# Authentication

## Certificate credentials

Async Sharepoint uses the Entra ID client-credentials flow with a certificate:

```python
from async_sharepoint import CertificateCredential, SharePointClient

credential = CertificateCredential(
    tenant_id="00000000-0000-0000-0000-000000000000",
    client_id="11111111-1111-1111-1111-111111111111",
    private_key_path="/path/to/azure-app-private.key",
    thumbprint="D33CFD3BB83E0EFB90AF709C897025286244FA0E",
)

async with SharePointClient("https://example.sharepoint.com/sites/Team", credential) as client:
    items = await client.ls()
```

The private key file may be a standalone PEM key or a combined certificate-and-key file; only the
`PRIVATE KEY` block is read. `thumbprint` is the certificate's hex-encoded SHA-1 digest, as shown
in the Azure portal. Malformed key and thumbprint values raise when the credential is constructed.

The OAuth scope is derived from the site host. For example,
`https://example.sharepoint.com/sites/Team` uses `https://example.sharepoint.com/.default`.

Entering the async context manager waits for the first token, so invalid credentials raise before
the first SharePoint request. A background task refreshes the token five minutes before expiry and
after a 401 or 403 response. Concurrent requests share that token and coalesce refresh attempts.

## Existing access tokens

Use `from_static_token()` when another component owns token acquisition:

```python
async with SharePointClient.from_static_token(site_url, access_token) as client:
    items = await client.ls()
```

A static token is never refreshed.

## Token management design

The rest of this page describes why the native refresher is shaped the way it is.

## Why this exists

The previous design took a Python callable (`get_token`) and invoked it on the request path via
`spawn_blocking` + `Python::attach`. That put the GIL in the middle of every HTTP request and was
the prime suspect for interactive Jupyter sessions hanging indefinitely after a period of idleness.

The replacement performs the Entra ID certificate client-credentials flow entirely in Rust. No
Python code runs while a request is in flight.

## The credential

`CertificateCredential` builds an RS256 **client assertion** — a JWT the application signs with its
own certificate private key and presents in place of a client secret.

- Header: `alg: RS256`, `typ: JWT`, `x5t` = base64url of the raw (hex-decoded) SHA-1 thumbprint.
- Claims: `aud` = the tenant's token endpoint, `iss` = `sub` = client ID, `jti` = a fresh UUID,
  `iat`/`nbf` = now, `exp` = now + 600s.
- POSTed form-encoded to `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token` with
  `grant_type=client_credentials` and
  `client_assertion_type=urn:ietf:params:oauth:client-assertion-type:jwt-bearer`.

The private key and thumbprint are parsed once in `CertCredential::load`, so malformed input raises
a `ValueError` at construction rather than inside the background task where nobody is listening.
`private_key_block` pulls just the `PRIVATE KEY` block out of the PEM, so a combined
certificate-and-key file works.

The scope is derived from the site host in `http.rs::default_scope`
(`https://contoso.sharepoint.com/sites/Team` becomes `https://contoso.sharepoint.com/.default`).

## Structure

Two channels, running in opposite directions. The field names don't help: `slot` is a **Receiver**,
`refresh` is a **Sender**.

| Channel | End in `Source::Refreshed` | End in the `refresher` task | Direction |
| --- | --- | --- | --- |
| `watch<TokenSlot>` | `slot` — Receiver | `sender` — Sender | refresher to request path |
| `mpsc<()>` capacity 1 | `refresh` — Sender | `refresh_rx` — Receiver | request path to refresher |

```
                      watch: token values
   refresher task  ──────────────────────────►  TokenManager::get
   (sender)                                     (slot)
        ▲
        └──────────────────────────────────────  TokenManager::get
                      mpsc: "refresh now"        (refresh.try_send)
                      (refresh_rx)
```

Every touch point, exhaustively:

- `slot` — cloned in `get`, then `borrow_and_update()` and `changed().await` in `wait_for_token`.
- `refresh` — `try_send(())` in `get`, only when `stale.is_some()`.
- `sender` — `send()` after every fetch attempt, `borrow()` on the failure path to read the token
  being carried forward, `closed()` in the `select!`.
- `refresh_rx` — `recv()` in the `select!`.

Callers of `get` are `ClientState::request()` (once per HTTP attempt) and `wait_ready()` from
`__aenter__`.

`ClientState::for_site` shares one `Arc<TokenManager>` across forked child clients. That is
deliberate — they are all the same tenant and scope, so they should share one refresher.

## `TokenSlot` and why both fields are `Option`

```rust
struct TokenSlot {
    generation: u64,
    token: Option<Arc<str>>,
    error: Option<String>,
}
```

All four combinations are reachable and each means something different:

| `token` | `error` | State | How you get here |
| --- | --- | --- | --- |
| `None` | `None` | Pending | `TokenSlot::default()` — the initial value `watch::channel` requires |
| `Some` | `None` | Healthy | A fetch succeeded |
| `None` | `Some` | Dead | The *first* fetch failed — bad cert, wrong tenant, no network |
| `Some` | `Some` | Degraded | A *refresh* failed, but the previous token is still valid |

**Pending is unavoidable.** `watch::channel` demands a value at construction, and construction
happens in `ClientState::new` (synchronous, from `__init__`) while the first fetch is inherently
async. Something has to occupy the channel in between.

**Degraded is the point of the 300s margin.** The refresher wakes five minutes before expiry. If a
transient network blip during that refresh wiped the token, the early refresh would buy nothing —
you would go straight from working to every-request-fails while still holding a token that is good
for another five minutes. Carrying it forward is what makes the margin useful, and it is why this
is not a `Result<Arc<str>, String>`: a `Result` cannot express "failed, but here is a still-valid
token".

## Why a generation counter

`request()` needs to say "this specific token got a 401, give me a different one". Comparing token
*strings* would look simpler, but **Entra returns the same cached token** for repeated
client-credentials requests within the token's lifetime. A value comparison would spin until the
60s timeout on any 401 that was not actually an expiry — a permissions change, say. The counter
lets the refresher say "this is genuinely a new response", so the request can fail fast with the
real error instead of hanging.

`generation += 1` happens at the *top* of the refresher loop, before the fetch, so a failure still
publishes a new generation. That is what wakes `get(Some(N))` waiters so they see the error rather
than blocking.

## Case walkthrough

**Startup.** `spawn` seeds the watch with `TokenSlot::default()` and the refresher fetches
immediately — the sleep is at the *bottom* of the loop. `__aenter__` calls `get(None)`, sees
generation 0 with no token, parks in `changed().await`. The refresher publishes generation 1 and the
waiter wakes. No mpsc traffic.

**Steady-state request.** `get(None)` then `borrow_and_update()` sees a token and returns. No await,
no channel traffic, just a read lock and an `Arc` clone. This is the hot path and it never touches
the mpsc.

**Proactive rotation.** Entirely refresher-internal. The `sleep` branch fires, the inner `while let`
recomputes `remaining` from the `SystemTime` deadline, and loops again if the `MAX_REFRESH_NAP` cap
cut the nap short. When `duration_since` errors the deadline has passed, the inner loop exits and
the outer loop refetches.

**401 on a live token.** `request()` sets `stale = Some(N)` and re-enters `get`. This is the only
thing that puts a message on the mpsc. The caller parks in `changed().await`; the refresher's
`recv()` branch fires, breaks the sleep loop, fetches, publishes generation N+1, and the waiter
wakes. A second 401 does not refresh again — `request()` only sets `stale` while it is `None`.

**Concurrent 401 stampede.** Twenty tasks holding generation N all call `try_send`. Capacity 1 means
one succeeds and nineteen get `Full`, which is discarded. All twenty park on their own cloned
receivers, the refresher does **one** fetch, and one `send` wakes all twenty. The coalescing is the
channel capacity doing the work, not any logic.

**Failed refresh.** `sender.borrow()` reads the outgoing token to carry forward (safe: single
sender, no interleaving). Callers passing `None` keep getting the old token; callers passing
`Some(N)` get the error. The backoff is `2^failures` capped at 60s plus jitter.

**Shutdown.** Dropping the last `ClientState` drops `TokenManager`, dropping both `slot` and
`refresh` at once. Three exit paths fire depending on where the refresher is parked: `send()`
returns `Err` if mid-fetch, `closed()` resolves if in the `select!`, `recv()` returns `None`
likewise. Redundant by design — they cover different await points. There is no abort handle and no
`Drop` impl; this is the entire shutdown mechanism.

**Static tokens.** `Source::Static` constructs neither channel and spawns no task. `get` returns
`(0, token)` synchronously. Used by `SharePointClient.from_static_token`, which is how the offline
mock-server tests authenticate.

## Subtleties that are load-bearing but invisible

**`slot.clone()` per call in `get`** is not defensive copying. `borrow_and_update` advances the
receiver's seen-version cursor, so concurrent callers sharing one receiver would steal each other's
change notifications.

**The bare block around the borrow in `wait_for_token`:**

```rust
{
    let current = slot.borrow_and_update();
    ...
}
if slot.changed().await.is_err() {
```

`watch::Ref` holds a read guard on the channel's internal `RwLock`. Holding it across
`changed().await` would block `sender.send()` — the refresher could never publish the value the
waiter is waiting for. That block is a deadlock guard.

**The `SystemTime` deadline with capped naps.** `Instant` is `CLOCK_MONOTONIC` on Linux and does
not advance while the machine is suspended. A single long `sleep()` would wake hours late still
believing it had time left. Recomputing `remaining` from a wall-clock deadline on every iteration
is what makes a closed laptop lid safe.

**`FORCED_REFRESH_DEBOUNCE`** covers a narrower case than the stampede above: requests that 401
while the refresher is *already mid-fetch*, so nothing is draining the mpsc. Their message sits in
the buffer, the in-flight fetch publishes a good token, waiters are satisfied — and then the
refresher enters the `select!` and finds a request that has already been answered.

Draining the channel after a successful publish would express this more precisely, and would not
also swallow a legitimate 401 on the *new* token arriving two seconds later:

```rust
while refresh_rx.try_recv().is_ok() {}
```

## Constants

| Constant | Value | Rationale |
| --- | --- | --- |
| `ASSERTION_LIFETIME` | 600s | JWT validity; Entra allows more but the assertion is single-use |
| `TOKEN_REFRESH_MARGIN` | 300s | Rotate this far ahead of expiry; also the window Degraded buys |
| `MIN_TOKEN_LIFETIME` | 60s | Floor, in case Entra returns a very short `expires_in` |
| `MAX_REFRESH_NAP` | 600s | Re-check the wall clock at least this often (suspend guard) |
| `TOKEN_WAIT_TIMEOUT` | 60s | Upper bound on `get`; prevents an unbounded hang |
| `MAX_BACKOFF` | 60s | Retry ceiling after repeated fetch failures |

## Testing

`SharePointClient.from_static_token` is the offline path — see `tests/test_api.py` and
`tests/test_native_http.py`, which run against a local mock HTTP server with no network.
`tests/test_comm.py` exercises the real certificate flow against live SharePoint.

Unit tests in `auth.rs` must avoid pyo3 types. The crate is `cdylib` + `extension-module`, so the
`cargo test` binary has no libpython; it links only because the linker dead-strips all
pyo3-touching code. A test that calls anything returning `PyResult` drags the whole pyo3 runtime
back in and breaks the link with hundreds of undefined `Py*` symbols. This is why
`http.rs::default_scope` returns `Option<String>` and the caller maps it to a `PyErr`.

Note that `uv run` re-syncs the project (`[tool.uv] package = true`) and will silently overwrite
the `.so` that `maturin develop` just installed. Use `uv run --no-sync` for both steps when testing
local Rust changes, or you will be exercising a stale binary.