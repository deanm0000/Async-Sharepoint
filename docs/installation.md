# Installation

## Stable release

To install Async Sharepoint, run this command in your terminal:

```sh
uv add Async-Sharepoint
```

Or if you prefer to use `pip`:

```sh
pip install Async-Sharepoint
```

## From source

The source files for Async Sharepoint can be downloaded from the [Github repo](https://github.com/deanm0000/Async-Sharepoint).

You can either clone the public repository:

```sh
git clone https://github.com/deanm0000/Async-Sharepoint
```

Or download the [tarball](https://github.com/deanm0000/Async-Sharepoint/tarball/main):

```sh
curl -OJL https://github.com/deanm0000/Async-Sharepoint/tarball/main
```

Once you have a copy of the source, you can install it with:

```sh
cd Async-Sharepoint
uv sync
uv run maturin develop --uv
```

Building from source requires Rust 1.88 or newer. Published packages currently target Linux.
