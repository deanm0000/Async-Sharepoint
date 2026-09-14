#!/bin/bash
uv build && uv pip install --python /tmp/async-sharepoint-py3.14/bin/python --reinstall dist/*.whl