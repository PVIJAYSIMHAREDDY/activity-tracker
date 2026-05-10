#!/bin/bash
# Launch the Daily Activity Tracker desktop app
cd "$(dirname "$(realpath "$0")")"
exec python3 main.py "$@"
