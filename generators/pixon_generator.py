# Deprecated — use projects/pixon.py
# This shim exists for backward compatibility only.
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from projects.pixon import PixonGenerator  # noqa: F401

if __name__ == "__main__":
    PixonGenerator.main()
