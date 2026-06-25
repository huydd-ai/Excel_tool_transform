import os
import sys

# Allow `python -m excel_tool` from a parent directory: the modules in this
# package use flat top-level imports (from excel_to_airtest import ...), so put
# this directory on sys.path before importing them.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_to_airtest import AirtestGenerator

AirtestGenerator.main()
