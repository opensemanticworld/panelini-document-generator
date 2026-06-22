"""Script for debugging Panel app in PyCharm or VSCode

Instructions:
- Set breakpoints in main.py or any imported module.
- Run this script in your IDE's debugger (set script path to run_debug.py).
- The Panel app will be served, and breakpoints will be hit as expected.

If INTERNAL_PORT is not set in your environment, defaults to 8000 for convenience.
"""

import os
from dotenv import load_dotenv
import panel as pn

load_dotenv(verbose=True)

# Import the Panel app object from main.py
from document_generator.main import app

if __name__ == "__main__":
    port = int(os.getenv("INTERNAL_PORT", "8000"))
    pn.serve(app, port=port, show=True)
