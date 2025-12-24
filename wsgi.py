#!/usr/bin/env python3
"""
WSGI entry point for xserver deployment.
"""

import sys
import os

# Add project directory to path
sys.path.insert(0, os.path.dirname(__file__))

# Import Flask app
from app import app as application

if __name__ == "__main__":
    application.run()

