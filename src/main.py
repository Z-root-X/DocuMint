import sys
import os

# Add the parent directory to sys.path when running directly
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from documint.gui import main

if __name__ == "__main__":
    main()
