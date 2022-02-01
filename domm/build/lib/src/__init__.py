"""
entry point for domain matching script
"""
from src.domm import run
import argparse

def main():
    """
    Parse command line arguemtns and start
    """
    parser = argparse.ArgumentParser(description="Match RCRAInfo handler by email domains. Without output flag specified will print stats to stdout")
    parser.add_argument('path',
                        help='path to the file')
    parser.add_argument('--sheet',
                        help='Excel sheet name, or 0-indexed integer, else will read all')
    parser.add_argument('--output', '-o',
                        help='Path of output to write to (.xlsx)',
                        action='store')
    parser.add_argument('--display', '-d',
                        help='Display parsed info',
                        choices=['contacts', 'stats' ],
                        action='store')
    args = parser.parse_args()
    # see run() in domm.py
    run(args)

if __name__ == '__main__':
    main()