#!/usr/bin/env python3
# coding=utf8

import argparse
import logging as log
import os

# external dependencies
import docx
import OleFileIO_PL


parser = argparse.ArgumentParser(
    formatter_class=argparse.RawDescriptionHelpFormatter)

parser.add_argument('--input', '-i', action='store',
                    help='Path to the directory that contains document files that you was restored')
parser.add_argument('--output', '-o', action='store',
                    help='Specified path to save.')
parser.add_argument('--verbose', '-v', action='store_true',
                    help='Verbose output')

args = parser.parse_args()

if args.verbose:
    log.basicConfig(format="%(levelname)s: %(message)s", level=log.INFO)
    log.info('Verbose output.')
else:
    log.basicConfig(format="%(levelname)s: %(message)s")

input_path = os.path.abspath(
    os.path.curdir) if not args.input else args.input
output_path = os.path.abspath(
    os.path.curdir) if not args.output else args.output
