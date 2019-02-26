# !/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Michael McKenzie"

import zipfile
import argparse
import sys
import os


def create_parser():
    parser = argparse.ArgumentParser(
        description='Searches .dotm file for text')
    parser.add_argument("--dir", help="directory to search", default=".")
    parser.add_argument("text", help="text to search for")
    return parser


def main(directory, text_search):
    pathway = os.listdir(directory)
    searched_files = 0
    matched_files = 0
    for file in pathway:
        searched_files += 1
        all_paths = os.path.join(directory, file)
        if not all_paths.endswith(".dotm"):
            print("this is not a dotm file {}".format(all_paths))
            continue
        if not zipfile.is_zipfile(all_paths):
            print("this is not a zip file {}".format(all_paths))
            continue
        with zipfile.ZipFile(all_paths, "r") as zipped:
            archive = zipped.namelist()
            if "word/document.xml" in archive:
                with zipped.open("word/document.xml", "r") as doc:
                    for line in doc:
                        i = line.find(text_search)
                        if i >= 0:
                            matched_files += 1
                            print(line[i - 40:i + 40])
    print("Searching dotm files in: {}".format(directory))
    print("I will look for magic text: {}".format(text_search))
    print("Total Number of matched files: {}".format(matched_files))
    print("Total Number of searched files: {}".format(searched_files))


if __name__ == '__main__':
    parser = create_parser()
    ns = parser.parse_args()
    main(ns.dir, ns.text)
