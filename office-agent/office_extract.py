#!/usr/bin/env python3
"""Extract text from MVP Office-agent HTML documents."""

import argparse
import json
from html.parser import HTMLParser


class TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []
    def handle_data(self, data):
        text = data.strip()
        if text:
            self.parts.append(text)


def main():
    parser = argparse.ArgumentParser(description='Extract text outline from generated Office files.')
    parser.add_argument('input')
    parser.add_argument('output')
    args = parser.parse_args()
    parser_obj = TextExtractor()
    with open(args.input, 'r', encoding='utf-8', errors='ignore') as f:
        parser_obj.feed(f.read())
    result = {'source': args.input, 'text': parser_obj.parts}
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(args.output)


if __name__ == '__main__':
    main()
