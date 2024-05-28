import argparse
from workers.converter import convert_word_to_pdf

def main():
    parser = argparse.ArgumentParser(description="Convert Word files to PDF and optionally print them.")
    parser.add_argument('files', metavar='F', type=str, nargs='+', help='Word files to convert')
    parser.add_argument('--print', dest='print_after_conversion', action='store_true', help='Print files after conversion')
    parser.add_argument('--formatted', dest='formatted', action='store_true', help='Apply additional formatting to the PDF')

    args = parser.parse_args()

    convert_word_to_pdf(args.files, args.print_after_conversion, args.formatted)

if __name__ == "__main__":
    main()
