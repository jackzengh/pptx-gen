import json
import sys

from harness import check_pptx


def main():
    if len(sys.argv) != 2:
        print("usage: python main.py <deck.pptx>", file=sys.stderr)
        raise SystemExit(2)
    result = check_pptx(sys.argv[1])
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
