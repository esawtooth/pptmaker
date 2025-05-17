import argparse
import sys
from .converter import pptx_to_yaml


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert PPTX to YAML")
    parser.add_argument("pptx", help="Path to PPTX file")
    parser.add_argument("output", help="Path to output YAML file")
    args = parser.parse_args(argv)

    yaml_text = pptx_to_yaml(args.pptx)
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(yaml_text)
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
