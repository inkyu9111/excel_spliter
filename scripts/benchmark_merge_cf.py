"""Measure only source conditional-format eligibility inspection; no Excel needed.

python scripts/benchmark_merge_cf.py [--baseline path/to/merge_conditional_formats.py]
Uses disposable synthetic ZIPs; excludes merging, copying, and output rewriting.
"""
import argparse
import gc
from pathlib import Path
import runpy
from statistics import median
import sys
from tempfile import TemporaryDirectory
from time import perf_counter
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from excel_splitter import merge_conditional_formats as cf


def inspect(module, path):
    package = None
    try:
        package = module["_read"](path)
        return module["_original"](package)[1]
    except ValueError:
        return None
    finally:
        if package is not None:
            package.sheet.unlink()
            package.styles.unlink()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", type=Path, help="Trusted saved module to execute for comparison")
    args = parser.parse_args()
    modules = {"current": vars(cf)}
    if args.baseline:
        modules = {"baseline": runpy.run_path(str(args.baseline)), **modules}
    data = ''.join(f'<row r="{row}">' + ''.join(
        f'<c r="{column}{row}"><v>{row}</v></c>' for column in "ABCDEFGH"
    ) + '</row>' for row in range(1, 20001))
    rule = '<conditionalFormatting sqref="H1:H1048576"><cfRule type="expression" priority="1"><formula>$H1&gt;0</formula></cfRule></conditionalFormatting>'
    print("20,000 x 8 synthetic cells; times cover CF eligibility inspection only, not Excel or Merge.", flush=True)
    print(f"Current: {cf.__file__}\nBaseline: {args.baseline or '(none)'}", flush=True)
    scratch = Path(__file__).resolve().parents[1] / "build"
    scratch.mkdir(exist_ok=True)
    with TemporaryDirectory(prefix="merge-cf-benchmark-", dir=scratch) as directory:
        for has_rule in (False, True):
            path = Path(directory) / f"synthetic-{has_rule}.xlsx"
            with ZipFile(path, "w", ZIP_DEFLATED) as archive:
                archive.writestr("xl/workbook.xml", f'<workbook xmlns="{cf._NS}"/>')
                archive.writestr("xl/tables/table1.xml", f'<table xmlns="{cf._NS}" ref="A1:H20000"/>')
                archive.writestr("xl/styles.xml", f'<styleSheet xmlns="{cf._NS}"/>')
                archive.writestr("xl/theme/theme1.xml", b"synthetic theme")
                archive.writestr("xl/worksheets/sheet1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    + f'<worksheet xmlns="{cf._NS}"><sheetData>{data}</sheetData>{rule if has_rule else ""}</worksheet>')
            before = path.read_bytes()
            expected = (rule, None, ("A", "1", "H", (("xl/theme/theme1.xml", b"synthetic theme"),), ())) if has_rule else None
            samples, reads = {name: [] for name in modules}, {}
            for name, module in modules.items():
                original_read = ZipFile.read
                members = []
                def counted(archive, member, *args, **kwargs):
                    members.append(member)
                    return original_read(archive, member, *args, **kwargs)
                with patch.object(ZipFile, "read", counted):
                    assert inspect(module, path) == expected, (name, has_rule)
                reads[name] = len(members)
            for round_index in range(5):
                order = list(modules.items())
                for name, module in (order if round_index % 2 == 0 else reversed(order)):
                    gc.collect()
                    started = perf_counter()
                    assert inspect(module, path) == expected
                    samples[name].append(perf_counter() - started)
            assert path.read_bytes() == before, "Synthetic input changed"
            for name in modules:
                print(f"CF={has_rule} {name}: median={median(samples[name]):.4f}s, ZIP-member reads={reads[name]}, samples={samples[name]}", flush=True)
    print("PASS: fixed expected rule/no-rule decisions, unchanged inputs, temporary cleanup", flush=True)


if __name__ == "__main__":
    main()
