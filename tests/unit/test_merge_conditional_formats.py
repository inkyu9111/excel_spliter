from xml.dom import minidom
from zipfile import ZipFile

import pytest


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RULE = '<cfRule type="expression" dxfId="0" priority="1" stopIfTrue="1"><formula>COUNTIF($H:$H,H1)&gt;1</formula></cfRule>'
STYLE = '<dxf><font><b/><color rgb="FFFF0000"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill><border><bottom style="thin"/></border><numFmt numFmtId="165" formatCode="0.00"/></dxf>'


def _package(path, *, rule=RULE, scope="H1:H1048576", table="B5:H7", style=STYLE,
             extra="", names="", theme=b"same theme"):
    with ZipFile(path, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", f'<worksheet xmlns="{NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xr="urn:test-revision" mc:Ignorable="xr"><sheetData><row r="6"><c r="H6"><v>7</v></c></row></sheetData><conditionalFormatting sqref="{scope}">{rule}</conditionalFormatting>{extra}<tableParts count="1"/></worksheet>')
        archive.writestr("xl/styles.xml", f'<styleSheet xmlns="{NS}"><dxfs count="1">{style}</dxfs><colors><indexedColors><rgbColor rgb="FF000000"/></indexedColors></colors></styleSheet>')
        archive.writestr("xl/tables/table1.xml", f'<table xmlns="{NS}" ref="{table}"/>')
        archive.writestr("xl/workbook.xml", f'<workbook xmlns="{NS}">{names}</workbook>')
        archive.writestr("xl/theme/theme1.xml", theme)
        archive.writestr("untouched.bin", b"values, comments, links, and other package parts")
    return path


def _parts(path):
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


@pytest.mark.parametrize("formula", ["COUNTIF($H:$H,H1)&gt;1", "AND($H1&gt;0,$C1=2)", "$H$2&gt;0", "SUM(H:H)&gt;10"])
def test_restores_one_original_full_column_rule_after_fifteen_row_block_pastes(tmp_path, formula):
    from excel_splitter import merge_conditional_formats as cf

    rule = RULE.replace("COUNTIF($H:$H,H1)&gt;1", formula)
    sources = tuple(_package(tmp_path / f"part{i}.xlsx", rule=rule) for i in range(15))
    output = _package(tmp_path / "merged.xlsx", rule=rule, scope="H6:H7",
                      extra=''.join(f'<conditionalFormatting sqref="H{8+i*2}:H{9+i*2}">{rule}</conditionalFormatting>' for i in range(14)))
    originals, before = tuple(_parts(path) for path in sources), _parts(output)

    assert cf.restore_full_column_rule(sources, output)

    after = _parts(output)
    sheet = minidom.parseString(after["xl/worksheets/sheet1.xml"])
    rules = sheet.getElementsByTagNameNS(NS, "conditionalFormatting")
    assert len(rules) == 1 and rules[0].getAttribute("sqref") == "H1:H1048576"
    assert rules[0].getElementsByTagNameNS(NS, "cfRule")[0].toxml() == rule
    assert sheet.documentElement.getAttribute("xmlns:xr") == "urn:test-revision"
    assert sheet.documentElement.getAttribute("mc:Ignorable") == "xr"
    assert sheet.getElementsByTagNameNS(NS, "v")[0].firstChild.data == "7"
    assert {k: v for k, v in after.items() if k != "xl/worksheets/sheet1.xml"} == {k: v for k, v in before.items() if k != "xl/worksheets/sheet1.xml"}
    assert tuple(_parts(path) for path in sources) == originals


@pytest.mark.parametrize("different", [
    {"scope": "H6:H7"}, {"scope": "H1:H1048576 G1:G1048576"},
    {"rule": RULE.replace("stopIfTrue=\"1\"", "stopIfTrue=\"0\"")},
    {"rule": RULE.replace("priority=\"1\"", "priority=\"2\"")},
    {"rule": RULE.replace("&gt;1", "&gt;2")},
    {"style": STYLE.replace("FFFF0000", "FF00FF00")}, {"theme": b"different theme"},
    {"table": "C5:I7"}, {"table": "B6:H8"},
    {"extra": f'<conditionalFormatting sqref="C1:C1048576">{RULE}</conditionalFormatting>'},
    {"extra": '<extLst><ext uri="test"><x:conditionalFormatting xmlns:x="urn:extended"/></ext></extLst>'},
    {"names": '<definedNames><definedName name="Threshold">2</definedName></definedNames>'},
])
def test_leaves_block_rules_untouched_when_source_rules_or_context_differ(tmp_path, different):
    from excel_splitter import merge_conditional_formats as cf

    sources = (_package(tmp_path / "a.xlsx"), _package(tmp_path / "b.xlsx", **different))
    output = _package(tmp_path / "merged.xlsx", scope="H6:H7")
    before = output.read_bytes()
    assert not cf.restore_full_column_rule(sources, output)
    assert output.read_bytes() == before


@pytest.mark.parametrize("case", ["encrypted", "dxf_changed", "external_formula", "extended_output"])
def test_unproven_packages_are_never_rewritten(tmp_path, case):
    from excel_splitter import merge_conditional_formats as cf

    rule = RULE.replace("COUNTIF($H:$H,H1)", "'[other.xlsx]Data'!H1") if case == "external_formula" else RULE
    sources = (_package(tmp_path / "a.xlsx", rule=rule), _package(tmp_path / "b.xlsx", rule=rule))
    if case == "encrypted":
        sources[1].write_bytes(b"DRM: Excel-only input")
    output = _package(tmp_path / "merged.xlsx", scope="H6:H7", style=STYLE.replace("0.00", "0.000") if case == "dxf_changed" else STYLE,
                      extra='<extLst><ext uri="test"><x:conditionalFormatting xmlns:x="urn:extended"/></ext></extLst>' if case == "extended_output" else "")
    before = output.read_bytes()
    assert not cf.restore_full_column_rule(sources, output)
    assert output.read_bytes() == before


def test_rewrite_failure_preserves_saved_output_and_cleans_its_temporary_file(tmp_path, monkeypatch):
    from excel_splitter import merge_conditional_formats as cf

    sources = (_package(tmp_path / "a.xlsx"), _package(tmp_path / "b.xlsx"))
    output = _package(tmp_path / "merged.xlsx", scope="H6:H7")
    before = {path: path.read_bytes() for path in (*sources, output)}

    def fail_replace(*_args):
        raise OSError("replace failed")

    monkeypatch.setattr(cf.os, "replace", fail_replace)
    with pytest.raises(OSError, match="replace failed"):
        cf.restore_full_column_rule(sources, output)
    assert {path: path.read_bytes() for path in tmp_path.iterdir()} == before


def test_unused_builtin_layout_names_do_not_prevent_restoration(tmp_path):
    from excel_splitter import merge_conditional_formats as cf

    names = '<definedNames><definedName name="_xlnm.Print_Area">Data!$A$1:$H$7</definedName></definedNames>'
    sources = (_package(tmp_path / "a.xlsx", names=names), _package(tmp_path / "b.xlsx", names=names))
    output = _package(tmp_path / "merged.xlsx", names=names, scope="H6:H7")
    assert cf.restore_full_column_rule(sources, output)

    rule = RULE.replace("COUNTIF($H:$H,H1)&gt;1", "SUM(Print_Area)&gt;1")
    for path in sources:
        _package(path, names=names, rule=rule)
    before = output.read_bytes()
    assert not cf.restore_full_column_rule(sources, output)
    assert output.read_bytes() == before


def test_external_formula_in_color_scale_threshold_is_not_restored(tmp_path):
    from excel_splitter import merge_conditional_formats as cf

    rule = '<cfRule type="colorScale" priority="1"><colorScale><cfvo type="formula" val="Other!A1"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule>'
    sources = (_package(tmp_path / "a.xlsx", rule=rule), _package(tmp_path / "b.xlsx", rule=rule))
    output = _package(tmp_path / "merged.xlsx", rule=rule, scope="H6:H7")
    before = output.read_bytes()
    assert not cf.restore_full_column_rule(sources, output)
    assert output.read_bytes() == before
