"""Excelブックのシート構成とセルの値・数式を比較する。"""

import argparse
from contextlib import ExitStack, closing
from difflib import SequenceMatcher
from pathlib import Path
import sys
from zipfile import BadZipFile

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

from excel_report import ReportWriteError, write_report


EXAMPLES_DIR = Path(__file__).resolve().parent / "examples"


class WorkbookReadError(Exception):
    """入力ブックを読み込めない場合の利用者向けエラー。"""


def load_workbook(path):
    """数式を文字列のまま読み込み、入力エラーにファイル名を添える。"""
    if not path.is_file():
        raise WorkbookReadError(f"ファイルが見つかりません: {path}")

    try:
        return openpyxl.load_workbook(path, data_only=False)
    except (
        OSError, ValueError, KeyError, IndexError, TypeError,
        SyntaxError, EOFError, BadZipFile, InvalidFileException,
    ) as exc:
        raise WorkbookReadError(
            f"Excelファイルを読み込めません: {path} ({exc})"
        ) from exc


# ============================================================
# 共通処理
# ============================================================

def display_value(value):
    """Noneを見やすく表示する"""
    if value is None:
        return "空欄"

    return value


def get_rows(ws, max_col):
    """
    ワークシートを行単位で取得する。

    例:
    [
        ("項目1", "説明1", 100),
        ("項目2", "説明2", 200),
        ...
    ]
    """
    rows = []

    for row in range(1, ws.max_row + 1):
        row_values = tuple(
            ws.cell(row, col).value
            for col in range(1, max_col + 1)
        )

        rows.append(row_values)

    return rows


# ============================================================
# シートの追加・削除・移動
# ============================================================

def detect_sheet_changes(old_wb, new_wb):
    old_names = old_wb.sheetnames
    new_names = new_wb.sheetnames

    old_set = set(old_names)
    new_set = set(new_names)

    changes = []

    # --------------------------------------------------------
    # シート追加
    # --------------------------------------------------------

    for index, sheet_name in enumerate(new_names):

        if sheet_name not in old_set:
            changes.append({
                "type": "シート追加",
                "sheet": sheet_name,
                "new_position": index + 1
            })

    # --------------------------------------------------------
    # シート削除
    # --------------------------------------------------------

    for index, sheet_name in enumerate(old_names):

        if sheet_name not in new_set:
            changes.append({
                "type": "シート削除",
                "sheet": sheet_name,
                "old_position": index + 1
            })

    # --------------------------------------------------------
    # 両方に存在するシートだけ抽出
    #
    # 例:
    #
    # 旧
    # A B C
    #
    # 新
    # A X B C
    #
    # ↓
    #
    # common_old = A B C
    # common_new = A B C
    #
    # なのでB,Cは移動扱いにならない
    # --------------------------------------------------------

    common_old = [
        name
        for name in old_names
        if name in new_set
    ]

    common_new = [
        name
        for name in new_names
        if name in old_set
    ]

    # --------------------------------------------------------
    # SequenceMatcherの一致ブロックを利用して
    # 「順序を維持しているシート」を探す
    # --------------------------------------------------------

    matcher = SequenceMatcher(
        None,
        common_old,
        common_new,
        autojunk=False
    )

    stable_sheets = set()

    for block in matcher.get_matching_blocks():

        for i in range(block.size):
            sheet_name = common_old[block.a + i]

            stable_sheets.add(sheet_name)

    # --------------------------------------------------------
    # 共通シートなのに一致ブロックから外れたもの
    # = 実際に移動した可能性が高いシート
    # --------------------------------------------------------

    common_set = old_set & new_set

    for sheet_name in old_names:

        if sheet_name not in common_set:
            continue

        if sheet_name in stable_sheets:
            continue

        old_position = old_names.index(sheet_name) + 1
        new_position = new_names.index(sheet_name) + 1

        changes.append({
            "type": "シート移動",
            "sheet": sheet_name,
            "old_position": old_position,
            "new_position": new_position
        })

    return changes


# ============================================================
# シート内部の比較
# ============================================================

def compare_sheet(sheet_name, old_ws, new_ws):
    changes = []

    max_col = max(
        old_ws.max_column,
        new_ws.max_column
    )

    old_rows = get_rows(old_ws, max_col)
    new_rows = get_rows(new_ws, max_col)

    # --------------------------------------------------------
    # 行単位で比較
    #
    # これによって途中に行が追加されても
    # 後続行を全部変更扱いしない
    # --------------------------------------------------------

    matcher = SequenceMatcher(
        None,
        old_rows,
        new_rows,
        autojunk=False
    )

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():

        # ----------------------------------------------------
        # 同一部分
        # ----------------------------------------------------

        if tag == "equal":
            continue

        # ----------------------------------------------------
        # 行追加
        # ----------------------------------------------------

        if tag == "insert":

            for new_index in range(j1, j2):

                new_row_number = new_index + 1

                for col in range(1, max_col + 1):

                    new_value = new_rows[new_index][col - 1]

                    # 空欄は出力しない
                    if new_value is None:
                        continue

                    cell = new_ws.cell(
                        new_row_number,
                        col
                    ).coordinate

                    changes.append({
                        "type": "追加",
                        "sheet": sheet_name,
                        "cell": cell,
                        "old": None,
                        "new": new_value
                    })

        # ----------------------------------------------------
        # 行削除
        # ----------------------------------------------------

        elif tag == "delete":

            for old_index in range(i1, i2):

                old_row_number = old_index + 1

                for col in range(1, max_col + 1):

                    old_value = old_rows[old_index][col - 1]

                    if old_value is None:
                        continue

                    cell = old_ws.cell(
                        old_row_number,
                        col
                    ).coordinate

                    changes.append({
                        "type": "削除",
                        "sheet": sheet_name,
                        "cell": cell,
                        "old": old_value,
                        "new": None
                    })

        # ----------------------------------------------------
        # 行の中身が変更された
        # ----------------------------------------------------

        elif tag == "replace":

            old_count = i2 - i1
            new_count = j2 - j1

            pair_count = min(
                old_count,
                new_count
            )

            # ------------------------------------------------
            # 対応する行同士をセル単位で比較
            # ------------------------------------------------

            for offset in range(pair_count):

                old_index = i1 + offset
                new_index = j1 + offset

                for col in range(1, max_col + 1):

                    old_value = old_rows[old_index][col - 1]
                    new_value = new_rows[new_index][col - 1]

                    if old_value == new_value:
                        continue

                    cell = new_ws.cell(
                        new_index + 1,
                        col
                    ).coordinate

                    changes.append({
                        "type": "変更",
                        "sheet": sheet_name,
                        "cell": cell,
                        "old": old_value,
                        "new": new_value
                    })

            # ------------------------------------------------
            # 旧側の方が行数が多い
            # → 余った行は削除
            # ------------------------------------------------

            if old_count > new_count:

                for old_index in range(
                    i1 + pair_count,
                    i2
                ):

                    for col in range(1, max_col + 1):

                        old_value = old_rows[old_index][col - 1]

                        if old_value is None:
                            continue

                        cell = old_ws.cell(
                            old_index + 1,
                            col
                        ).coordinate

                        changes.append({
                            "type": "削除",
                            "sheet": sheet_name,
                            "cell": cell,
                            "old": old_value,
                            "new": None
                        })

            # ------------------------------------------------
            # 新側の方が行数が多い
            # → 余った行は追加
            # ------------------------------------------------

            elif new_count > old_count:

                for new_index in range(
                    j1 + pair_count,
                    j2
                ):

                    for col in range(1, max_col + 1):

                        new_value = new_rows[new_index][col - 1]

                        if new_value is None:
                            continue

                        cell = new_ws.cell(
                            new_index + 1,
                            col
                        ).coordinate

                        changes.append({
                            "type": "追加",
                            "sheet": sheet_name,
                            "cell": cell,
                            "old": None,
                            "new": new_value
                        })

    return changes


# ============================================================
# 結果表示
# ============================================================

def print_results(sheet_changes, cell_changes):

    if not sheet_changes and not cell_changes:
        print("変更なし")
        return

    print("変更あり")
    print()

    # --------------------------------------------------------
    # シート変更
    # --------------------------------------------------------

    for change in sheet_changes:

        if change["type"] == "シート追加":

            print("種類: シート追加")
            print(f"シート名: {change['sheet']}")
            print(f"場所: {change['new_position']}番目")
            print()

        elif change["type"] == "シート削除":

            print("種類: シート削除")
            print(f"シート名: {change['sheet']}")
            print(f"変更前: {change['old_position']}番目")
            print()

        elif change["type"] == "シート移動":

            print("種類: シート移動")
            print(f"シート名: {change['sheet']}")
            if change["old_position"] == change["new_position"]:
                print(
                    f"場所: {change['new_position']}番目"
                    "（共通シート間の順序変更）"
                )
            else:
                print(f"変更前: {change['old_position']}番目")
                print(f"変更後: {change['new_position']}番目")
            print()

    # --------------------------------------------------------
    # セル変更
    # --------------------------------------------------------

    for change in cell_changes:

        print(f"種類: {change['type']}")
        print(f"シート名: {change['sheet']}")
        print(f"場所: {change['cell']}")
        print(
            f"変更前: "
            f"{display_value(change['old'])}"
        )
        print(
            f"変更後: "
            f"{display_value(change['new'])}"
        )
        print()


# ============================================================
# メイン処理
# ============================================================

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Excelファイルのシート構成とセルの値・数式を比較します。",
        epilog="引数を省略すると、同梱のexamples/old.xlsxとexamples/new.xlsxを比較します。",
    )
    parser.add_argument("old_file", nargs="?", type=Path, help="変更前のExcelファイル")
    parser.add_argument("new_file", nargs="?", type=Path, help="変更後のExcelファイル")
    parser.add_argument(
        "-o", "--output", type=Path, metavar="PATH",
        help="Excelレポートの保存先（省略時: reports/変更点まとめ_日時.xlsx）",
    )
    args = parser.parse_args(argv)

    if (args.old_file is None) != (args.new_file is None):
        parser.error("変更前・変更後のファイルを両方指定してください。")

    old_file = args.old_file or EXAMPLES_DIR / "old.xlsx"
    new_file = args.new_file or EXAMPLES_DIR / "new.xlsx"

    # 後のファイルの読み込みや比較に失敗した場合も、読み込んだブックを閉じる。
    with ExitStack() as stack:
        try:
            old_wb = stack.enter_context(closing(load_workbook(old_file)))
            new_wb = stack.enter_context(closing(load_workbook(new_file)))
        except WorkbookReadError as exc:
            print(f"エラー: {exc}", file=sys.stderr)
            return 1

        sheet_changes = detect_sheet_changes(old_wb, new_wb)
        cell_changes = []
        new_sheet_set = set(new_wb.sheetnames)

        # 両方に存在するシートだけ比較する。
        for sheet_name in old_wb.sheetnames:
            if sheet_name not in new_sheet_set:
                continue

            cell_changes.extend(compare_sheet(
                sheet_name,
                old_wb[sheet_name],
                new_wb[sheet_name],
            ))

        print_results(sheet_changes, cell_changes)

    try:
        report_path = write_report(
            sheet_changes, cell_changes, old_file, new_file,
            output_path=args.output,
        )
    except ReportWriteError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1

    print(f"レポート保存先: {report_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
