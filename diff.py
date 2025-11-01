# =============================================================================
# ファイル差分比較ツール
# =============================================================================
# [変更履歴]
# Ver1.00  2023/11/04 新規作成
# Ver1.01  2023/12/23 第2引数がpathの場合、第1引数のファイル名補完追加
# Ver1.02  2023/12/29 html形式の出力機能追加
# Ver1.03  2024/01/04 Excelへの出力機能追加
# Ver1.04  2024/01/06 左右に並べて出力での参照行数の誤り修正

# import
import os
import sys
import chardet
import locale
import re
import math
import shutil
import unicodedata
import argparse
from pathlib import Path
from unicodedata import east_asian_width

if os.name == "nt":
	# ESCシーケンス有効化用
	from ctypes import windll, wintypes, byref

# 扱える文字コード
USABLE_CODE = ("ascii", "SHIFT_JIS","utf-8","UTF-8-SIG", "UTF-16","EUC-JP")

# ESCシーケンス
RESET  = "\033[0m"
REVERSE= "\033[7m"
BLACK  = "\033[30m"
RED    = "\033[31m"
GREEN  = "\033[32m"
YELLOW = "\033[33m"
BLUE   = "\033[34m"
MAGENTA= "\033[35m"
CYAN   = "\033[36m"
WHITE  = "\033[37m"
B_YELLOW = "\033[30;43m"
B_CYAN   = "\033[30;46m"

HEAD_HTML = """
<!DOCTYPE html>
<html>
<head>
<title>diff.py compare result</title>
<style>
  table {
    table-layout: fixed;
    margin: 0;
    border: 1px solid #a0a0a0;
    box-shadow: 1px 1px 2px rgba(0, 0, 0, 0.15);
  }
  th {
    position: sticky;
    top: 0;
  }
  td,th {
    word-break: break-all;
    font-size: 12pt;
    padding: 0 3px;
  }
  tr {
    vertical-align: top;
  }
  .title {
    font-weight: bold;
    color: white;
    background-color: blue;
    vertical-align: top;
    text-align: center;
    padding: 4px 4px;
    background: linear-gradient(mediumblue, darkblue);
  }
  .line {color: #0000ff; background-color: #ffffff; text-align: right;}
  .mark {color: #000f00; background-color: #ffffff; text-align: center;}
  .same {color: #000000; background-color: #ffffff;}
  .diff {color: #000000; background-color: #efcb05;}
  .null {color: #000000; background-color: #00ccff;}

</style>
</head>

"""

THREAD_HTML = """
<body>
  <table style="width: 100%%; border-collapse: collapse;">
    <colgroup>
      <col style="width: {num_width}em;" />
      <col style="width: calc(100% / 2 - 0.5em - {num_width}em);" />
      <col style="width: 0,5em;" />
      <col style="width: {num_width}em;" />
      <col style="width: calc(100% / 2 - 0.5em - {num_width}em);" />
    </colgroup>

    <thead>
      <tr>
        <th colspan="2" class="title">{file1}</th>
        <th></th>
        <th colspan="2" class="title">{file2}</th>
      </tr>
    </thead>
"""

TAIL_HTML = """
    </tbody>
  </table>
</body>
</html>
"""


class FileCtrl():
	"""
	入力ファイルに関する制御

	Attributes:
		filename  : 入力ファイル名
		lines[]   : 行データ
		count     : 行数
		max_width : 1行の最大幅
		num_width : 行数の最大幅
	"""

	def __init__(self, path_file, org_file=""):
		path = Path(path_file)
		if org_file == "":
			# ファイル存在チェック
			if path.is_file():
				file = path_file
			else:
				print(f"Not exist file({path_file})")
				sys.exit(1)
		else:
			# ファイル存在チェック
			if path.is_file():
				file = path_file
			# DIR存在チェック
			elif path.is_dir():
				# ファイル名を補完
				org_name = Path(org_file).name
				file = Path(path_file, org_file)
				if not file.is_file():
					print(f"Not exist file({file})")
					sys.exit(1)
			else:
				print(f"Not exist file or path({path_file})")
				sys.exit(1)

		# ファイル読込み
		with open(file, "rb") as f:
			txt = f.read()

		# 文字コード変換+行データとして格納
		guess = chardet.detect(txt).get("encoding")
		if not(guess in USABLE_CODE):
			guess = locale.getpreferredencoding()
		self.lines = txt.decode(guess).splitlines()
		self.count = len(self.lines)
		self.filename = file

		self.max_width = 0
		for row in self.lines:
			width = self.unicode_len(row)
			if self.max_width < width:
				self.max_width = width

		num = f"{self.count}"
		self.num_width = len(num)


	def unicode_len(self, text):
		"""
		漢字コードを含む文字列の長さ取得

		Args:
			text (str): 文字列

		Returns:
			int: 画面表示上の幅(全角:2, 半角:1)
		"""
		width = 0
		for char in text:
			if re.fullmatch(r"[┌ ┬ ┐ ├ ┼ ┤ └ ┴ ┘ │ ]", char) != None:
				width += 1
			elif unicodedata.east_asian_width(char) in ("F", "W", "A"):
				width += 2
			else:
				width += 1
		return width


class DiffCtrl():
	"""
	ファイル差分比較に関する制御

	Attributes:
		file1     : FileCtrlオブジェクト1
		file2     : FileCtrlオブジェクト2
		lcs_cnt   : LCSインデックス数
		lcs_idx1[]: LCSインデックス値1
		lcs_idx2[]: LCSインデックス値2
		swap_flag : file1, file2の入替え表示
		mx        : 検索したX方向の最大位置(?)
		my        : 検索したY方向の最大位置(?)
		num_width : 行数の最大幅
		args      : 引数情報
		wb        : Excelのワークブック
		ws        : Excelのワークシート
		row       : Excelの出力行
	"""
	def __init__(self, file1, file2):
		if file1.count > file2.count:
			self.file1 = file2
			self.file2 = file1
			self.swap_flag = True
		else:
			self.file1 = file1
			self.file2 = file2
			self.swap_flag = False

		self.lcs_cnt = 0
		self.lcs_idx1 = []
		self.lcs_idx2 = []
		self.lcs_idx1.append(-1)
		self.lcs_idx2.append(-1)
		self.mx = -1
		self.my = -1
		self.num_width = max(self.file1.num_width, self.file2.num_width)
		self.args = None
		self.wb = None
		self.ws = None
		self.row = 1

		# 画面幅からside-by-sideの片側に表示する行の幅を算出
		terminal_size = shutil.get_terminal_size()
		term_width = terminal_size.columns
		self.sbs_width = int((term_width - 4) / 2) - (self.num_width + 1)

	def set_args(self, args):
		"""
		オプションの設定
		Args:
			tab_stp (int): タブ幅
		"""

		self.args = args

	def recover_swap(self):
		"""
		ファイル1の行数が多い時にファイル2との入替えの戻し
		"""

		if self.swap_flag:
			save_file = self.file1
			self.file1 = self.file2
			self.file2 = save_file

			save_lcs = self.lcs_idx1
			self.lcs_idx1 = self.lcs_idx2
			self.lcs_idx2 = save_lcs

	def compare(self, str1, str2):
		"""
		オプションに応じた2つの文字列の一致判定
		Args:
			str1 (str): 比較文字列1
			str2 (str): 比較文字列2
		Returns:
			bool: True:一致、False:不一致
		"""

		edit1 = str1
		edit2 = str2

		# タブ/空白の変化を無視
		if self.args.ignore_space:
			# 行頭/行末のタブ/空白削除
			edit1 = re.sub(r"^( |\t)+|( |\t)+$", "", edit1)
			edit2 = re.sub(r"^( |\t)+|( |\t)+$", "", edit2)
			# 連続するタブ/空白を1空白へ変換
			edit1 = re.sub(r"( |\t)+", " ", edit1)
			edit2 = re.sub(r"( |\t)+", " ", edit2)

		# (シングル/ダブル)クォーテーションの違いを無視
		if self.args.ignore_quote:
			edit1 = re.sub(r"\"", "'", edit1)
			edit2 = re.sub(r"\"", "'", edit2)

		if edit1 == edit2:
			return True
		else:
			return False

	def snake(self, k, y):
		"""
		fp(k,p)からエディットグラフを対角線上に進んで到達できる最大のy座標を返す
		Args:
			k (int): エディットグラフのX座標
			y (int): エディットグラフのY座標
		Returns:
			int: 到達できる最大のy座標
		"""

		x = y - k

		while (x < self.file1.count and y < self.file2.count \
				and (self.compare(self.file1.lines[x], self.file2.lines[y]))):
			if self.compare(self.file1.lines[x], self.file2.lines[y]):
				if (self.mx < x) and (self.my < y):
					self.lcs_idx1[self.lcs_cnt] = x;
					self.lcs_idx2[self.lcs_cnt] = y;
					self.lcs_idx1.append(-1)
					self.lcs_idx2.append(-1)
					self.lcs_cnt += 1
					self.mx = x;
					self.my = y

			x += 1;
			y += 1;

		return y

	def edit_dist(self):
		"""
		2つの要素列の違いを数値化した編集距離(Edit Distance)の算出
		Args:
			なし
		Returns:
			int: 編集距離
		"""

		ofset = self.file1.count + 1
		delta = self.file2.count - self.file1.count
		size  = self.file1.count + self.file2.count + 3

		fp = []
		for i in range(0, size):
			fp.append(-1)

		p = -1
		while True:
			p += 1
			for k in range(-p, delta):
				y = max(fp[k - 1 + ofset] + 1, fp[k + 1 + ofset])
				fp[k + ofset] = self.snake(k, y)

			for k in range(delta + p, delta, -1):
				y = max(fp[k - 1 + ofset] + 1, fp[k + 1 + ofset])
				fp[k + ofset] = self.snake(k, y)

			y = max(fp[delta - 1 + ofset] + 1, fp[delta + 1 + ofset])
			fp[delta + ofset] = self.snake(delta, y)

			if fp[delta + ofset] == self.file2.count:
				break

		return (delta + 2 * p)

	def from_to(self, from_line, to_line):
		"""
		From-To表示の文字列生成
		"""

		if from_line == to_line:
			return f"{from_line}"
		else:
			return f"{from_line},{to_line}"

	def get_end(self, lcs_idx, idx, file, line):
		"""
		LCS_Idxと合致しない最終行位置の取得
		"""

		while (lcs_idx[idx] != line and line < file.count):
			line += 1

		return line

	def prt_diff(self, file, stt_idx, end_idx, format):
		"""
		LCS_Idxと合致しない行の表示
		"""

		for idx in range(stt_idx, end_idx):
			#print(format % file.lines[idx])
			edit = self.edit_text(file.lines[idx])
			print(format % edit)

		return (idx + 1)

	def disp_diff(self):
		"""
		差分の表示
		"""

		i = 0  # LCSの行位置
		x = 0  # file1の行位置
		y = 0  # file2の行位置
		while (x < self.file1.count) or (y < self.file2.count):
			# 表示パターンの判定
			pid = ""
			if self.lcs_idx1[i] == -1:
				if x < self.file1.count:
					if y < self.file2.count:
						pid = "change"
					else:
						pid = "del"
				else:
					pid = "add"
			elif self.lcs_idx1[i] == x:
				if self.lcs_idx2[i] != y:
					pid = "add"
				else:
					pid = "next"
			else:
				if self.lcs_idx2[i] != y:
					pid = "change"
				else:
					pid = "del"

			# パターンに応じた表示
			if pid == "add":
				end_y = self.get_end(self.lcs_idx2, i, self.file2, y)
				line2 = self.from_to(y + 1, end_y)
				print(f"{x}a{line2}")
				y = self.prt_diff(self.file2, y, end_y, "> %s")
			elif pid == "del":
				end_x = self.get_end(self.lcs_idx1, i, self.file1, x)
				line1 = self.from_to(x + 1, end_x)
				print(f"{line1}d{y}")
				x = self.prt_diff(self.file1, x, end_x, "< %s")
			elif pid == "change":
				end_x = self.get_end(self.lcs_idx1, i, self.file1, x);
				end_y = self.get_end(self.lcs_idx2, i, self.file2, y);

				line1 = self.from_to(x + 1, end_x)
				line2 = self.from_to(y + 1, end_y)
				print(f"{line1}c{line2}")
				x = self.prt_diff(self.file1, x, end_x, "< %s")
				print("---")
				y = self.prt_diff(self.file2, y, end_y, "> %s")
			elif pid == "next":
				i += 1
				x += 1
				y += 1

	def edit_tabstop(self, org_str):
		"""
		TAB→SPACE変換
		"""
		edt_str = ""
		pos = 0
		for c in org_str:
			if c == "\t":
				cnt = self.args.tabstop - (pos % self.args.tabstop)
				edt_str += (" " * cnt)
				pos += cnt
			else:
				edt_str += c
				if re.match(r"[ a-zA-Z0-9!-\~'\-\\=+<>]", c) is not None:
					pos += 1
				else:
					pos += 2

		return edt_str

	def edit_line_width(self, org_str):
		"""
		SideBySide幅への編集
		"""
		EDGE_CHAR = "$"

		edt_str = org_str
		edge_str = ""
		pos = 0
		for i, c in enumerate(edt_str):
			if re.match(r"[ a-zA-Z0-9!-\~'\-\\=+<>]", c) is not None:
				char_width = 1
			else:
				char_width = 2

			pos += char_width

			if self.sbs_width == pos:
				# SideBySide幅と同じ場合、次のループで超えていれば
				# 最後の文字を境界文字へ変換するための前準備
				pre_str = edt_str[:i] 
				edge_str = EDGE_CHAR * char_width

			if self.sbs_width < pos:
				# SideBySide幅を超えた場合、境界文字を行末に付与
				if edge_str != "":
					edt_str = pre_str + REVERSE + edge_str + RESET
				else:
					edt_str = edt_str[:i] + REVERSE + EDGE_CHAR + RESET
				break

		if pos < self.sbs_width:
			# SideBySide幅に不足する分を空白追加
			edt_str += (" " * (self.sbs_width - pos))

		return edt_str

	def output_2column(self, fmt):
		"""
		差分を左右に並べて出力(全行)

		Args:
			fmt (str): 出力形式("text", "html", "excel")
		"""

		# 事前出力
		self.pre_output(fmt)

		i = 0  # LCSの行位置
		x = 0  # Aの行位置
		y = 0  # Bの行位置
		while (x < self.file1.count) or (y < self.file2.count):
			# 表示パターンの判定
			pid = ""
			if self.lcs_idx1[i] == -1:
				if x < self.file1.count:
					pid = "change"
				else:
					pid = "add"
			elif self.lcs_idx1[i] == x:
				if self.lcs_idx2[i] != y:
					pid = "add"
				else:
					pid = "next"
			else:
				pid = "change"

			# パターンに応じた表示
			if pid == "add":
				while (self.lcs_idx2[i] != y) and (y < self.file2.count):
					self.put_row(fmt, "", "", ">", y + 1, self.file2.lines[y])
					y += 1
			elif pid == "change":
				end_x = self.get_end(self.lcs_idx1, i, self.file1, x)
				end_y = self.get_end(self.lcs_idx2, i, self.file2, y)

				while (x < end_x) or (y < end_y):
					if x >= end_x:
						line2 = self.file2.lines[y]
						self.put_row(fmt, "", "", ">", y + 1, line2)
						y += 1
					elif y >= end_y:
						line1 = self.file1.lines[x]
						self.put_row(fmt, x + 1, line1, "<", "", "")
						x += 1
					else:
						line1 = self.file1.lines[x]
						line2 = self.file2.lines[y]
						self.put_row(fmt, x + 1, line1, "|", y + 1, line2)
						x += 1
						y += 1
			elif pid == "next":
				line1 = self.file1.lines[x]
				line2 = self.file2.lines[y]
				self.put_row(fmt, x + 1, line1, " ", y + 1, line2)
				i += 1
				x += 1
				y += 1

		# 事後出力
		self.post_output(fmt)

	def pre_output(self, fmt):
		"""
		事前出力
		"""
		if fmt == "html":
			print(HEAD_HTML)

			num_wd = self.num_width * 0.5 + 1.0
			thread = THREAD_HTML 
			thread = thread.replace("{num_width}", str(num_wd))
			thread = thread.replace("{file1}", self.file1.filename)
			thread = thread.replace("{file2}", self.file2.filename)
			print(thread)
			print("    <tbody>")
		elif fmt == "excel":
			wb = Workbook()
			ws = wb.active
			ws.title = "比較結果"
			self.wb = wb
			self.ws = ws
			self.set_title(self.file1.filename, self.file2.filename)

	def put_row(self, fmt, num1, line1, mark, num2, line2):
		"""
		差分行出力
		"""
		if fmt == "text":
			self.put_text(num1, line1, mark, num2, line2)
		elif fmt == "html":
			self.put_html(num1, line1, mark, num2, line2)
		elif fmt == "excel":
			self.put_excel(num1, line1, mark, num2, line2)

	def put_text(self, num1, line1, mark, num2, line2):
		"""
		差分を左右に並べて表示(text)
		"""
		sw = self.sbs_width
		nw = self.num_width
		rst = RESET

		base = f"%s%%{nw}s:%s%%s{rst} {mark} %s%%{nw}s:%s%%s{rst} "
		if mark == " ":
			format = base % (GREEN, RESET, GREEN, RESET)
		elif mark == "|":
			format = base % (GREEN, B_YELLOW, GREEN, B_YELLOW)
		elif mark == "<":
			format = base % (GREEN, B_YELLOW, GREEN, B_CYAN)
		elif mark == ">":
			format = base % (GREEN, B_CYAN, GREEN, B_YELLOW)
		else:
			format = ""

		edit1 = self.edit_text(line1)
		edit2 = self.edit_text(line2)
		output = format % (num1, edit1, num2, edit2)

		print(output)

	def edit_text(self, org_str):
		"""
		出力行の編集(text)
		"""
		# TAB→SPACE変換
		edt_str = self.edit_tabstop(org_str)

		# SideBySide幅への編集
		if self.args.side_by_side:
			edt_str = self.edit_line_width(edt_str)

		return edt_str

	def put_html(self, num1, line1, mark, num2, line2):
		"""
		差分行出力(html)
		"""
		if mark == " ":
			cls1 = "same"
			cls2 = "same"
		elif mark == "|":
			cls1 = "diff"
			cls2 = "diff"
		elif mark == "<":
			cls1 = "diff"
			cls2 = "null"
		elif mark == ">":
			cls1 = "null"
			cls2 = "diff"

		line1 = self.edit_html(line1)
		line2 = self.edit_html(line2)
		mark  = self.edit_html(mark)

		print('    <tr>')
		print(f'      <td class="line"><code>{num1}:</code></td>')
		print(f'      <td class="{cls1}"><code>{line1}</code></td>')
		print(f'      <td class="mark"><code>{mark}</code></td>')
		print(f'      <td class="line"><code>{num2}:</code></td>')
		print(f'      <td class="{cls2}"><code>{line2}</code></td>')
		print('    </tr>')

	def edit_html(self, org_str):
		"""
		html出力文字の編集
		"""
		# TAB→SPACE変換
		edt_str = self.edit_tabstop(org_str)

		# 特殊記号変換
		edt_str = edt_str.replace("&", "&amp;")
		edt_str = edt_str.replace("<", "&lt;")
		edt_str = edt_str.replace(">", "&gt;")
		edt_str = edt_str.replace("|", "&verbar;")

		# スペース変換
		while True:
			m = re.search(r"(^ +)|(  +)", edt_str)
			if m is None:
				break

			start = m.start()
			end   = m.end()
			count = end - start
			edt_str = edt_str[:start] + "&ensp;" * count + edt_str[end:]

		return edt_str

	def put_excel(self, num1, line1, mark, num2, line2):
		"""
		Excelの差分行設定
		"""
		ws = self.ws
		row = self.row

		# 値設定
		edit1 = self.edit_tabstop(line1)
		edit2 = self.edit_tabstop(line2)
		ws[f"A{row}"] = f"{num1}:"
		ws[f"B{row}"] = edit1
		ws[f"C{row}"] = mark
		ws[f"D{row}"] = f"{num2}:"
		ws[f"E{row}"] = edit2

		# 文字色/配置設定
		self.format(f"A{row},D{row}", color="0000ff", align="right")
		self.format(f"B{row},E{row}", color="000000", align="left", wrap=True)
		self.format(f"C{row}"       , color="000000", align="center")

		# 背景色設定
		if mark == "|":
			self.format(f"B{row}", bg_color="efcb05")
			self.format(f"E{row}", bg_color="efcb05")
		elif mark == "<":
			self.format(f"B{row}", bg_color="efcb05")
			self.format(f"E{row}", bg_color="00ccff")
		elif mark == ">":
			self.format(f"B{row}", bg_color="00ccff")
			self.format(f"E{row}", bg_color="efcb05")

		self.row += 1


	def post_output(self, fmt):
		"""
		事後出力
		"""
		if fmt == "html":
			print(TAIL_HTML)
		elif fmt == "excel":
			self.adjust_column(2)

			file = self.args.xls
			if is_open(file):
				print(f"{file}が既に開かれているため、出力できません。")
			else:
				self.wb.save(file)

	def set_title(self, title1, title2):
		"""
		Excelのタイトル行設定
		"""
		ws = self.ws
		row = self.row

		# 値設定
		ws[f"A{row}"] = title1
		ws[f"D{row}"] = title2
		ws.merge_cells(f"A{row}:B{row}")
		ws.merge_cells(f"D{row}:E{row}")

		# 文字色/背景色/配置設定
		self.format(f"A{row},B{row},D{row},E{row}", size=14,
				color="ffffff", bg_color="0000ff", align="center", wrap=True)

		# 枠固定設定
		ws.freeze_panes = f"A{row+1}"

		self.row += 1

	def adjust_column(self, start_row=1):
		"""
		Excelの列幅調整
		"""
		font_depend = 1.2
		for col in self.ws.columns:
			max_length = 0
			column = col[1].column_letter # Get the column name
			row = 0
			for cell in col:
				row += 1
				if row < start_row:
					continue
				if (cell.font.size is None) or (cell.value is None):
					continue

				cell_length = 0
				for char in str(cell.value):
					if east_asian_width(char) in ("F", "W", "A"):
						# 全角文字の幅設定
						cell_length += 2
					else:
						# 半角文字の幅設定
						cell_length += 1

				if cell_length > max_length:
					max_length = cell_length

			self.ws.column_dimensions[column].width = (max_length + 2) * 1.2

	# アドレスの書式チェック
	def is_adr(self, adr):
		if re.fullmatch(r"[a-zA-Z]+[1-9][0-9]*", adr):
			return True
		else:
			return False

	# アドレス→行列変換(A1->row, col)
	def adr_to_cell(self, adr):
		col = 0
		row = 0
		adr = adr.upper()
		for c in adr:
			if c >= "A" and c <= "Z":
				col = 26 * col + (ord(c) - ord("A") + 1)
			elif c >= "0" and c <= "9":
				row = 10 * row + (ord(c) - ord("0"))

		return row, col


	# 行列→アドレス変換(row, col->A1)
	def cell_to_adr(self, row, col):
		adr = ""
		while col > 0:
			idx = (col - 1) % 26 + 1
			col = (col - 1) // 26
			adr = chr(64 + idx) + adr

		adr = adr + str(row)
		return adr


	# From-Toを含む複数アドレスの分解
	def adr_to_list(self, org_adrs):
		adrs = []
		org_adrs = org_adrs.upper()
		for col_adr in org_adrs.split(","):
			range_adr = col_adr.split(":")
			count = len(range_adr)
			if count == 1:
				# 単一アドレスの格納
				adr1 = range_adr[0]
				if self.is_adr(adr1):
					adrs.append(adr1)
				else:
					return f"Format error({adr1})"
			elif count == 2:
				# From-To形式のアドレス展開
				adr1 = range_adr[0]
				adr2 = range_adr[1]
				# アドレスの書式チェック
				if not self.is_adr(adr1):
					return f"Format error({adr1})"
				if not self.is_adr(adr2):
					return f"Format error({adr2})"

				# アドレスから行列への変換
				(row1, col1) = self.adr_to_cell(adr1)
				(row2, col2) = self.adr_to_cell(adr2)

				# 行列からアドレスへ変換してリストへ格納
				if (row1 <= row2 and col1 <= col2):
					for row in range(row1, row2 + 1):
						for col in range(col1, col2 + 1):
							adrs.append(self.cell_to_adr(row, col))
				elif (row1 >= row2 and col1 >= col2):
					for row in range(row2, row1 + 1):
						for col in range(col2, col1 + 1):
							adrs.append(self.cell_to_adr(row, col))
				else:
					return f"Format error({col_adr})"

			else:
				return f"Format error({col_adr})"

		return adrs


	# セルのフォント/文字色/背景色/配置設定
	def format(self, adrs, name="", size="", color="", bg_color="",
				align="", vertical="", wrap=""):

		ws = self.ws

		# 複数アドレスの展開
		adr_list = self.adr_to_list(adrs)
		if type(adr_list) is str: 
			print(result)
			sys.exit(1)

		for adr in adr_list:
			# フォント/文字色設定
			if (name or size or color):
				if name == "":
					name="ＭＳ ゴシック"
				if size == "":
					size = 12
				if color == "":
					color = "000000"
				ws[adr].font = Font(name=name, size=size, color=color)

			# 背景色設定
			if bg_color:
				ws[adr].fill = PatternFill("solid", fgColor=bg_color)

			# 配置設定
			if (align or vertical or wrap):
				if align == "":
					align="center"
				if vertical == "":
					vertical="top"
				if wrap == "":
					wrap = False
				ws[adr].alignment \
				= Alignment(horizontal=align, vertical=vertical, wrapText=wrap)


# === common function ==========================================================
if os.name == "nt":
	def enable():
		"""
		コマンドプロンプトのESCシーケンス有効化
		"""

		INVALID_HANDLE_VALUE = -1
		STD_INPUT_HANDLE  = -10
		STD_OUTPUT_HANDLE = -11
		STD_ERROR_HANDLE  = -12
		ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
		ENABLE_LVB_GRID_WORLDWIDE = 0x0010

		hOut = windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
		if hOut == INVALID_HANDLE_VALUE:
			return False
		dwMode = wintypes.DWORD()
		if windll.kernel32.GetConsoleMode(hOut, byref(dwMode)) == 0:
			return False
		dwMode.value |= ENABLE_VIRTUAL_TERMINAL_PROCESSING
		if windll.kernel32.SetConsoleMode(hOut, dwMode) == 0:
			return False
		return True


def is_open(filepath):
	"""
	ファイルを追記モードで開けるかチェック
	"""
	try:
		f = open(filepath, 'a')
		f.close()
	except:
		return True
	else:
		return False


def get_args():
	"""
	引数を解析する
	Args:
		なし
	Returns:
		obj: parser.parse_args()
	"""

	parser = argparse.ArgumentParser(
				prog="diff.py",
				description="Compare 2 files and print difference."
				)

	parser.add_argument("filename1", help="Compare filename1")
	parser.add_argument("filename2", help="Compare filename2")

	parser.add_argument(
			"-s", "--side_by_side", action="store_true",
			help="output in two columns"
			)
	parser.add_argument(
			"-m", "--html", action="store_true",
			help="output html format in two columns"
			)
	parser.add_argument(
			"-e", "--xls", type=str, default="", metavar="FILE",
			help="output excel format in two columns"
			)
	parser.add_argument(
			"-x", "--tabstop", type=int, default=4, metavar="N",
			help="set tab stops"
			)
	parser.add_argument(
			"-b", "--ignore_space", action="store_true",
			help="ignore changes in the amount of white space"
			)
	parser.add_argument(
			"-q", "--ignore_quote", action="store_true",
			help="ignore quotation difference"
			)

	return parser.parse_args()


# === main =====================================================================
if __name__ == "__main__":
	if os.name == "nt":
		# ESCシーケンス有効化
		enable()

	# 引数の解析
	args = get_args()

	# FileCtrlオブジェクト生成
	file1 = FileCtrl(args.filename1)
	file2 = FileCtrl(args.filename2, args.filename1)

	# DiffCtrlオブジェクト生成
	diff = DiffCtrl(file1, file2)

	# ファイルの差分検出
	diff.set_args(args)
	diff.edit_dist()
	diff.recover_swap()

	# ファイル差分の出力
	if args.side_by_side:
		# text形式で左右に並べての出力
		diff.output_2column("text")
	elif args.html:
		# html形式で左右に並べての出力
		diff.output_2column("html")
	elif args.xls:
		# Excel形式で左右に並べての出力
		from openpyxl import Workbook
		from openpyxl.styles import (
				PatternFill, GradientFill, Border, Font, Alignment)
		diff.output_2column("excel")
	else:
		# text形式の差分出力
		diff.disp_diff()
