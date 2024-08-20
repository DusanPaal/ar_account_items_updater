# pylint: disable = E0110, E1101

"""
Description:
------------
The module creates user reports from the processing output.
"""

from os.path import exists, dirname
from pandas import DataFrame, Series, ExcelWriter

FilePath = str

class FolderNotFoundError(Exception):
	"""Raised when a directory is
	requested but doesn't exist.
	"""

def _get_col_width(vals: Series, col_name: str, add_width: int = 0) -> int:
	"""
	Returns an integer representing the width of a column calculated \n
	as the maximum number of characters contained in column name and \n
	column values plus additional points provided with the 'add_width'
	argument (default 0 points).
	"""

	alpha = 1 # additional offset factor

	data_vals = vals.astype("string").dropna().str.len()
	data_vals = list(data_vals)
	data_vals.append(len(str(col_name)))
	width = max(data_vals) + alpha + add_width

	return width

def _col_to_rng(
		data: DataFrame,
		first_col: str,
		last_col: str = "",
		row: int = -1,
		last_row: int = -1
	) -> str:
	"""
	Converts data position in a DataFrame object into excel range notation (e.g. 'A1:D1', 'B2:G2'). \n
	If 'last_col' is None, then only single-column range will be generated (e.g. 'A:A', 'B1:B1'). \n
	If 'row' is '-1', then the generated range will span all the column(s) rows (e.g. 'A:A', 'E:E'). \n
	If 'last_row' is provided, then the generated range will include all data records up to the last
	row (including).

	Parameters:
	-----------
	data:
		Data for which colum names should be converted to a range.

	first_col:
		Name of the first column.

	last_col:
		Name of the last column.

	row:
		Index of the row for which the range will be generated.

	last_row:
		Index of the last data row which location will be considered in the resulting range.

	Returns:
	--------
	Excel data range string.
	"""

	empty_str = ""

	if isinstance(first_col, str):
		first_col_idx = data.columns.get_loc(first_col)
	elif isinstance(first_col, int):
		first_col_idx = first_col
	else:
		raise TypeError(f"Argument 'first_col' has invalid type: {type(first_col)}")

	first_col_idx += 1
	prim_lett_idx = first_col_idx // 26
	sec_lett_idx = first_col_idx % 26

	lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else empty_str
	lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else empty_str
	lett = "".join([lett_a, lett_b])

	if last_col == "":
		last_lett = lett
	else:

		if isinstance(last_col, str):
			last_col_idx = data.columns.get_loc(last_col)
		elif isinstance(last_col, int):
			last_col_idx = last_col
		else:
			raise TypeError(f"Argument 'last_col' has invalid type: {type(last_col)}")

		last_col_idx += 1
		prim_lett_idx = last_col_idx // 26
		sec_lett_idx = last_col_idx % 26

		lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else empty_str
		lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else empty_str
		last_lett = "".join([lett_a, lett_b])

	if row == -1:
		rng = ":".join([lett, last_lett])
	elif first_col == last_col and row != -1 and last_row == -1:
		rng = f"{lett}{row}"
	elif first_col == last_col and row != -1 and last_row != -1:
		rng = ":".join([f"{lett}{row}", f"{lett}{last_row}"])
	elif first_col != last_col and row != -1 and last_row == -1:
		rng = ":".join([f"{lett}{row}", f"{last_lett}{row}"])
	elif first_col != last_col and row != -1 and last_row != -1:
		rng = ":".join([f"{lett}{row}", f"{last_lett}{last_row}"])
	else:
		assert False, "Undefined argument combination!"

	return rng

def generate_excel_report(file: FilePath, data: DataFrame, sht_name: str) -> None:
	"""Creates an excel report from the processing outcome.

	Parameters:
	-----------
	file:
		Path to the .xlsx report file to create.

	data:
		The processing outcome containing records to write.

	sht_name:
		Name of the report sheet.
	"""

	dst_dir = dirname(file)

	if not file.lower().endswith(".xlsx"):
		raise ValueError(f"TUnsupported report file format: '{file}'")

	if not exists(dst_dir):
		raise FolderNotFoundError(f"Destination folder not found: {dst_dir}")

	if sht_name == "":
		raise ValueError("Sheet name cannot be an empty string!")

	header_row_index = 1

	with ExcelWriter(file, engine="xlsxwriter") as wrtr:

		report = wrtr.book

		# define visual formats which will be applied to particular fields
		align_fmt = report.add_format({ # type: ignore
			"align": "center"
		})

		header_fmt = report.add_format({ # type: ignore
			"align": "center",
			"bg_color": "#09275E",
			"font_color": "white",
			"bold": True
		})

		data.columns = ["Account", "Old Text", "New Text", "New Assignment", "Message"]
		data.to_excel(wrtr, sht_name, index = False)

		# get the workbook sheet and apply field formats
		sht = wrtr.sheets[sht_name]

		for idx, col in enumerate(data.columns):
			col_width = _get_col_width(data[col], col)
			sht.set_column(idx, idx, col_width, align_fmt)

		rng = _col_to_rng(data, data.columns[0], data.columns[-1], header_row_index)
		sht.conditional_format(rng, {"type": "no_errors", "format": header_fmt})
