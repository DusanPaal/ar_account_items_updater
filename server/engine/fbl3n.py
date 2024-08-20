# pylint: disable = C0103, C0123, W0603, W0703, W1203

"""
Description:
------------
The module provides the following exclusive procedures:
	- change_document_parameters():
		For modification of 'Text' and 'Assignment' fields 
		of items posted to G/L accounts.

	- export_line_items():
		For exporting accounting item data from G/L accounts
		to raw text.

How to use:
-----------
The FBL3N must be started by calling the `start()` procedure.

Attempt to use an exclusive procedure when FBL3N has not been 
started results in the `UninitializedModuleError` exception.

After using the module, the transaction should be closed
and the resources released by calling the `close()` procedure.

Version history:
----------------
1.0.20210720 - Initial version.
"""

import os
from copy import deepcopy
from datetime import date
from logging import Logger, getLogger
from os.path import exists, split
from typing import overload, Union
from pyperclip import copy as copy_to_clipboard
from win32com.client import CDispatch

FilePath = str

_sess: CDispatch = None     	# type: ignore
_main_wnd: CDispatch = None 	# type: ignore
_stat_bar: CDispatch = None 	# type: ignore

_virtual_keys = {
	"Enter":        0,
	"F2":           2,
	"F3":           3,
	"F4":           4,
	"F6":           6,
	"F8":           8,
	"F9":           9,
	"CtrlS":        11,
	"F12":          12,
	"ShiftF1":      13,
	"ShiftF2":      14,
	"ShiftF4":      16,
	"ShiftF12":     24,
	"CtrlF1":       25,
	"CtrlF8":       32,
	"CtrlShiftF2":  38,
	"CtrlShiftF6":  42
}

log: Logger = getLogger("master")

# custom exceptions and warnings
class ItemLoadingError(Exception):
	"""Raised when loading of accounting items fails.
	
	The possible reasons include:
		- no account fulfils the selection conditions
	"""

class NoItemsFoundWarning(Warning):
	"""Raised when no items are found on account(s)
	using the specified selection criteria.
	"""

class NoItemsFoundError(Exception):
	"""Raised when no items are found by filtering
	a field on the specified values.
	"""

class DataExportError(Exception):
	"""Raised when data export fails."""

class UninitializedModuleError(Exception):
	"""Raised when attempting to use an exclusive
	procedure before starting the transaction.
	"""

class FolderNotFoundError(Exception):
	"""Raised when a folder is reruired but doesn't exist."""

class SapConnectionLostError(Exception):
	"""Raised when the connection to SAP is lost."""


def _clear_clipboard() -> None:
	"""Clears the contents of the clipboard."""
	copy_to_clipboard("")

def _press_key(name: str) -> None:
	"""Simulates pressing a keyboard button."""
	_main_wnd.SendVKey(_virtual_keys[name])

def _is_popup_dialog() -> bool:
	"""Checks if the active window is a popup dialog window."""
	return _sess.ActiveWindow.type == "GuiModalWindow"

def _close_popup_dialog(confirm: bool) -> None:
	"""Confirms or declines a pop-up dialog."""

	if _sess.ActiveWindow.text == "Information":
		if confirm:
			_press_key("Enter") # confirm
		else:
			_press_key("F12")   # decline
		return

	btn_caption = "Yes" if confirm else "No"

	for child in _sess.ActiveWindow.Children:
		for grandchild in child.Children:
			if grandchild.Type != "GuiButton":
				continue
			if btn_caption != grandchild.text.strip():
				continue
			grandchild.Press()
			return

def _set_text(val: str) -> None:
	"""Enters new value into the item 'Text' field."""

	if len(val) > 50:
		raise ValueError(
			f"The length of the entered value '{val}' exceeds "
			"the allowed maximum of 50 chars for this field!")

	_main_wnd.findByName("BSEG-SGTXT", "GuiCTextField").text = val

def _set_assignment(val: str) -> None:
	"""Enters new value into the item 'Text' field."""

	if _main_wnd.findAllByName("BSEG-ZUONR", "GuiTextField").count == 0:
		return

	if len(val) > 18:
		raise ValueError(f"The length of the entered value '{val}' "
		"exceeds the allowed maximum of 18 chars for this field!")

	_main_wnd.findByName("BSEG-ZUONR", "GuiTextField").text = val

def _set_layout(val: str) -> None:
	"""Enters layout name into the 'Layout' field
	located on the main transaction window.
	"""
	_main_wnd.findByName("PA_VARI", "GuiCTextField").text = val

def _set_worklist(val: str) -> None:
	"""Enters worklist name into the 'Worklist' field
	located on the main transaction window.
	"""
	_main_wnd.findByName("PA_WLSAK", "GuiCTextField").text = val

def _set_company_code(val: str) -> None:
	"""Enters company code into the 'Company code' field
	located on the transaction main window.
	"""

	if not (val == "" or (len(val) == 4 and val.isnumeric())):
		raise ValueError(f"Company code has incorrect value: '{val}'!")

	if _main_wnd.findAllByName("SD_BUKRS-LOW", "GuiCTextField").count > 0:
		_main_wnd.findByName("SD_BUKRS-LOW", "GuiCTextField").text = val
	elif _main_wnd.findAllByName("SO_WLBUK-LOW", "GuiCTextField").count > 0:
		_main_wnd.findByName("SO_WLBUK-LOW", "GuiCTextField").text = val

def _set_posting_dates(first: date, last: date) -> None:
	"""Enters start and end posting dates in the transaction main window
	that define the date range for which accounting data will be loaded.
	"""

	if not (first is None and last is None):
		log.debug(f"Export date range: {first} - {last}")
		if first > last:
			raise ValueError(
				"Lower posting date is greater than upper posting date!")

	sap_date_format = "%d.%m.%Y"
	date_from = "" if first is None else first.strftime(sap_date_format)
	date_to = "" if last is None else last.strftime(sap_date_format)

	_main_wnd.FindByName("SO_BUDAT-LOW", "GuiCTextField").text = date_from
	_main_wnd.FindByName("SO_BUDAT-HIGH", "GuiCTextField").text = date_to

def _set_accounts(vals: list) -> None:
	"""Inserts GL accounts into the appropriate
	search field in the FBL3N main mask.
	"""

	if len(vals) == 0:
		raise ValueError("No G/L accounts found!")

	for val in vals:
		if not str(val).isnumeric():
			raise ValueError(f"Invalid account number: {val}!")

	# open selection table for company codes
	_main_wnd.findByName("%_SD_SAKNR_%_APP_%-VALU_PUSH", "GuiButton").press()

	accs = list(map(str, vals))
	_press_key("ShiftF4")   				# clear any previous values
	copy_to_clipboard("\r\n".join(accs))    # copy accounts to clipboard
	_press_key("ShiftF12")  				# confirm selection
	_clear_clipboard()
	_press_key("F8")        				# confirm the entered values

def _set_line_items_selection(status: str) -> None:
	"""Sets line item selection mode by item status."""

	if status == "open":
		obj_name = "X_OPSEL"
	elif status == "cleared":
		obj_name = "X_CLSEL"
	elif status =="all":
		obj_name = "X_AISEL"
	else:
		raise ValueError(f"Unrecognized item status: '{status}'")

	_main_wnd.FindByName(obj_name, "GuiRadioButton").Select()

def _toggle_worklist(activate: bool) -> None:
	"""Activates or deactivates the 'Use worklist'
	option in the transaction main search mask.
	"""

	used = _main_wnd.FindAllByName("PA_WLSAK", "GuiCTextField").Count != 0

	if (activate and not used) or (not activate and used):
		_press_key("CtrlF1")

def _add_filter_criterion() -> None:
	"""Adds a filter criterion to a filtered field."""

	_sess.findById("wnd[1]").findByName("APP_WL_SING", "GuiButton").press()

def _get_filter_list() -> CDispatch:
	"""Returns a reference to the table object containing a list of unused filters."""

	return _sess.FindById("wnd[1]").FindAllByName("shell", "GuiApoGrid")(1)

def _set_filter(vals: list, fld_tech_name: str = "SGTXT") -> None:
	"""Applies a filter on a table field (default 'Text' field)."""

	_press_key("CtrlShiftF2")  # open Set Filter Values dialog
	_press_key("CtrlShiftF6")  # toggle technical names

	filters = _get_filter_list()

	for row_idx in range(0, filters.RowCount):
		if filters.GetCellValue(row_idx, "FIELDNAME") == fld_tech_name:
			filters.selectedRows = row_idx
			_add_filter_criterion()
			break

	# press "Define filter values" button
	_sess.findById("wnd[1]").findByName("600_BUTTON", "GuiButton").press()

	# open value list box
	_sess.findById("wnd[2]").findByName("%_%%DYN001_%_APP_%-VALU_PUSH", "GuiButton").press()

	copy_to_clipboard("\r\n".join(vals))    # copy data to clipboard
	_press_key("ShiftF12")  				# paste data from clipboard
	_clear_clipboard()
	_press_key("F8")        				# confirm inserted values
	_press_key("Enter")     				# confirm the filter

def _load_items() -> CDispatch:
	"""Loads items located on G/L account(s) and returns the item table."""

	try:
		_press_key("F8")
	except Exception as exc:
		raise ItemLoadingError("Could not load account data!") from exc

	try: # SAP crash can be caught only after next statement following item loading
		msg = _stat_bar.Text
	except Exception as exc:
		raise SapConnectionLostError("Connection to SAP lost!") from exc

	if "No items selected" in msg:
		raise NoItemsFoundWarning("No items found using your selection criteria!")

	if "items displayed" not in msg:
		raise ItemLoadingError(msg)

	items = _main_wnd.FindById("usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")

	return items

def _get_item_params(tbl: CDispatch, idx: int) -> tuple:
	"""Returns a tuple of 'Text' and 'Assignment' values of an item."""

	tbl.selectedRows = idx
	tbl.currentCellRow = idx
	text = tbl.GetCellValue(idx, "SGTXT")
	assignment = tbl.GetCellValue(idx, "ZUONR")

	return (text, assignment)

def _export_to_file(file: FilePath) -> None:
	"""Exports data to a local file."""

	folder_path, file_name = split(file)

	if not exists(folder_path):
		raise FolderNotFoundError(
			"The export folder not found at the "
			f"path specified: '{folder_path}'!")

	# open local data file export dialog
	_press_key("F9")

	# set plain text data export format and confirm
	_sess.FindById("wnd[1]").FindAllByName("SPOPLI-SELFLAG", "GuiRadioButton")(0).Select()
	_press_key("Enter")

	# enter data export file name, folder path and encoding
	# then click 'Replace' an existing file button
	folder_path = "".join([folder_path, "\\"])
	utf_8 = "4120"

	_sess.FindById("wnd[1]").FindByName("DY_PATH", "GuiCTextField").text = folder_path
	_sess.FindById("wnd[1]").FindByName("DY_FILENAME", "GuiCTextField").text = file_name
	_sess.FindById("wnd[1]").FindByName("DY_FILE_ENCODING", "GuiCTextField").text = utf_8
	_press_key("CtrlS")

def _read_exported_data(file_path: str) -> str:
	"""Reads exported FBL3N data from the text file."""

	with open(file_path, encoding = "utf-8") as stream:
		text = stream.read()

	return text

def _check_prerequisities():
	"""Verifies that the prerequisites
	for using the module are met."""

	if _sess is None:
		raise UninitializedModuleError(
			"Uninitialized module! Use the start() "
			"procedure to run the transaction first!")

def start(sess: CDispatch) -> None:
	"""Starts the FBL3N transaction.

	If the FBL3N has already been started,
	then the transaction will be restarted.

	Parameters:
	-----------
	sess:
		An SAP `Gui_sess` object (wrapped in
		the win32:CDispatch class)that represents
		an active user SAP GUI _sess.
	"""

	global _sess
	global _main_wnd
	global _stat_bar

	if sess is None:
		raise UnboundLocalError("Argument 'sess' is unbound!")

	# close the transaction
	# if it is already running
	close()

	_sess = sess
	_main_wnd = _sess.findById("wnd[0]")
	_stat_bar = _main_wnd.findById("sbar")

	_sess.StartTransaction("FBL3N")

def close() -> None:
	"""Closes a running FBL3N transaction.

	Attempt to close the transaction that has not been
	started by the `start()` procedure is ignored.
	"""

	global _sess
	global _main_wnd
	global _stat_bar

	if _sess is None:
		return

	_sess.EndTransaction()

	if _is_popup_dialog():
		_close_popup_dialog(confirm = True)

	_sess = None        # type: ignore
	_main_wnd = None    # type: ignore
	_stat_bar = None    # type: ignore

def change_document_parameters(
		accounts: list,
		company_code: str,
		parameters: dict,
		status: str = "open",
		layout: str = ""
	) -> dict:
	"""Replaces the 'Text' and 'Assignment' parameters of accounting items.

	If loading of accounting items fails, then an `ItemLoadingError` exception is raised.

	If no items are found using the company code and G/L accounts,
	then an `NoItemsFoundWarning` warning is raised.

	If no items are found when filtering the table on the 'Text' field
	using the specified text values, then `NoItemsFoundError` is raised.

	Parameters:
	-----------
	accounts:
		Numbers of G/L accounts with line items to change.

	company_code:
		The company code for which the accounting data will is changed.

	status:
		Item status to consider for selection:
		- "open": Open items will be exported (default).
		- "cleared": Cleared items will be exported.
		- "all": All items will be exported.

	parameters:
		Original item text values mapped to their new 'Text' and
		'Assignment' values. The data is structured as follows:

		{
			"old_text_value_1": {
				"new_text": "value"
				"new_assignment": "value"
			},

			"old_text_value_2": {
				"new_text": "value"
				"new_assignment": "value"
			},

			...

			"old_text_value_n": {
				"new_text": "value"
				"new_assignment": "value"
			}
		}

	layout:
		The name of the layout that defines the format of the loaded data.

		By default, no specific layout name is used,
		and a the predefined FLB3N layout is used.

	Returns:
	--------
	The processing results in the following structure:
		{
			"old_text_value_1": {
				"new_text": "value"
				"new_assignment": "value"
				“message”: “value”
			},
			"old_text_value_2": {
				"new_text": "value"
				"new_assignment": "value"
				message”: “value”
			},
			...
			"old_text_value_n": {
				"new_text": "value"
				"new_assignment": "value"
				message”: “value”
			}
		}
	"""

	_check_prerequisities()
	_toggle_worklist(activate = False)
	_set_company_code(company_code)
	_set_layout(layout)
	_set_accounts(accounts)
	_set_line_items_selection(status)

	item_table = _load_items()
	_set_filter(list(parameters.keys()))

	if item_table.RowCount == 0:
		_press_key("F3")
		raise NoItemsFoundError(
			"Filtering on the searched text "
			"values returned no results!")

	output = deepcopy(parameters)

	# write a default message to the  output
	for itm in output:
		output[itm]["message"] = "Document not found on the account!"

	for idx in range(0, item_table.RowCount):

		text_differs = True
		assign_differs = True
		old_text, old_assign = _get_item_params(item_table, idx)

		new_text = parameters[old_text]["new_text"]
		new_assign = parameters[old_text]["new_assignment"]

		output[old_text]["message"] = ""

		if new_text is not None and new_text == old_text:
			output[old_text]["message"] += "Text aready contains the desired value. "
			text_differs = False

		if new_assign is not None and old_assign == new_assign:
			output[old_text]["message"] += "Assignment aready contains the desired value. "
			assign_differs = False

		if not (text_differs or assign_differs):
			output[old_text]["message"] = output[old_text]["message"].strip()
			continue

		_press_key("ShiftF2")
		_press_key("ShiftF1")

		if new_text is not None and text_differs:
			_set_text(str(new_text))
			output[old_text]["message"] += "Text updated. "

		if new_assign is not None and assign_differs:
			_set_assignment(str(new_assign))
			output[old_text]["message"] += "Assignment updated. "

		output[old_text]["message"] = output[old_text]["message"].strip()

		_press_key("CtrlS")

	_press_key("F3")

	return output

@overload
def export_line_items(
		file: FilePath,
		worklist: str,
		company_code: str,
		status: str = "open",
		from_day: date = None,  # type: ignore
		to_day: date = None, 	# type: ignore
		layout: str = ""
	) -> str:
	"""Exports item data from G/L accounts.

	A `NoItemsFoundWarning` warning will be raised
	if no items are found for the given selection criteria.

	A `DataExportError` exception is raised
	if the attempt to expot accounting data fails.

	A `SapConnectionLostError` exception is raised
	when the connection to SAP is lost due to an error.

	Parameters:
	-----------
	file:
		Path to a temporary .txt file to which the data will be exported.
		If the file path points to an invalid folder, a `FolderNotFoundError`
		exception is raised. The file is removed when the data reading is complete.

	worklist:
		Name of the worklist that contains
		G/L accounts from which data is exported.

	company_code:
		The company code to which the accounts are assigned.
		A valid company code is a 4-digit string (e.g. '0075').

	status:
		Item status to consider for selection:
			- "open": Open items will be exported (default).
			- "cleared": Cleared items will be exported.
			- "all": All items will be exported.

	from_day:
		Posting date from which accounting data is exported.
		By default, no past limit is used.

	to_day:
		Posting date up to which accounting data is exported.
		By default, no future limit is used.

	layout:
		The name of the layout that defines the format of the loaded data.
		
		By default, no specific layout name is used,
		and a the predefined FLB3N layout is used.

	Returns:
	--------
	The exported line item data as plain text.
	"""

@overload
def export_line_items(
		file: FilePath,
		accounts: list,
		company_code: str,
		status: str = "open",
		from_day: date = None,  # type: ignore
		to_day: date = None, 	# type: ignore
		layout: str = ""
	) -> str:
	"""Exports item data from G/L accounts.

	A `NoItemsFoundWarning` warning will be raised
	if no items are found for the given selection criteria.

	A `DataExportError` exception is raised
	if the attempt to expot accounting data fails.

	A `SapConnectionLostError` exception is raised
	when the connection to SAP is lost due to an error.

	Parameters:
	-----------
	file:
		Path to a temporary .txt file to which the data will be exported.
		If the file path points to an invalid folder, a `FolderNotFoundError`
		exception is raised. The file is removed when the data reading is complete.

	accounts:
		G/L accounts stored as `int` from which data is exported.

	company_code:
		The company code to which the accounts are assigned.
		A valid company code is a 4-digit string (e.g. '0075').

	status:
		Item status to consider for selection:
			- "open": Open items will be exported (default).
			- "cleared": Cleared items will be exported.
			- "all": All items will be exported.

	from_day:
		Posting date from which the accounting data will be loaded.
		By default, no historical date limit is used.

	to_day:
		Posting date up to which the accounting data will be loaded.
		By default, no future date limit is used.

	layout:
		The name of the layout that defines the format of the loaded data.

		By default, no specific layout name is used,
		and a the predefined FLB3N layout is used.

	Returns:
	--------
	The exported line item data as plain text.
	"""

def export_line_items(
		file: FilePath,
		selection: Union[list,str],
		company_code: str,
		status: str = "open",
		from_day: date = None,  # type: ignore
		to_day: date = None, 	# type: ignore
		layout: str = ""
	) -> str:
	"""Exports item data from G/L accounts.

	A `NoItemsFoundWarning` warning will be raised
	if no items are found for the given selection criteria.

	A `DataExportError` exception is raised
	if the attempt to expot accounting data fails.

	A `SapConnectionLostError` exception is raised
	when the connection to SAP is lost due to an error.

	Parameters:
	-----------
	file:
		Path to a temporary .txt file to which the data will be exported.
		If the file path points to an invalid folder, a `FolderNotFoundError`
		exception is raised. The file is removed when the data reading is complete.

	selection:
		Criteria used to load items:
		- A `list[str, int]` object is interpreted as
			a list of G/L accounts from which to export data.
		- A `str` object is interpreted as the name of the worklist
			of G/L accounts from which to export data.

	company_code:
		The company code to which the accounts are assigned.
		A valid company code is a 4-digit string (e.g. '0075').

	status:
		Item status to consider for selection:
			- "open": Open items will be exported (default).
			- "cleared": Cleared items will be exported.
			- "all": All items will be exported.

	from_day:
		Posting date from which the accounting data will be loaded.
		By default, no historical date limit is used.

	to_day:
		Posting date up to which the accounting data will be loaded.
		By default, no future date limit is used.

	layout:
		The name of the layout that defines the format of the loaded data.

		By default, no specific layout name is used,
		and a the predefined FLB3N layout is used.

	Returns:
	--------
	The exported line item data as plain text.
	"""

	_check_prerequisities()

	if isinstance(selection, list):
		_toggle_worklist(activate = False)
	else:
		_toggle_worklist(activate = True)
		_set_worklist(selection)

	_set_company_code(company_code)
	_set_layout(layout)

	if isinstance(selection, list):
		_set_accounts(selection)

	_set_line_items_selection(status)
	_set_posting_dates(from_day, to_day)
	_load_items()
	_press_key("CtrlF8")        # open layout mgmt dialog
	_press_key("CtrlShiftF6")   # toggle technical names
	_press_key("Enter")         # Confirm Layout Changes
	_export_to_file(file)
	_press_key("F3")            # Load main mask
	data = _read_exported_data(file)

	try:
		os.remove(file)
	except (PermissionError, FileNotFoundError) as exc:
		log.error(exc)

	return data
