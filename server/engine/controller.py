"""
Description:
------------
The controller.py represents the middle layer of the application design
and mediates communication between the top layer (app.py) and the
highly specialized modules situated on the bottom layer of the design
(fbl3n.py, fbl5n.py, mails.py, report.py sap.py).

Version history:
----------------
1.0.20210720 - Initial version.
"""

import logging
import os
import re
from datetime import datetime as dt
from datetime import timedelta
from glob import glob
from logging import Logger, config
from os.path import basename, isfile, join
from typing import Union

import pandas as pd
import yaml
from pandas import DataFrame
from win32com.client import CDispatch

from . import fbl3n, fbl5n, mails, report, sap

ACCOUNT_TYPE_CUSTOMER = "customer"
ACCOUNT_TYPE_GENERAL_LEDGER = "general_lenger"

log = logging.getLogger("master")


# ====================================
# initialization of the logging system
# ====================================

def _compile_log_path(log_dir: str) -> str:
	"""Compiles the path to the log file
	by generating a log file name and then
	concatenating it to the specified log
	directory path."""

	date_tag = dt.now().strftime("%Y-%m-%d")
	nth = 0

	while True:
		nth += 1
		nth_file = str(nth).zfill(3)
		log_name = f"{date_tag}_{nth_file}.log"
		log_path = join(log_dir, log_name)

		if not isfile(log_path):
			break

	return log_path

def _read_log_config(cfg_path: str) -> dict:
	"""Reads logging configuration parameters from a yaml file."""

	# Load the logging configuration from an external file
	# and configure the logging using the loaded parameters.

	if not isfile(cfg_path):
		raise FileNotFoundError(f"The logging configuration file not found: '{cfg_path}'")

	with open(cfg_path, encoding = "utf-8") as stream:
		content = stream.read()

	return yaml.safe_load(content)

def _update_log_filehandler(log_path: str, logger: Logger) -> None:
	"""Changes the log path of a logger file handler."""

	prev_file_handler = logger.handlers.pop(1)
	new_file_handler = logging.FileHandler(log_path)
	new_file_handler.setFormatter(prev_file_handler.formatter)
	logger.addHandler(new_file_handler)

def _print_log_header(logger: Logger, header: list, terminate: str = "\n") -> None:
	"""Prints header to a log file."""

	for nth, line in enumerate(header, start = 1):
		if nth == len(header):
			line = f"{line}{terminate}"
		logger.info(line)

def _remove_old_logs(logger: Logger, log_dir: str, n_days: int) -> None:
	"""Removes old logs older than the specified number of days."""

	old_logs = glob(join(log_dir, "*.log"))
	n_days = max(1, n_days)
	curr_date = dt.now().date()

	for log_file in old_logs:
		log_name = basename(log_file)
		date_token = log_name.split("_")[0]
		log_date = dt.strptime(date_token, "%Y-%m-%d").date()
		thresh_date = curr_date - timedelta(days = n_days)

		if log_date < thresh_date:
			try:
				logger.info(f"Removing obsolete log file: '{log_file}' ...")
				os.remove(log_file)
			except PermissionError as exc:
				logger.error(str(exc))

def configure_logger(log_dir: str, cfg_path: str, *header: str) -> None:
	"""Configures the application's logging system.

	Parameters:
	-----------
	log_dir:
		The path to the directory where 
	    the log file will be stored.

	cfg_path:
		The path to a YAML/YML file containing 
		the application’s configuration parameters

	header:
		A sequence of lines to be printed  
		into the log header.
	"""

	log_path = _compile_log_path(log_dir)
	log_cfg = _read_log_config(cfg_path)
	config.dictConfig(log_cfg)
	logger = logging.getLogger("master")
	_update_log_filehandler(log_path, logger)
	if header is not None:
		_print_log_header(logger, list(header))
	_remove_old_logs(logger, log_dir, log_cfg.get("retain_logs_days", 1))


# ====================================
# 		application configuration
# ====================================

def load_app_config(cfg_path: str) -> dict:
	"""Reads application configuration 
	parameters from a YAML/YML file.

	Parameters:
	-----------
	cfg_path:
		The path to the YAML/YML file 
		containing the application's 
		configuration parameters

	Returns:
	--------
	A dictionary containing the application 
	configuration parameters.
	"""

	log.info("Loading application configuration ...")

	if not cfg_path.endswith((".yaml", ".yml")):
		raise ValueError("The configuration file not a YAML/YML type!")

	with open(cfg_path, encoding = "utf-8") as stream:
		content = stream.read()

	cfg = yaml.safe_load(content)
	log.info("Configuration loaded.")

	return cfg


# ====================================
# 		Fetching of user input
# ====================================

def _generate_processing_input(data: DataFrame) -> dict:
	"""Creates processing input for 
   FBL3N/FBL5N form the user data."""

	result = {}

	records = data[[
		"old_text",
		"new_text",
		"new_assignment"
	]].to_dict("records")

	for rec in records:
		result.update({
			rec["old_text"]: {
				"new_text": rec["new_text"],
				"new_assignment": rec["new_assignment"],
				"message": ""
			}
		}
	)

	return result

def _convert_user_data(att_data: bytes) -> DataFrame:
	"""Converts user-attached content and
	identifies the type of used accounts.

	Parameters:
	-----------
	att_data:
		Email attachment data as a bytes-like object.

	Returns:
	--------
	The result of conversion.
	"""

	log.info("Converting user data ...")

	data = pd.read_excel(att_data,
		dtype = {
			"account": "UInt32",
			"old_text": "string",
			"new_text": "string",
			"new_assignment": "string"
		},
		names = [
			"account",
			"old_text",
			"new_text",
			"new_assignment"
		]
	)

	# the user accidentally sends unfilled new text and assignment values
	if data["new_text"].isna().all() and data["new_assignment"].isna().all():
		raise ValueError(
			"The supplied data contains no entries in "
			"'New Text' and 'New Assignment' columns!")

	if data["old_text"].isna().all():
		raise ValueError("The supplied data contains no entries in 'Old Text' column!")

	# replace missing values with pythonic None type
	data.replace({pd.NA: None}, inplace = True)

	data = data.assign(message = pd.NA)

	return data

def get_user_input(msg_cfg: dict, email_id: str) -> dict:
	"""Retrieves the processing parameters and data provided by the user.

	If the user message is no longer available (e.g. it gets accidentally
	deleted), then a `RuntimeError` exception is raised.

	If the proessing parameters are not found in the user message or 
	the user provides incorrect values, then `ValueError` exception is raised.

	Parameters:
	-----------
	msg_cfg:
		Configuration parameters for application messages.

	email_id:
		The string ID associated with the message.

	Returns:
	--------
	A dictionary of the processing parameters and their values:
		- "error_message": `str`
			A detailed error message if an exception occurs.
		- "email": `str`
			Email address of the sender.
		- "company_code": `str`
			Company code provided by the sender.
		- "data": pandas.DataFrame
			A DataFrame containing the converted attachment data.
		- "account_type": `str`
			Type of the accounts:
				- "customer": The accounts contained in the data are customer accounts.
				- "general_ledger": The accounts contained in the data are G/L accounts.
		- "attachment": `dict`
			Details of the attachment: 
				- "name": `str`
					The name of the attachment.
				- "content": `any`
					The content of the attachment before conversion. 
	"""

	log.info("Retrieving user message ...")

	params = {
		"error_message": None,
		"email": None,
		"company_code": None,
		"data": None,
		"account_type": None,
		"attachment": {
			"name": None,
			"content": None
		}
	}

	acc = mails.get_account(
		msg_cfg["requests"]["mailbox"],
		msg_cfg["requests"]["account"],
		msg_cfg["requests"]["server"]
	)

	msg = mails.get_messages(acc, email_id)[0]

	if msg is None:
		raise RuntimeError(
			f"Could not find message with the specified ID: '{email_id}'")

	log.info("User message retrieved.")

	log.info("Extracting relevant contents from the user message ...")
	email_addr = msg.sender.email_address
	params.update({"email": email_addr})
	log.info(f"User email address: '{email_addr}'")
	cocd_patt = r"Company code:\s*(?P<cocd>\d{4})"
	cocd_match = re.search(cocd_patt, msg.text_body, re.I|re.M)

	if cocd_match is None:
		params.update({"error_message": "The message contains no valid company code!"})
		return params

	cocd = cocd_match.group("cocd")
	params.update({"company_code": cocd})
	attachments = mails.get_attachments(msg, ".xlsm")

	if len(attachments) == 0:
		params.update({"error_message": "The message contains no attachment!"})
		return params

	params["attachment"].update({
		"name": attachments[0]["name"],
		"content": attachments[0]["content"]
	})

	log.info("The contents successfully extracted.")

	log.info("Converting attachment data ...")
	converted = _convert_user_data(attachments[0]["content"])
	params.update({"data": converted})
	log.info("Data conversion completed.")

	log.info("Validating data ...")
	accs = converted["account"].unique()
	acc_lens = {len(str(acc)) for acc in accs}

	if len(acc_lens) > 1:
		params.update({"error_message": "Cannot combine customer and GL accounts in data!"})
		return params

	if acc_lens == {7}:
		acc_type = ACCOUNT_TYPE_CUSTOMER
	elif acc_lens == {8}:
		acc_type = ACCOUNT_TYPE_GENERAL_LEDGER
	else:
		acc_type = None
		params.update({"error_message": "Cannot combine customer and GL accounts in data!"})

	params.update({"account_type": acc_type})
	log.info("Data validation completed.")

	return params


# ====================================
# 		Management of SAP connection
# ====================================

def connect_to_sap(system: str) -> CDispatch:
	"""Establishes a connection to the SAP GUI scripting engine.

	Parameters:
	-----------
	system:
		The SAP system to connect to using the scripting engine.

	Returns:
	--------
	An SAP `GuiSession` object (wrapped in the `win32:CDispatch class`)
	representing the active SAP GUI session.
	"""

	log.info("Connecting to SAP ...")
	sess = sap.connect(system)
	log.info("Connection created.")

	return sess

def disconnect_from_sap(sess: CDispatch) -> None:
	"""Closes the connection to the SAP GUI scripting engine.

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object (wrapped in the `win32:CDispatch class`)
		representing the active SAP GUI session.
	"""

	log.info("Disconnecting from SAP ...")
	sap.disconnect(sess)
	log.info("Connection to SAP closed.")


# ====================================
# 			Item processing
# ====================================

def modify_accounting_parameters(
		sess: CDispatch,
		data: DataFrame,
		acc_type: str,
		cocd: str, 
		data_cfg: dict
	) -> dict:
	"""Modifies accounting item parameters in FBL3N/FBL5N based on 
	user-supplied data. Currently, only the 'Text' and 'Assignment' fields
	can be modified. 

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object (wrapped in the `win32:CDispatch` class)
		representing an active user SAP GUI session.

	data:
		The message attachment data in a pandas DataFrame.

	acc_type:
		The type of accounts being processed (e.g., "customer" or "general_ledger").

	cocd:
		The company code associated with the accounts being processed.

	data_cfg:
		Configuration parameters for handling the data.

	Returns:
	--------
	A dictionary containing the following keys and values:
		- “data”: `pandas.DataFrame`
			The original user data and the processing status stored in its 'message' field.
		- “error_message”: `str`, `None`
			  An error message if an exception occurs, otherwise None.
	"""

	if data.empty:
		raise ValueError("Argument 'data' has no records!")

	if not (len(cocd) == 4 and cocd.isnumeric()):
		raise ValueError(f"Incorrect company code: {cocd}")

	accs = data["account"].unique()
	parameters = _generate_processing_input(data)

	if acc_type == ACCOUNT_TYPE_GENERAL_LEDGER:

		fbl3n.start(sess)

		try:
			log.info("Modifying FBL3N item params ...")
			result = fbl3n.change_document_parameters(
				list(accs), cocd, parameters,
				layout = data_cfg["fbl3n_layout"])
		except fbl3n.NoItemsFoundError:
			err_msg = "No items with the text values you supplied " \
						"were found on the account (s)!"
			return {"data": data, "error_message": err_msg}
		finally:
			fbl3n.close()

	elif acc_type == ACCOUNT_TYPE_CUSTOMER:

		fbl5n.start(sess)

		try:
			log.info("Modifying FBL5N item params ...")
			result = fbl5n.change_document_parameters(
				list(accs), cocd, parameters,
				layout = data_cfg["fbl5n_layout"])
		except fbl5n.NoItemsFoundError:
			err_msg = "No items with the text values you supplied " \
						"were found on the account (s)!"
			return {"data": data, "error_message": err_msg}
		finally:
			fbl5n.close()
	else:
		raise ValueError(f"Unrecognized account type: {acc_type}")

	log.info("Item modification completed.")

	for old_txt in result:
		idx = data[data["old_text"] == old_txt].index
		data.loc[idx, "message"] = result[old_txt]["message"]

	return {"data": data, "error_message": None}


# ====================================
# 	Reporting of processing output
# ====================================

def create_report(temp_dir: str, data_cfg: dict, data: DataFrame) -> str:
	"""Generates a user report based on the processing outcome.

	Parameters:
	-----------
	temp_dir:
		The path to the directory where temporary files are stored.

	data_cfg:
		Configuration parameters for handling the data.

	data:
		The data containing the processing outcome 
		from which the report will be generated.

	Returns:
	--------
	Path to the report file.
	"""

	log.info("Creating user report ...")
	rep_path = join(temp_dir, data_cfg["report_name"])
	report.generate_excel_report(rep_path, data, data_cfg["sheet_name"])
	log.info("Report successfully created.")

	return rep_path

def send_notification(
		msg_cfg: dict,
		user_mail: str,
		template_dir: str,
		attachment: Union[dict, str] = None, # type: ignore
		error_msg: str = ""
	) -> None:
	"""Sends a notification with the processing result to the user.

	Parameters:
	-----------
	msg_cfg:
		Configuration parameters for application messages.

	user_mail:
		The email address of the user who requested processing.

	template_dir:
		The path to the directory containing notification templates.

	attachment:
		Either a dictionary containing attachment name  
		and data, or a file path to the attachment.

	error_msg:
		An error message to include in the notification.
		Default is an empty string (no error message).
	"""

	log.info("Sending notification to user ...")

	notif_cfg = msg_cfg["notifications"]

	if not notif_cfg["send"]:
		log.warning(
			"Sending of notifications to users "
			"is disabled in 'appconfig.yaml'.")
		return

	if error_msg != "":
		templ_name = "template_error.html"
	else:
		templ_name = "template_completed.html"

	templ_path = join(template_dir, templ_name)

	with open(templ_path, encoding = "utf-8") as stream:
		html_body = stream.read()

	if error_msg != "":
		html_body = html_body.replace("$error_msg$", error_msg)

	if attachment is None:
		msg = mails.create_smtp_message(
			notif_cfg["sender"], user_mail,
			notif_cfg["subject"], html_body
		)
	elif isinstance(attachment, dict):
		msg = mails.create_smtp_message(
			notif_cfg["sender"], user_mail,
			notif_cfg["subject"], html_body,
			{attachment["name"]: attachment["content"]}
		)
	elif isinstance(attachment, str):
		msg = mails.create_smtp_message(
			notif_cfg["sender"], user_mail,
			notif_cfg["subject"], html_body,
			attachment
		)
	else:
		raise ValueError(f"Unsupported data type: '{type(attachment)}'!")

	try:
		mails.send_smtp_message(msg, notif_cfg["host"], notif_cfg["port"])
	except Exception as exc:
		log.error(exc)
		return

	log.info("Notification sent.")


# ====================================
# 			Data cleanup
# ====================================

def delete_temp_files(temp_dir: str) -> None:
	"""Removes all temporary files from the specified directory.

	Parameters:
	-----------
	temp_dir:
		The path to the directory where temporary files are stored.
	"""

	file_paths = glob(join(temp_dir, "*.*"))

	if len(file_paths) == 0:
		return

	log.info("Removing temporary files ...")

	for file_path in file_paths:
		try:
			os.remove(file_path)
		except Exception as exc:
			log.exception(exc)

	log.info("Files successfully removed.")
