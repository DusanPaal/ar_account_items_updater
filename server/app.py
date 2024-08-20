# pylint: disable = C0103, C0301, R0911, W1203, W0718

"""
The 'AR Account Items Updater' application automates updating of 'Text' field
values of items located on customer or general ledger accounts. The user sends a mail
request containing input data via the client part of the application to a specified
email address. The data must contain a list of accounts, old text values and new text
values to a specified mail address. The server part of the application then processes
the data in FBL3N or FBL5N, respectively. An excel report file is generated and sent
back to the user.

Version history:
----------------
1.0.20220902 - Initial version.
1.0.20221013 - Minor code refactoring acoss modues.
			 - Updated doctrings.
1.0.20230704 - Minor code refactoring acoss modues.
			 - Updated doctrings.
"""

from os.path import join
from datetime import datetime as dt
import argparse
import logging
import sys
from engine import controller

log = logging.getLogger("master")

def main(args: dict) -> int:
	"""Program entry point.

	Controls the overall execution
	of the program.

	Parameters (args):
	------------------
	- "email_id":
		The string ID of the user message
		that triggers the application.

	Returns:
	--------
	Program completion state:
	- 0: Program successfully completes.
	- 1: Program fails during the initialization phase.
	- 2: Program fails during the user input fetch phase.
	- 3: Program fails during the processing phase.
	- 4: Program fails during the reporting phase.
	"""

	app_dir = sys.path[0]
	log_dir = join(app_dir, "logs")
	temp_dir = join(app_dir, "temp")
	template_dir = join(app_dir, "notifications")
	app_cfg_path = join(app_dir, "app_config.yaml")
	log_cfg_path = join(app_dir, "log_config.yaml")
	curr_date = dt.now().strftime("%d-%b-%Y")

	try:
		controller.configure_logger(
			log_dir, log_cfg_path,
			"Application name: AR Account Items Updater",
			"Application version: 1.0.20230704",
			f"Log date: {curr_date}")
	except Exception as exc:
		print(exc)
		print("CRITICAL: Unhandled exception while trying to configuring the logging system!")
		return 1

	try:
		log.info("=== Initialization START ===")
		cfg = controller.load_app_config(app_cfg_path)
		sess = controller.connect_to_sap(cfg["sap"]["system"])
		log.info("=== Initialization END ===\n")
	except Exception as exc:
		log.exception(exc)
		log.critical("Unhandled exception while trying to initialize the application!")
		return 2

	try:
		log.info("=== Fetching user input START ===")
		user_input = controller.get_user_input(cfg["messages"], args["email_id"])
		log.info("=== Fetching user input END ===\n")
	except Exception as exc:
		log.exception(exc)
		log.critical("=== Fetching user input FAILURE ===")
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(sess)
		controller.delete_temp_files(temp_dir)
		log.info("=== Cleanup END ===\n")
		return 2
	  
	if user_input["error_message"] is not None:
		log.error(user_input["error_message"])
		controller.send_notification(
			cfg["messages"], user_input["email"], template_dir,
			error_msg = user_input["error_message"])
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(sess)
		controller.delete_temp_files(temp_dir)
		log.info("=== Cleanup END ===\n")
		return 2

	log.info("=== Processing START ===")
	 
	try:
		output = controller.modify_accounting_parameters(
			sess, user_input["data"],
			user_input["account_type"],
			data_cfg = cfg["data"],
			cocd = user_input["company_code"])
	except Exception as exc:
		log.exception(exc)
		log.critical("=== Processing FAILURE ===\n")
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(sess)
		controller.delete_temp_files(temp_dir)
		log.info("=== Cleanup END ===\n")
		return 3 

	if output["error_message"] is not None:
		log.warning(output["error_message"])
		controller.send_notification(
			cfg["messages"], user_input["email"],
			template_dir, user_input["attachment"],
			error_msg = output["error_message"])
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(sess)
		controller.delete_temp_files(temp_dir)
		log.info("=== Cleanup END ===\n")
		return 3 

	log.info("=== Processing END ===\n")

	try: 
		log.info("=== Reporting START ===")
		report_path = controller.create_report(temp_dir, cfg["data"], output["data"])
		controller.send_notification(
			cfg["messages"], user_input["email"],
			template_dir, report_path)
		log.info("=== Reporting END ===\n")
	except Exception as exc:
		log.exception(exc)
		return 4
	finally:
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(sess)
		controller.delete_temp_files(temp_dir)
		log.info("=== Cleanup END ===\n")

	return 0


if __name__ == "__main__":

	parser = argparse.ArgumentParser()

	parser.add_argument(
		"-e", "--email_id",
		required = True,
		help = "Sender email id."
	)

	exit_code = main(vars(parser.parse_args()))
	log.info(f"=== System shutdown with return code: {exit_code} ===")
	logging.shutdown()
	sys.exit(exit_code)
