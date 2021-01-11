#!/usr/bin/env python

import sys
import logging
import traceback
from sys import platform
import json

EXIT_BOT = "exit bot"
BOT_ERROR = "bot error"
WINDOWS_FILE = "C:/ProgramData/AutomationAnywhere/BotRunner/Logs/python3wrapper.log"
LINUX_FILE = "/var/log/automationanywhere/python3wrapper.log"
OUTPUT_PREFIX = "#output_start#"
OUTPUT_POSTFIX = "#output_end#"


def main():
    # Setting up logging
    file_location = WINDOWS_FILE if platform == "win32" else LINUX_FILE
    FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    logging.basicConfig(filename=file_location, level=logging.DEBUG, format=FORMAT)
    bot_imported = False
    while True:
        try:
            msg = sys.stdin.readline().strip()
            logging.debug("msg: %s", msg)
            if msg == EXIT_BOT:
                complete = OUTPUT_PREFIX + "true" + OUTPUT_POSTFIX + "\n"
                sys.stdout.write(complete)
                sys.stdout.flush()
                logging.info("Exiting the run !!")
                break
            data = json.loads(msg)
            function_output = OUTPUT_PREFIX + "true" + OUTPUT_POSTFIX + "\n"
            if 'functionName' not in data:
               logging.info("importing and running bot ...")
               import bot
            else:
               if (not bot_imported):
                  logging.info("importing bot for execution..")
                  import bot
                  bot_imported = True
               result = bot.play(data)   
               function_output = OUTPUT_PREFIX + json.dumps(result, ensure_ascii=False) + OUTPUT_POSTFIX + "\n"
            sys.stdout.buffer.write(function_output.encode('utf-8'))
            sys.stdout.buffer.flush()
            logging.debug("Bot output : %s", function_output)
        except KeyboardInterrupt:
            pass
        except:
            err_output = OUTPUT_PREFIX + BOT_ERROR + OUTPUT_POSTFIX + "\n"
            sys.stdout.write(err_output)
            sys.stdout.flush()
            logging.error("Error Running bot function: %s", traceback.format_exc())
            sys.stderr.write(traceback.format_exc())
            sys.stderr.flush()

    logging.info("Bot run complete !!!")

if __name__ == "__main__":
    main()
