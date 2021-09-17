import logging
import sys
import traceback

import pandas as pd
from absl import app
from absl import flags
import crawler as cw
import alarm_mail as am
import handler

logging.basicConfig(filename='esh.log', encoding='utf-8', level=logging.DEBUG, format='%(asctime)s %(message)s',
                    datefmt='%m/%d/%Y %I:%M:%S %p')

# set the cmd-line tool parameters.
FLAGS = flags.FLAGS
flags.DEFINE_boolean('debug', False, 'Turn on/off the debug mode (crawler will show the browser).')
flags.DEFINE_boolean('log', False, 'Turn logging on/off.')
flags.DEFINE_string('config', "setting.ini", 'the configuration path.')
flags.DEFINE_string('action', "crawler", 'running specific action [crawler|mail].')


class ArgumentException(Exception): pass


def main(argv):
    config_path = FLAGS.config
    log_mode = FLAGS.log

    with cw.Crawler(config_path, FLAGS.debug) as crawler:
        action = FLAGS.action
        if action == "crawler":
            try:
                # run the  crawler program.
                history_xlsx_path = crawler.run_crawler()
                # send the wrong format mail
                am.Alarm_mail(config_path, history_xlsx_path).send_wrong_format_mail()
                # upload the result file to EDA ftp
                handler.upload_files(config_path)
            except Exception as e:
                if log_mode:
                    logging.error("Crawler Error:" + traceback.format_exc())
                else:
                    print(e)

        elif action == "mail":
            try:
                am.Alarm_mail(config_path).send_overdue_mail()
            except Exception as e:
                if log_mode:
                    logging.error("Overdue Mail Error:" + traceback.format_exc())
                else:
                    print(e)
                return
        else:
            raise ArgumentException("Pleas verify your 'action' value ,the action only allow 'crawler' or 'alert'.")


if __name__ == '__main__':
    app.run(main)
