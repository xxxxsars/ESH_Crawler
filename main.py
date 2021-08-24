from absl import app
from absl import flags
import crawler as cw

# set the cmd-line tool parameters.
FLAGS = flags.FLAGS
flags.DEFINE_string('config', "setting.ini", 'the configuration path.')
flags.DEFINE_string('action', "crawler", 'running specific action [crawler|alert].')


class ArgumentException(Exception): pass

#todo add debug mode (crawler will open browser)
def main(argv):
    with cw.Crawler(FLAGS.config) as crawler:
        action = FLAGS.action
        if action == "crawler":
            crawler.run_crawler()
        elif action == "alert":
            crawler.send_alert_mail()
        else:
            raise ArgumentException("Pleas verify your 'action' value ,the action only allow 'crawler' or 'alert'.")


if __name__ == '__main__':
    app.run(main)
