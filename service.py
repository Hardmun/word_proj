# Adapted from http://pypi.python.org/pypi/watchdog
import sys, os, re
import time
import logging
from optparse import OptionParser
import subprocess
from watchdog.observers import Observer
import watchdog.events

import pythoncom
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket

help_text = """Monitor directory(s) for new files and take action.
%prog [options, -h for details]"""

# List of tuple consisting of (directory, include regexp, execute)
# "execute" may be a list in case of multiple dispatch commands
configs = (
    #("/var/tmp", '^etil', "echo %(filename)s"),
)

class eventHandler(watchdog.events.FileSystemEventHandler):
    '''File create/move event handler'''
    def __init__(self, include_regexp=None, actions=None):
        '''Filename include-match regular expression and action may be given'''
        super(eventHandler, self).__init__()
        if include_regexp:
            self.regexp = re.compile(include_regexp)
        else:
            self.regexp = None
        if not actions or hasattr(actions, '__iter__'):
            self.actions = actions
        else:
            # If action is not tuple/list, make it iterable
            self.actions = [ actions ]
    def is_matching(self, file_path):
        if self.regexp:
            return self.regexp.search( os.path.basename(file_path), flags=re.IGNORECASE )
        return True
    def do_actions(self, file_path):
        if not self.actions: return
        for action in self.actions:
            if re.search('%\(', action):
                command = action % { 'filename':file_path }
            else:
                command = action
            msg = "Executing '%s'" % command
            logging.debug(msg)
            servicemanager.LogInfoMsg(msg)
            p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            stdout, stderr = p.communicate()
            p.wait()
            if p.returncode != 0:
                msg = "returncode=%d stdout=%s stderr=%s" % (p.returncode, stdout, stderr)
                logging.error(msg)
                servicemanager.LogErrorMsg(msg)
                raise ValueError("Action failed %d: %s: stdout=%s stderr=%s" % (p.returncode, command, stdout, stderr))
    def on_moved(self, event):
        super(eventHandler, self).on_moved(event)
        if event.is_directory: return
        logging.debug("moved: %s -> %s" % (event.src_path, event.dest_path))
        if not self.is_matching(event.dest_path):
            logging.debug("Skipped moved file: %s" % event.dest_path)
            return
        try:
            self.do_actions(event.dest_path)
        except Exception as e:
            msg = "moved %s action error: %s" % (event.dest_path, e)
            logging.error(msg)
            servicemanager.LogErrorMsg(msg)
    def on_created(self, event):
        super(eventHandler, self).on_moved(event)
        if event.is_directory: return
        logging.debug("created: %s" % (event.src_path))
        if not self.is_matching(event.src_path):
            logging.debug("Skipped created file: %s" % event.src_path)
            return
        try:
            self.do_actions(event.src_path)
        except Exception as e:
            msg = "created %s action error: %s" % (event.src_path, e)
            logging.error(msg)
            servicemanager.LogErrorMsg(msg)

# http://stackoverflow.com/questions/32404/can-i-run-a-python-script-as-a-service-in-windows-how
class AppServerSvc (win32serviceutil.ServiceFramework):
    _svc_name_ = "MSG Incoming Fire"
    _svc_display_name_ = "MSG Incoming Fire"
    def __init__(self,args):
        win32serviceutil.ServiceFramework.__init__(self,args)
        self.hWaitStop = win32event.CreateEvent(None,0,0,None)
        socket.setdefaulttimeout(60)
    def SvcStop(self):
        self.run_flag = False
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_,''))
        self.run_flag = True
        self.main()
    def main(self):
        observer = Observer()
        #event_handler = watchdog.events.LoggingEventHandler()
        #observer.schedule(event_handler, path=sys.argv[1], recursive=True
        for dir, include_regexp, actions in configs:
            eh = eventHandler(include_regexp=include_regexp, actions=actions)
            observer.schedule(eh, path=dir, recursive=True)
            logging.info("Watching %s" % dir)
        observer.start()
        try:
            while True:
                if self.run_flag is False: break
                time.sleep(1)
        #except KeyboardInterrupt:
        except Exception as e:
            servicemanager.LogInfoMsg("Caught exception: %s" % e)
        finally:
            observer.stop()
        observer.join()

def main(argv=None):
    if argv is None:
        argv = sys.argv

    # http://docs.python.org/library/logging.html
    debuglevelD = {
        'debug': logging.DEBUG,
        'info': logging.INFO,
        'warning': logging.WARNING,
        'error': logging.ERROR,
        'critical': logging.CRITICAL
    }
    defvals = {}
    parser = OptionParser(usage=help_text)
    parser.add_option("-d", dest="debug", type="string", help="Logging verbosity %s"%debuglevelD.keys(), metavar='LEVEL')
    parser.add_option("-D", dest="no_daemon", action="store_true", help="Do not run as a service")
    parser.set_defaults(**defvals)
    (options, args) = parser.parse_args()
    if options.debug:
        if options.debug not in debuglevelD:
            parser.error("Logging verbosity must be one of: %s"%debuglevelD.keys())
        dbglevel = debuglevelD[options.debug]
    else:
        dbglevel = logging.WARNING
    logging.basicConfig(level=dbglevel,
                        format='%(asctime)s - %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S')
    win32serviceutil.HandleCommandLine(AppServerSvc)


if __name__ == "__main__":
    sys.exit(main())