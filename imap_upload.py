import email
import getpass
import imaplib
import mailbox
import optparse
import re
import socket
import sys
import time
import urllib
from optparse import OptionParser
from urlparse import urlparse

__version__ = "1.0"

if sys.version_info < (2, 5):
    print >>sys.stderr, "IMAP Upload requires Python 2.5 or later."
    sys.exit(1)

class MyOptionParser(OptionParser):
    def __init__(self):
        usage = "usage: python %prog [options] MBOX [DEST]\n"\
                "  MBOX UNIX style mbox file.\n"\
                "  DEST is imap[s]://[USER[:PASSWORD]@]HOST[:PORT][/BOX]\n"\
                "  DEST has a priority over the options."
        OptionParser.__init__(self, usage,
                              version="IMAP Upload " + __version__)
        self.add_option("--gmail", action="callback", nargs=0, 
                        callback=self.enable_gmail, 
                        help="setup for Gmail. Equivalents to "
                             "--host=imap.gmail.com --port=993 "
                             "--ssl --retry=3")
        self.add_option("--host", 
                        help="destination hostname [default: %default]")
        self.add_option("--port", type="int", 
                        help="destination port number [default: %default]")
        self.add_option("--ssl", action="store_true", 
                        help="use SSL connection")
        self.add_option("--box", 
                        help="destination mail box name [default: %default]")
        self.add_option("--user", help="login name [default: %default]")
        self.add_option("--password", help="login password")
        self.add_option("--retry", type="int", metavar="COUNT", 
                        help="retry COUNT times on connection abort. "
                             "0 disables [default: %default]")
        self.add_option("--error", metavar="ERR_MBOX", 
                        help="append failured messages to the file ERR_MBOX")
        self.set_defaults(host="localhost",
                          port=143,
                          ssl=False,
                          user="",
                          password="",
                          box="INBOX", 
                          retry=0,
                          error=None)
    def enable_gmail(self, option, opt_str, value, parser):
        parser.values.ssl = True
        parser.values.host = "imap.gmail.com"
        parser.values.port = 993
        parser.values.retry = 3
        
    def parse_args(self, args):
        (options, args) = OptionParser.parse_args(self, args)
        if len(args) < 1:
            self.error("Missing MBOX")
        if len(args) > 2:
            self.error("Extra argugment")
        if len(args) > 1:
            dest = self.parse_dest(args[1])
            for (k, v) in dest.__dict__.iteritems():
                setattr(options, k, v)
        options.src = args[0]
        return options

    def parse_dest(self, dest):
        try:
            dest, ssl = re.subn("^imaps:", "imap:", dest)
            dest = urlparse(dest)
            options = optparse.Values()
            options.ssl = bool(ssl)
            options.host = dest.hostname
            if dest.port:
                options.port = dest.port
            if dest.username:
                options.user = urllib.unquote(dest.username)
            if dest.password:
                options.password = urllib.unquote(dest.password)
            if len(dest.path):
                options.box = dest.path[1:] # trim the first `/'
            return options
        except:
            self.error("Invalid DEST")

    def error(self, msg):
        raise optparse.OptParseError(self.get_usage() + "\n" + msg)


def upload(imap, src, err):
    print >>sys.stderr, \
          "Counting the mailbox (it could take a while for the large one)."
    total_count = len(src)
    ok_count = 0
    for i, msg in src.iteritems():
        print >>sys.stderr, str(i + 1) + "/" + str(total_count), 
        try:
            delivery_time = msg.get_delivery_time()
            r, r2 = imap.upload(delivery_time, msg.as_string(), 3)
            if r != "OK":
                raise Exception(r2[0]) # FIXME: Should use custom class
            ok_count += 1
            print >>sys.stderr, "OK"
            continue
        except InvalidDeliveryTime, e:
            print >>sys.stderr, "NG: Invalid delivery time: ", e
        except socket.error, e:
            print >>sys.stderr, "NG: Socket error: ", e
        except Exception, e:
            print >>sys.stderr, "NG:", e
        if err is not None:
            err.add(msg)
        
    print >>sys.stderr, "Done. (OK: %d, NG: %d)" % (ok_count, total_count - ok_count)


class InvalidDeliveryTime(Exception):
    """The delivery time in the From_ line is malformatted."""

def get_delivery_time(self):
    """Extract delivery time from the From_ line. 
    
    Directly attach to the mailbox.mboxMessage as a method 
    because the factory parameter of mailbox.mbox() seems
    not to work in Python 2.5.4.
    """
    try:
        time_str = re.sub(r'^[^ ]* (.{24}).*', r'\1', self.get_from())
        try:
            t = time.strptime(time_str, "%a %b %d %H:%M:%S %Y")
            t = time.mktime(t)
            return t
        except:
            t = email.utils.parsedate_tz(time_str)
            t = email.utils.mktime_tz(t)
            return t
    except:
        raise InvalidDeliveryTime(time_str)

mailbox.mboxMessage.get_delivery_time = get_delivery_time


class IMAPUploader:
    def __init__(self, host, port, ssl, box, user, password, retry):
        self.imap = None
        self.host = host
        self.port = port
        self.ssl = ssl
        self.box = box
        self.user = user
        self.password = password
        self.retry = retry

    def upload(self, delivery_time, message, retry = None):
        if retry is None:
            retry = self.retry
        try:
            self.open()
            return self.imap.append(self.box, [], delivery_time, message)
        except (imaplib.IMAP4.abort, socket.error):
            self.close()
            if retry == 0:
                raise
        print >>sys.stderr, "(Reconnect)",
        time.sleep(5)
        return self.upload(delivery_time, message, retry - 1)

    def open(self):
        if self.imap:
            return
        imap_class = [imaplib.IMAP4, imaplib.IMAP4_SSL][self.ssl];
        self.imap = imap_class(self.host, self.port)
        self.imap.socket().settimeout(60)
        self.imap.login(self.user, self.password)

    def close(self):
        if not self.imap:
            return
        self.imap.shutdown()
        self.imap = None

def main(args=None):
    try:
        # Parse arguments
        if args is None:
            args = sys.argv[1:]
        parser = MyOptionParser()
        options = parser.parse_args(args)
        if len(str(options.user)) == 0:
            print "User name: ",
            options.user = sys.stdin.readline().rstrip("\n")
        if len(str(options.password)) == 0:
            options.password = getpass.getpass()
        options = options.__dict__
        src = options.pop("src")
        err = options.pop("error")
        # Connect to the server and login
        print >>sys.stderr, \
              "Connecting to %s:%s." % (options["host"], options["port"])
        uploader = IMAPUploader(**options)
        uploader.open()
        # Prepare source and error mbox
        src = mailbox.mbox(src, create=False)
        if err:
            err = mailbox.mbox(err)
        # Upload
        print >>sys.stderr, "Uploading..."
        upload(uploader, src, err)
        return 0
    except optparse.OptParseError, e:
        print >>sys.stderr, e
        return 2
    except mailbox.NoSuchMailboxError, e:
        print >>sys.stderr, "No such mailbox:", e
        return 1
    except socket.timeout, e:
        print >>sys.stderr, "Timed out"
        return 1
    except imaplib.IMAP4.error, e:
        print >>sys.stderr, "IMAP4 error:", e
        return 1
    except KeyboardInterrupt, e:
        print >>sys.stderr, "Interrupted"
        return 130
    except Exception, e:
        print >>sys.stderr, "An unknown error has occurred: ", e
        return 1


if __name__ == "__main__":
    sys.exit(main())
