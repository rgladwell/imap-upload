import codecs
import email
import email.header
import getpass
import imaplib
import locale
import mailbox
import math
import optparse
import re
import socket
import sys
import time
import unicodedata
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
                        help="destination port number [default: 143, 993 for SSL]")
        self.add_option("--ssl", action="store_true", 
                        help="use SSL connection")
        self.add_option("--box", 
                        help="destination mail box name [default: %default]")
        self.add_option("--user", help="login name [default: empty]")
        self.add_option("--password", help="login password")
        self.add_option("--retry", type="int", metavar="COUNT", 
                        help="retry COUNT times on connection abort. "
                             "0 disables [default: %default]")
        self.add_option("--error", metavar="ERR_MBOX", 
                        help="append failured messages to the file ERR_MBOX")
        self.set_defaults(host="localhost",
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
        if options.port is None:
            options.port = [143, 993][options.ssl]
        options.src = args[0]
        return options

    def parse_dest(self, dest):
        try:
            dest, ssl = re.subn("^imaps:", "imap:", dest)
            dest = urlparse(dest)
            options = optparse.Values()
            options.ssl = bool(ssl)
            options.host = dest.hostname
            options.port = [143, 993][options.ssl]
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


def si_prefix(n, prefixes=("", "k", "M", "G", "T", "P", "E", "Z", "Y"), 
              block=1024, threshold=1):
    """Get SI prefix and reduced number."""
    if (n < block * threshold or len(prefixes) == 1):
        return (n, prefixes[0])
    return si_prefix(n / block, prefixes[1:])


def str_width(s):
    """Get string width."""
    w = 0
    for c in unicode(s):
        w += 1 + (unicodedata.east_asian_width(c) in "FWA")
    return w


def trim_width(s, width):
    """Get truncated string with specified width."""
    trimed = []
    for c in unicode(s):
        width -= str_width(c)
        if width <= 0:
            break
        trimed.append(c)
    return "".join(trimed)


def left_fit_width(s, width, fill=' '):
    """Make a string fixed width by padding or truncating.

    Note: fill can't be full width character.
    """
    s = trim_width(s, width)
    s += fill * (width - str_width(s))
    return s


class Progress():
    """Store and output progress information."""

    def __init__(self, total_count):
        self.total_count = total_count
        self.ok_count = 0
        self.count = 0
        self.format = "%" + str(len(str(total_count))) + "d/" + \
                      str(total_count) + " %5.1f %-2s  %s  "

    def begin(self, msg):
        """Called when start proccessing of a new message."""
        self.time_began = time.time()
        size, prefix = si_prefix(float(len(msg.as_string())), threshold=0.8)
        sbj = self.decode_subject(msg["subject"] or "")
        print >>sys.stderr, self.format % \
              (self.count + 1, size, prefix + "B", left_fit_width(sbj, 30)),

    def decode_subject(self, sbj):
        decoded = []
        try:
            parts = email.header.decode_header(sbj)
            for s, codec in parts:
                decoded.append(s.decode(codec or "ascii"))
        except Exception, e:
            pass
        return "".join(decoded)

    def endOk(self):
        """Called when a message was processed successfully."""
        self.count += 1
        self.ok_count += 1
        print >>sys.stderr, "OK (%d sec)" % \
              math.ceil(time.time() - self.time_began)

    def endNg(self, err):
        """Called when an error has occurred while processing a message."""
        print >>sys.stderr, "NG (%s)" % err

    def endAll(self):
        """Called when all message was processed."""
        print >>sys.stderr, "Done. (OK: %d, NG: %d)" % \
              (self.ok_count, self.total_count - self.ok_count)


def upload(imap, src, err):
    print >>sys.stderr, \
          "Counting the mailbox (it could take a while for the large one)."
    p = Progress(len(src))
    for i, msg in src.iteritems():
        try:
            p.begin(msg)
            r, r2 = imap.upload(msg.get_delivery_time(), msg.as_string(), 3)
            if r != "OK":
                raise Exception(r2[0]) # FIXME: Should use custom class
            p.endOk()
            continue
        except InvalidDeliveryTime, e:
            p.endNg("Invalid delivery time: " + str(e))
        except socket.error, e:
            p.endNg("Socket error: " + str(e))
        except Exception, e:
            p.endNg(e)
        if err is not None:
            err.add(msg)
    p.endAll()


class InvalidDeliveryTime(Exception):
    """The delivery time in the From_ line is malformatted."""

def get_delivery_time(self):
    """Extract delivery time from the From_ line. 
    
    Directly attach to the mailbox.mboxMessage as a method 
    because the factory parameter of mailbox.mbox() seems
    not to work in Python 2.5.4.
    """
    try:
        time_str = self.get_from().split(" ", 1)[1]
        t = time_str.replace(",", " ").lower()
        t = re.sub(" (sun|mon|tue|wed|thu|fri|sat) ", " ", 
                   " " + t + " ")
        if t.find(":") == -1:
            t += " 00:00:00"
        t = email.utils.parsedate_tz(t)
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
        # Setup locale
        # Set LC_TIME to "C" so that imaplib.Time2Internaldate() 
        # uses English month name.
        locale.setlocale(locale.LC_ALL, "")
        locale.setlocale(locale.LC_TIME, "C")
        #  Encoding of the sys.stderr
        enc = locale.getlocale()[1] or "utf_8"
        sys.stderr = codecs.lookup(enc)[-1](sys.stderr, errors="ignore")

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
