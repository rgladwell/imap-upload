#!/usr/bin/python3
# coding=utf-8
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
import urllib.request, urllib.parse, urllib.error
import os
import traceback
from optparse import OptionParser
from urllib.parse import urlparse
from imapclient import imap_utf7

__version__ = "2.0.0"

if sys.version_info < (3, 5):
    print("IMAP Upload requires Python 3.5 or later.")
    sys.exit(1)

class MyOptionParser(OptionParser):
    def __init__(self):
        usage = "usage: python %prog [options] (MBOX|-r MBOX_FOLDER) [DEST]\n"\
                "  MBOX UNIX style mbox file.\n"\
                "  MBOX_FOLDER folder containing subfolder trees of mbox files\n"\
                "  DEST is imap[s]://[USER[:PASSWORD]@]HOST[:PORT][/BOX]\n"\
                "  DEST has a priority over the options."
        OptionParser.__init__(self, usage,
                              version="IMAP Upload " + __version__)
        self.add_option("-r", action="store_true",
                        help="recursively search sub-folders")
        self.add_option("--gmail", action="callback", nargs=0,
                        callback=self.enable_gmail,
                        help="setup for Gmail. Equivalents to "
                             "--host=imap.gmail.com --port=993 "
                             "--ssl --retry=3")
        self.add_option("--office365", action="callback", nargs=0,
                        callback=self.enable_office365,
                        help="setup for Office365. Equivalents to "
                             "--host=outlook.office365.com --port=993 "
                             "--ssl --retry=3")
        self.add_option("--fastmail", action="callback", nargs=0,
                        callback=self.enable_fastmail,
                        help="setup for Fastmail hosted IMAP. Equivalent to "
                             "--host=imap.fastmail.com --port=993 "
                             "--ssl --retry=3")
        self.add_option("--email-only-folders", action="store_true",
                        help="use for servers that do not allow storing emails and subfolders in the same folder"
                            "only works with -r")
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
        self.add_option("--time-fields", metavar="LIST", type="string", nargs=1,
                        action="callback", callback=self.set_time_fields,
                        help="try to get delivery time of message from "
                             "the fields in the LIST. "
                             'Specify any of "from", "received" and '
                             '"date" separated with comma in order of '
                             'priority (e.g. "date,received"). '
                             '"from" is From_ line of mbox format. '
                             '"received" is "Received:" field and "date" '
                             'is "Date:" field in RFC 2822. '
                             '[default: from,received,date]')
        self.add_option("--list_boxes", action="store_true",
                        help="list all mail boxes in the IMAP server")
        self.add_option("--folder-separator", type="string",
                        help="change folder separator-character default")
        self.add_option("--google-takeout", action="store_true",
                        help="Import Google Takeout using labels as folders.")
        self.add_option("--google-takeout-box-as-base-folder", action="store_true",
                        help="Use given box as base folder.")
        self.add_option("--google-takeout-first-label", action="store_true",
                        help="Only import first label from the email.")
        self.add_option("--google-takeout-rename-label-ampersand", action="store_true",
                        help="Rename ampersand in labels")
        self.add_option("--debug", action="store_true",
                        help="Debug: Make some error messages more verbose.")
        self.set_defaults(host="localhost",
                          ssl=False,
                          r=False,
                          email_only_folders=False,
                          user="",
                          password="",
                          box="INBOX",
                          retry=0,
                          error=None,
                          time_fields=["from", "received", "date"],
                          folder_separator="/",
                          google_takeout=False,
                          google_takeout_box_as_base_folder=False,
                          google_takeout_first_label=False,
                          google_takeout_rename_label_ampersand=False,
                          debug=False
                          )

    def enable_gmail(self, option, opt_str, value, parser):
        parser.values.ssl = True
        parser.values.host = "imap.gmail.com"
        parser.values.port = 993
        parser.values.retry = 3

    def enable_office365(self, option, opt_str, value, parser):
        parser.values.ssl = True
        parser.values.host = "outlook.office365.com"
        parser.values.port = 993
        parser.values.retry = 3

    def enable_fastmail(self, option, opt_str, value, parser):
        parser.values.ssl = True
        parser.values.host = "imap.fastmail.com"
        parser.values.port = 993
        parser.values.retry = 3

    def set_time_fields(self, option, opt_str, value, parser):
        fields = []
        if value != "":
            fields = value.split(",")
        # Assert that list contains only valid fields
        if set(fields) - set(["from", "received", "date"]):
            self.error("Invalid value '%s' for --time-fields" % value)
        self.values.time_fields = fields

    def parse_args(self, args):
        (options, args) = OptionParser.parse_args(self, args)
        if len(args) < 1 and not options.list_boxes:
            self.error("Missing MBOX")
        if len(args) > 2:
            self.error("Extra argument")
        if len(args) > 1:
            dest = self.parse_dest(args[1])
            for (k, v) in dest.__dict__.items():
                setattr(options, k, v)
        if options.port is None:
            options.port = [143, 993][options.ssl]
        if not options.list_boxes:
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
                options.user = urllib.parse.unquote(dest.username)
            if dest.password:
                options.password = urllib.parse.unquote(dest.password)
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
    for c in str(s):
        w += 1 + (unicodedata.east_asian_width(c) in "FWA")
    return w


def trim_width(s, width):
    """Get truncated string with specified width."""
    trimed = []
    for c in str(s):
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


def decode_header_to_string(header):
    """Decodes an email message header (possibly RFC2047-encoded)
    into a string, while working around https://bugs.python.org/issue22833"""

    return "".join(
        alleged_string if isinstance(alleged_string, str) else alleged_string.decode(
            alleged_charset or 'ascii')
        for alleged_string, alleged_charset in email.header.decode_header(header))


class Progress():
    """Store and output progress information."""

    def __init__(self, total_count, google_takeout = False):
        self.total_count = total_count
        self.ok_count = 0
        self.count = 0
        self.format = "%" + str(len(str(total_count))) + "d/" + \
                      str(total_count) + " %5.1f %-2s  %s  "
        self.google_takeout = google_takeout

    def begin(self, msg):
        """Called when start proccessing of a new message."""
        self.time_began = time.time()
        size, prefix = si_prefix(float(len(msg.as_string())), threshold=0.8)
        sbj = decode_header_to_string(msg["subject"] or "")
        google_takeout_language = "es"
        if self.google_takeout:
            if (google_takeout_language == "es"):
                gmail_inbox_str = r"Recibidos"
                gmail_sent_str = r"Enviados"
                gmail_draft_str = "Borradores"
                gmail_important_str = u'Importante'
                gmail_open_str = u'Abierto'
                gmail_unseen_str = u"No leídos"
                gmail_category_str = r"^Categor.a:"
                gmail_imap_str = r'^IMAP_'
                gmail_trash_str = "Papelera"
            else:
                gmail_inbox_str = r"Safata d'entrada"
                gmail_sent_str = r"Enviats"
                gmail_draft_str = "Esborranys"
                gmail_important_str = u'Importants'
                gmail_open_str = u'Oberts'
                gmail_unseen_str = u"No llegits"
                gmail_category_str = r"^Categor.a"
                gmail_imap_str = r'^IMAP_'
                gmail_trash_str = "Paperera"
            label = decode_header_to_string(msg["x-gmail-labels"] or "")
            sanitized_label = re.sub(r"\n\r", "", label)
            sanitized_label = re.sub(r"\r\n", "", sanitized_label)
            sanitized_label = re.sub(r"\r", " ", sanitized_label)
            sanitized_label = re.sub(r"\n", "", sanitized_label)
            label = sanitized_label
            label = re.sub(gmail_inbox_str, "INBOX", label)
            label = re.sub(gmail_sent_str, "Sent", label)
            labels = label.split(",")

            labels_without_categories = []
            for i in range(len(labels)):
                if (not (re.match(gmail_category_str,labels[i]))):
                    labels_without_categories.append(labels[i])

            labels = labels_without_categories

            labels_without_special_imap_dirs = []
            for i in range(len(labels)):
                if (not (re.match(gmail_imap_str,labels[i]))):
                    labels_without_special_imap_dirs.append(labels[i])

            labels = labels_without_special_imap_dirs

            sanitized_labels = []
            for i in range(len(labels)):
                sanitized_label = re.sub(r":", "_", labels[i])
                sanitized_labels.append(sanitized_label)
            labels = sanitized_labels

            if labels.count(gmail_open_str) > 0:
                labels.remove(gmail_open_str)

            if labels.count(u'INBOX') > 0:
                labels.remove(u'INBOX')

            flags = []
            if labels.count(gmail_unseen_str) > 0:
                labels.remove(gmail_unseen_str)
            else:
                flags.append('\Seen')

            if labels.count(gmail_important_str) > 0:
                flags.append('\Flagged')
                labels.remove(gmail_important_str)

            if ((labels.count(gmail_sent_str) > 0) and (len(labels) > 1)):
                labels.remove(gmail_sent_str)

            if labels.count(gmail_trash_str) > 0:
                labels.remove(gmail_trash_str)
                labels.append('Trash')

            if len(labels):
                msg.flags = " ".join(flags)
            else:
                msg.flags = []

            msg.boxes = []
            if len(labels) != 0:
                if labels.count(gmail_draft_str):
                    msg.boxes.append(['Drafts'])
                else:
                    if labels.count('Spam'):
                        msg.boxes.append(['Junk'])
                    else:
                        for i in range(len(labels)):
                            box = re.sub(r"\?", "", labels[i])
                            msg.boxes.append(box.split("/"))
            if len(msg.boxes) == 0:
                msg.boxes.append(["INBOX"])

        print(self.format % \
              (self.count + 1, size, prefix + "B", left_fit_width(sbj, 30)), end=' ')

    def endOk(self):
        """Called when a message was processed successfully."""
        self.count += 1
        self.ok_count += 1
        print("OK (%d sec)" % \
              math.ceil(time.time() - self.time_began))

    def endNg(self, err):
        """Called when an error has occurred while processing a message."""
        print("NG (%s)" % err)

    def endAll(self):
        """Called when all message was processed."""
        print("Done. (OK: %d, NG: %d)" % \
              (self.ok_count, self.total_count - self.ok_count))


def upload(imap, box, src, err, time_fields, google_takeout = False, debug = False):
    print("Uploading to {}...".format(box))
    print("Counting the mailbox (it could take a while for the large one).")
    p = Progress(len(src), google_takeout=google_takeout)
    for i, msg in src.items():
        try:
            p.begin(msg)
            if google_takeout:
                for i in range(len(msg.boxes)):
                    r, r2 = imap.upload(box, msg.get_delivery_time(time_fields),
                                        msg.as_string(), msg.flags, msg.boxes[i], 3)
                    if r != "OK":
                        raise Exception(r2[0]) # FIXME: Should use custom class
            else:
                r, r2 = imap.upload(box, msg.get_delivery_time(time_fields),
                                    msg.as_string(), None, None, 3)
                if r != "OK":
                    raise Exception(r2[0]) # FIXME: Should use custom class

            p.endOk()
            continue
        except socket.error as e:
            p.endNg("Socket error: " + str(e))
        except Exception as e:
            if debug:
                p.endNg(traceback.format_exc())
            else:
                p.endNg(e)
        if err is not None:
            err.add(msg)
    p.endAll()


def recursive_upload(imap, box, src, err, time_fields, email_only_folders, separator):
    usrc = str(src)
    for file in os.listdir(usrc):
        path = usrc + os.sep + file
        if os.path.isdir(path):
            fileName, fileExtension = os.path.splitext(file)
            if not box:
                subbox = fileName
            else:
                subbox = box + separator + fileName
            recursive_upload(imap, subbox, path, err, time_fields, email_only_folders, separator)
        elif file.endswith("mbox"):
            print("Found mailbox at {}...".format(path))
            mbox = mailbox.mbox(path, create=False)
            if (email_only_folders and has_mixed_content(src)):
                target_box = box + separator + src.split(os.sep)[-1]
            else:
                target_box = file.split('.')[0] if (box is None or box == "") else box
            if err:
                err = mailbox.mbox(err)
            upload(imap, target_box, mbox, err, time_fields)

def has_mixed_content(src):
    dirFound = False
    mboxFound = False

    for file in os.listdir(src):
        path = src + os.sep + file
        if (os.path.isdir(path)):
            dirFound = True
        elif file.endswith("mbox"):
            mboxFound = True

    return dirFound and mboxFound

def pretty_print_mailboxes(boxes):
    for box in boxes:
        x = re.search("\(((\\\\[A-Za-z]+\s*)+)\) \"(.*)\" \"?(.*)\"?",box)
        if not x:
            print("Could not parse: {}".format(box))
            continue
        raw_name = x.group(4)
        sep = x.group(3)
        raw_flags = x.group(1)
        print("{:40s}{}".format(pretty_mailboxes_name(raw_name, sep), pretty_flags(raw_flags)))

def pretty_mailboxes_name(name, sep):
    depth = name.count(sep)
    spacer = "  "
    branch = "+- " if (depth>0) else ""
    slash = name.rfind(sep)
    clean_name = name if (slash == -1) else name[slash+1:]
    return "{0}{1}\"{2}\"".format( spacer*depth, branch, clean_name)

def pretty_flags(raw_flags):
    flags = raw_flags.replace("\\HasChildren", "")
    flags = flags.replace("\\HasNoChildren", "")
    flags = flags.replace("\\", "#")
    flags = flags.split()
    return "\t".join(flags)

def get_delivery_time(self, fields):
    """Extract delivery time from message.

    Try to extract the time data from given fields of message.
    The fields is a list and can consist of any of the following:
      * "from"      From_ line of mbox format.
      * "received"  The first "Received:" field in RFC 2822.
      * "date"      "Date:" field in RFC 2822.
    Return the current time if the fields is empty or no field
    had valid value.
    """
    def get_from_time(self):
        """Extract the time from From_ line."""
        time_str = self.get_from().split(" ", 1)[1]
        t = time_str.replace(",", " ").lower()
        t = re.sub(" (sun|mon|tue|wed|thu|fri|sat) ", " ",
                   " " + t + " ")
        if t.find(":") == -1:
            t += " 00:00:00"
        return t
    def get_received_time(self):
        """Extract the time from the first "Received:" field."""
        t = self["received"]
        t = t.split(";", 1)[1]
        t = t.lstrip()
        return t
    def get_date_time(self):
        """Extract the time from "Date:" field."""
        return self["date"]

    for field in fields:
        try:
            t = vars()["get_" + field + "_time"](self)
            t = email.utils.parsedate_tz(t)
            t = email.utils.mktime_tz(t)
            # Do not allow the time before 1970-01-01 because
            # some IMAP server (i.e. Gmail) ignore it, and
            # some MUA (Outlook Express?) set From_ date to
            # 1965-01-01 for all messages.
            if t < 0:
                continue
            return t
        except:
            pass
    # All failed. Return current time.
    return time.time()

# Directly attach get_delivery_time() to the mailbox.mboxMessage
# as a method.
# I want to use the factory parameter of mailbox.mbox()
# but it seems not to work in Python 2.5.4.
mailbox.mboxMessage.get_delivery_time = get_delivery_time


class IMAPUploader:
    def __init__(self, host, port, ssl, box, user, password, retry):
        self.imap = None
        self.host = host
        self.port = port
        self.ssl = ssl
        self.user = user
        self.password = password
        self.retry = retry
        self.box = box

    def upload(self, box, delivery_time, message, flags = None, boxes = None, retry = None):
        if retry is None:
            retry = self.retry
        if flags is None:
            flags = []
        try:
            self.open()
            if boxes is not None: # Google Takeout
                if type(message) == str:
                    message = message.encode('utf-8', 'surrogateescape').decode('utf-8')
                    message = bytes(message, 'utf-8')
                try:
                    self.create_folders(boxes)
                    google_takeout_box = "/".join(boxes)
                    google_takeout_box_imap_command = '"' + google_takeout_box + '"'
                    res = self.imap.append(imap_utf7.encode(google_takeout_box_imap_command), flags, delivery_time, message)
                except:
                    google_takeout_box = "/".join(boxes)
                    google_takeout_box_imap_command = '"' + google_takeout_box + '"'
                    res = self.imap.append(imap_utf7.encode(google_takeout_box_imap_command), flags, delivery_time, message)
                return res
            else: # Default behaviour
                self.imap.create(box)
                if type(message) == str:
                    message = bytes(message, 'utf-8')
                return self.imap.append(box, flags, delivery_time, message)
        except (imaplib.IMAP4.abort, socket.error):
            self.close()
            if retry == 0:
                raise
        print("(Reconnect)", end=' ')
        time.sleep(5)
        return self.upload(box, delivery_time, message, flags, boxes, retry - 1)

    def create_folders(self, boxes):
        i = 1
        while i <= len(boxes):
            google_takeout_box = "/".join(boxes[0:i])
            google_takeout_box_imap_command = '"' + google_takeout_box + '"'
            if google_takeout_box != "INBOX":
                try:
                    # self.imap.enable("UTF8=ACCEPT")
                    self.imap.create(imap_utf7.encode(google_takeout_box_imap_command))
                except:
                    print ("Cannot create box %s" % google_takeout_box)
            i += 1

    def open(self):
        if self.imap:
            return
        imap_class = [imaplib.IMAP4, imaplib.IMAP4_SSL][self.ssl]
        self.imap = imap_class(self.host, self.port)
        self.imap.socket().settimeout(60)
        self.imap.login(self.user, self.password)

        try:
            self.imap.create(self.box)
        except Exception as e:
            print("(create error: )" + str(e))

    def close(self):
        if not self.imap:
            return
        self.imap.shutdown()
        self.imap = None

    def list_boxes(self):
        try:
            self.open()
            status, mailboxes = self.imap.list()
            return mailboxes
        except (imaplib.IMAP4.abort, socket.error):
            self.close()


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
            print("User name: ", end=' ', flush=True)
            options.user = sys.stdin.readline().rstrip("\n")
        if len(str(options.password)) == 0:
            options.password = getpass.getpass()
        options = options.__dict__
        list_boxes = options.pop("list_boxes")
        err = options.pop("error")
        time_fields = options.pop("time_fields")

        recurse = options.pop("r")
        email_only_folders = options.pop("email_only_folders")
        separator = options.pop("folder_separator")
        google_takeout = options.pop("google_takeout")
        google_takeout_box_as_base_folder = options.pop("google_takeout_box_as_base_folder")
        google_takeout_first_label = options.pop("google_takeout_first_label")
        google_takeout_rename_label_ampersand = options.pop("google_takeout_rename_label_ampersand")
        debug = options.pop("debug")


        # Connect to the server and login
        print("Connecting to %s:%s." % (options["host"], options["port"]))

        if (list_boxes):
            print("Just list mail boxes!")

            uploader = IMAPUploader(**options)
            uploader.open()
            pretty_print_mailboxes(uploader.list_boxes())
        else:
            src = options.pop("src")

            uploader = IMAPUploader(**options)
            uploader.open()

            if(not recurse):
                # Prepare source and error mbox
                src = mailbox.mbox(src, create=False)
                if err:
                    err = mailbox.mbox(err)
                upload(uploader, options["box"], src, err, time_fields, google_takeout, debug)
            else:
                recursive_upload(uploader, "", src, err, time_fields, email_only_folders, separator)

        return 0

    except optparse.OptParseError as e:
        print(e)
        return 2
    except mailbox.NoSuchMailboxError as e:
        print("No such mailbox:", e)
        return 1
    except socket.timeout as e:
        print("Timed out")
        return 1
    except imaplib.IMAP4.error as e:
        print("IMAP4 error:", e)
        return 1
    except KeyboardInterrupt as e:
        print("Interrupted")
        return 130
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print("An unknown error has occurred [{}]: ".format(exc_tb.tb_lineno), e)
        return 1


if __name__ == "__main__":
    print("IMAP Upload (v{})".format(__version__))
    sys.exit(main())
