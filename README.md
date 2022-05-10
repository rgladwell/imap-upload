# IMAP Upload

_Copyright &copy; 2009-2014 [OZAWA Masayuki](http://coroq.com/) and [Ricardo Gladwell](http://gladwell.me)_

IMAP Upload is a tool for uploading a local mbox file to IMAP4 server. The most stable way to migrate to Gmail.

### Features

*   Recursively import mbox sub-folders, currently supports Mac Mail MBOX export folder format.
*   Read messages stored in mbox format which is used by many mail clients such as Thunderbird.
*   Upload messages to IMAP4 server.
*   Preserve the delivery time of the message. (support date time in From_ line / &ldquo;Received:&rdquo; field / &ldquo;Date:&rdquo; field)
*   Automatic retry when the connection was aborted which happens frequently on Gmail.
*   Can write out failed messages in mbox format. (Easy to retry for the failed messages)
*   Supports IMAP servers that can only store either folders or emails in a folder
*   Support SSL.
*   Run on Windows, Mac OS X, Linux, *BSD, and so on.
*   Command line interface. (No friendly GUI, sorry...)
*   Free of charge.
*   Open source.

### Requirements

*   Python 3.5 or later.
*   imapclient ( `pip3 install imapclient` )

### Quick Start

Uploading a local mail box file &ldquo;Friends.mbox&rdquo; to the remote mail box &ldquo;imported&rdquo; on the server &ldquo;example.com&rdquo; using SSL:

```sh
python imap_upload.py Friends.mbox imaps://example.com/imported
```

You can specify the destination by options instead of URL:

```sh
python imap_upload.py --host example.com --port 993 --ssl --box imported Friends.mbox
```

You can use a shortcut option for the Gmail server:

```sh
python imap_upload.py --gmail --box imported Friends.mbox
```

You can use a shortcut option for the Office 365 server:

```sh
python imap_upload.py --office365 --box imported Friends.mbox
```

There's an `--error` option so that you can store the failed messages in mbox format and retry for them later:

```sh
python imap_upload.py --gmail --box imported --error Friends.err Friends.mbox
```

You can also recursively import mbox sub-folders using th `-r` option:

```
python imap_upload.py --gmail -r path
```

If your server only supports email or folders per folder you can use the `--email-only-folders` option together with `-r`.
If a mixed content folder is found, the emails of the folder are uploaded to a sub-folder of the same name:

```sh
python imap_upload.py -r path --email-only-folders
```

Example:
```
**Local**
  Foo (Folder)
   -> Bar (Folder)
   -> Email 1
   -> Email 2

**Remote**
  Foo
    -> Bar (Folder)
    -> Foo (Folder)
      -> Email 1
      -> Email 2
```

You can use just output the account's mailboxes (folders/labels) list. This is useful if you need to upload to an existing special mailbox (i.e.: Gmail's Send Email label, when using a language different from English):

```sh
python imap_upload.py --gmail --list_boxes
```
If you prefer a tree-like view of the mailboxes:
```sh
python imap_upload.py --gmail --list_boxes --treeview
```

Some email providers use alternative IMAP folder separators (for example, Hetzner uses the `.` separator character). You can change this default using the `--folder-separator` argument, as follows:

```sh
python imap_upload.py -r path --folder-separator '.' --email-only-folders
```

Google Takeout example (might duplicates mails):
```sh
python imap_upload.py --ssl --user=login@example.net --password=MyS3cr3t --host=mail.example.net --port=993 --error='All mail Including Spam and Trash_errors.mbox' --google-takeout 'All mail Including Spam and Trash.mbox'
```

Google Takeout example using only one label per mail:
```sh
python imap_upload.py --ssl --user=login@example.net --password=MyS3cr3t --host=mail.example.net --port=993 --error='All mail Including Spam and Trash_errors.mbox' --google-takeout --google-take-out-one-label 'All mail Including Spam and Trash.mbox'
```



For more details, please refer to the --help message:

```sh
python imap_upload.py --help
```

### Usage

```
IMAP Upload (v2.0.0)
Usage: python imap_upload.py [options] (MBOX|-r MBOX_FOLDER) [DEST]
  MBOX UNIX style mbox file.
  MBOX_FOLDER folder containing subfolder trees of mbox files
  DEST is imap[s]://[USER[:PASSWORD]@]HOST[:PORT][/BOX]
  DEST has a priority over the options.

Options:
  --version             show program's version number and exit
  -h, --help            show this help message and exit
  -r                    recursively search sub-folders
  --gmail               setup for Gmail. Equivalents to --host=imap.gmail.com
                        --port=993 --ssl --retry=3
  --office365           setup for Office365. Equivalents to
                        --host=outlook.office365.com --port=993 --ssl
                        --retry=3
  --fastmail            setup for Fastmail hosted IMAP. Equivalent to
                        --host=imap.fastmail.com --port=993 --ssl --retry=3
  --email-only-folders  use for servers that do not allow storing emails and
                        subfolders in the same folderonly works with -r
  --host=HOST           destination hostname [default: localhost]
  --port=PORT           destination port number [default: 143, 993 for SSL]
  --ssl                 use SSL connection
  --box=BOX             destination mail box name [default: INBOX]
  --user=USER           login name [default: empty]
  --password=PASSWORD   login password
  --retry=COUNT         retry COUNT times on connection abort. 0 disables
                        [default: 0]
  --error=ERR_MBOX      append failured messages to the file ERR_MBOX
  --time-fields=LIST    try to get delivery time of message from the fields in
                        the LIST. Specify any of "from", "received" and "date"
                        separated with comma in order of priority (e.g.
                        "date,received"). "from" is From_ line of mbox format.
                        "received" is "Received:" field and "date" is "Date:"
                        field in RFC 2822. [default: from,received,date]
  --list_boxes          list all mail boxes in the IMAP server
  --folder-separator=FOLDER_SEPARATOR
                        change folder separator-character default
  --google-takeout      Import Google Takeout using labels as folders.
  --google-takeout-box-as-base-folder
                        Use given box as base folder.
  --google-takeout-first-label
                        Only import first label from the email.
  --google-takeout-label-priority=GOOGLE_TAKEOUT_LABEL_PRIORITY
                        Priority of labels, if --google-takeout-first-label is
                        used
  --google-takeout-language=GOOGLE_TAKEOUT_LANGUAGE
                        [Use specific language. Supported languages: 'en es ca
                        de'. default: en]
  --google-takeout-flagged-labels=GOOGLE_TAKEOUT_FLAGGED_LABELS
                        Mark Mails with given labels (comma separated) as
                        flagged, by default the Important flag is used
  --debug               Debug: Make some error messages more verbose.
  --dry-run             Do not perform IMAP writing actions
```

