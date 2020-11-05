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

*   Python 2.5 or later.

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

For more details, please refer to the --help message:

```sh
python imap_upload.py --help
```
