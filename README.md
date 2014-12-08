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

There's an `--error` option so that you can store the failed messages in mbox format and retry for them later:

```sh
python imap_upload.py --gmail --box imported --error Friends.err Friends.mbox
```

You can also recursively import mbox sub-folders using th `-r` option:

```
python imap_upload.py --gmail -r path
```

For more details, please refer to the --help message:

```sh
python imap_upload.py --help
```
