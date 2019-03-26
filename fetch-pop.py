#!/bin/env python

"""
Fetch contents of POP3S mailbox and extract it's contents into
specified directory
"""

import poplib
import email
import mimetypes
import os
import argparse
from ConfigParser import ConfigParser
from UserDict import UserDict


class POPBox(object):
    def __init__(self, host, user, password):
        self.host = host
        self.user = user
        self.password = password
        self.connection = poplib.POP3_SSL(self.host)
        self.connection.user(self.user)
        self.connection.pass_(password)

    def pop_message(self, delete_message=True):
        """Fetch and delete message
            return: (msgnum, msguid, message)
        """
        for i in range(len(self.connection.list()[1])):
            message_lines = self.connection.retr(i+1)[1]
            message = email.message_from_string(b'\n'.join(message_lines))
            (response, mnum, uid) = self.connection.uidl(i+1).split()
            if delete_message:
                self.connection.dele(i+1)
            yield (i, uid, message)

    def close(self):
        self.connection.quit()


class AppConfig(UserDict):
    def __getattr__(self, attr):
        return getattr(self, 'data')[attr]

    def __str__(self):
        return str(getattr(self, 'data'))


if __name__ == '__main__':
    config = AppConfig()
    config_file = ConfigParser()
    opened_configs = config_file.read(['.fetchrc',
                                       os.path.expanduser('~/.fetchrc')])
    if opened_configs:
        for i in ('user', 'password', 'server'):
            config[i] = config_file.get('auth', i, raw=False)
        config['directory'] = config_file.get('locations', 'directory',
                                              raw=False)
    parser = argparse.ArgumentParser(
        description="Fetch contents of mailbox and unpack it's contents")
    parser.add_argument('--user', default=config.user,
                        help='User name (email)')
    parser.add_argument('--password', default=config.password, help='password')
    parser.add_argument('--server', default=config.server, help='POP3 server')
    subdirs_group = parser.add_mutually_exclusive_group()
    subdirs_group.add_argument('--subdirs', action='store_true', default=False,
                               help='create sub-directories based on messages UIDs')
    subdirs_group.add_argument('--subject', action='store_true', default=False,
                               help='create sub-directories based on messages Subject')
    parser.add_argument('--no-delete', action='store_true', default=False,
                        help='do not delete message on server')
    parser.add_argument('--directory',
                        default=config.directory,
                        help='directory to store message contents into')
    args = parser.parse_args()

    for i in ('user', 'password', 'server', 'directory'):
        config[i] = getattr(args, i, config[i])

    config.user = os.environ.get('POP3_USER', config.user)
    config.password = os.environ.get('POP3_PASSWORD', config.password)
    config.server = os.environ.get('POP3_SERVER', config.server)
    config.directory = os.environ.get('POP3_DIRECTORY', config.directory)

    mbox = POPBox(host=config.server, user=config.user,
                  password=config.password)
    counter = 1
    if not os.path.exists(config.directory):
        os.makedirs(config.directory)
    for (msgid, msguid, message) in mbox.pop_message(not args.no_delete):
        subject = message.get('Subject', '-No subject-')
        print("Processing {0}/{1}: {2}".format(msgid, msguid, subject))
        if args.subdirs:
            counter = 1
            my_dir = os.path.join(config.directory, msguid)
            if not os.path.exists(my_dir):
                os.makedirs(my_dir)
        elif args.subject:
            counter = 1
            my_dir = os.path.join(config.directory, subject)
            if not os.path.exists(my_dir):
                os.makedirs(my_dir)
        else:
            my_dir = config.directory
        for part in message.walk():

            if part.get_content_maintype() == 'multipart':
                continue
            # Applications should really sanitize the given filename so that an
            # email message can't be used to overwrite important files
            filename = part.get_filename()
            if not filename:
                ext = mimetypes.guess_extension(part.get_content_type())
                if not ext:
                    # Use a generic bag-of-bits extension
                    ext = '.bin'
                filename = 'part-%05d%s' % (counter, ext)
            counter += 1
            fp = open(os.path.join(my_dir, filename), 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()

    mbox.close()
