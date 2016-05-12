#!/usr/bin/env python
# -*- coding: latin-1 -*-
"""
Extracts emails and attachments saved in Microsoft Outlook's .msg files
    https://github.com/mattgwwalker/msg-extractor"""

__author__ = "Matthew Walker"
__date__ = "2013-11-19"
__version__ = '0.2'

# --- LICENSE -----------------------------------------------------------------
#
#    Copyright 2013 Matthew Walker
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.

import argparse
import email.utils
import glob
import json
import olefile as OleFile
import os
import random
import string
import sys
import tempfile
import traceback

from email.parser import Parser as EmailParser

# This property information was sourced from
# http://www.fileformat.info/format/outlookmsg/index.htm
# on 2013-07-22.
properties = {
    '001A': 'Message class',
    '0037': 'Subject',
    '003D': 'Subject prefix',
    '0040': 'Received by name',
    '0042': 'Sent repr name',
    '0044': 'Rcvd repr name',
    '004D': 'Org author name',
    '0050': 'Reply rcipnt names',
    '005A': 'Org sender name',
    '0064': 'Sent repr adrtype',
    '0065': 'Sent repr email',
    '0070': 'Topic',
    '0075': 'Rcvd by adrtype',
    '0076': 'Rcvd by email',
    '0077': 'Repr adrtype',
    '0078': 'Repr email',
    '007d': 'Message header',
    '0C1A': 'Sender name',
    '0C1E': 'Sender adr type',
    '0C1F': 'Sender email',
    '0E02': 'Display BCC',
    '0E03': 'Display CC',
    '0E04': 'Display To',
    '0E1D': 'Subject (normalized)',
    '0E28': 'Recvd account1 (uncertain)',
    '0E29': 'Recvd account2 (uncertain)',
    '1000': 'Message body',
    '1008': 'RTF sync body tag',
    '1035': 'Message ID (uncertain)',
    '1046': 'Sender email (uncertain)',
    '3001': 'Display name',
    '3002': 'Address type',
    '3003': 'Email address',
    '39FE': '7-bit email (uncertain)',
    '39FF': '7-bit display name',

    # Attachments (37xx)
    '3701': 'Attachment data',
    '3703': 'Attachment extension',
    '3704': 'Attachment short filename',
    '3707': 'Attachment long filename',
    '370E': 'Attachment mime tag',
    '3712': 'Attachment ID (uncertain)',

    # Address book (3Axx):
    '3A00': 'Account',
    '3A02': 'Callback phone no',
    '3A05': 'Generation',
    '3A06': 'Given name',
    '3A08': 'Business phone',
    '3A09': 'Home phone',
    '3A0A': 'Initials',
    '3A0B': 'Keyword',
    '3A0C': 'Language',
    '3A0D': 'Location',
    '3A11': 'Surname',
    '3A15': 'Postal address',
    '3A16': 'Company name',
    '3A17': 'Title',
    '3A18': 'Department',
    '3A19': 'Office location',
    '3A1A': 'Primary phone',
    '3A1B': 'Business phone 2',
    '3A1C': 'Mobile phone',
    '3A1D': 'Radio phone no',
    '3A1E': 'Car phone no',
    '3A1F': 'Other phone',
    '3A20': 'Transmit dispname',
    '3A21': 'Pager',
    '3A22': 'User certificate',
    '3A23': 'Primary Fax',
    '3A24': 'Business Fax',
    '3A25': 'Home Fax',
    '3A26': 'Country',
    '3A27': 'Locality',
    '3A28': 'State/Province',
    '3A29': 'Street address',
    '3A2A': 'Postal Code',
    '3A2B': 'Post Office Box',
    '3A2C': 'Telex',
    '3A2D': 'ISDN',
    '3A2E': 'Assistant phone',
    '3A2F': 'Home phone 2',
    '3A30': 'Assistant',
    '3A44': 'Middle name',
    '3A45': 'Dispname prefix',
    '3A46': 'Profession',
    '3A48': 'Spouse name',
    '3A4B': 'TTYTTD radio phone',
    '3A4C': 'FTP site',
    '3A4E': 'Manager name',
    '3A4F': 'Nickname',
    '3A51': 'Business homepage',
    '3A57': 'Company main phone',
    '3A58': 'Childrens names',
    '3A59': 'Home City',
    '3A5A': 'Home Country',
    '3A5B': 'Home Postal Code',
    '3A5C': 'Home State/Provnce',
    '3A5D': 'Home Street',
    '3A5F': 'Other adr City',
    '3A60': 'Other adr Country',
    '3A61': 'Other adr PostCode',
    '3A62': 'Other adr Province',
    '3A63': 'Other adr Street',
    '3A64': 'Other adr PO box',

    '3FF7': 'Server (uncertain)',
    '3FF8': 'Creator1 (uncertain)',
    '3FFA': 'Creator2 (uncertain)',
    '3FFC': 'To email (uncertain)',
    '403D': 'To adrtype (uncertain)',
    '403E': 'To email (uncertain)',
    '5FF6': 'To (uncertain)'}


def windowsUnicode(string):
    if string is None:
        return None
    if sys.version_info[0] >= 3:  # Python 3
        return str(string, 'utf_16_le')
    else:  # Python 2
        return unicode(string, 'utf_16_le')


def createNumDirIfNotExists(baseDir):

    for count in xrange(0, 100):
        suffix = '-' + str(count).zfill(2) if count else ''
        newDirName = baseDir + suffix
        if not os.path.isdir(newDirName):
            os.makedirs(newDirName)
            return newDirName

    msg = ("Failed to create a directory based on '%s'. ",
           "Does it already exist?") % baseDir
    raise FileIO(msg)


class Attachment(object):
    def __init__(self, msg, dir_):
        # Get long filename
        self.longFilename = msg._getStringStream([dir_, '__substg1.0_3707'])

        # Get short filename
        self.shortFilename = msg._getStringStream([dir_, '__substg1.0_3704'])

        # Get attachment data
        self.data = msg._getStream([dir_, '__substg1.0_37010102'])

    def save(self, newDirName):
        # Use long filename as first preference
        filename = self.longFilename
        # Otherwise use the short filename
        if filename is None:
            filename = self.shortFilename
        # Otherwise just make something up!
        if filename is None:
            handle, path = tempfile.mkstemp('.bin', 'UnknownFilename-',
                                            text=False, dir=newDirName)
            f = os.fdopen(handle, 'wb')
        else:
            path = os.path.join(newDirName, filename)
            f = open(path, 'wb')
        f.write(self.data)
        f.close()
        return path


class Message(OleFile.OleFileIO):
    def __init__(self, filename):
        OleFile.OleFileIO.__init__(self, filename)

    def _getStream(self, filename):
        if self.exists(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            return None

    def _getStringStream(self, filename, prefer='unicode'):
        """Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then the parameter /prefer/ specifies which will be
        returned.
        """

        if isinstance(filename, list):
            # Join with slashes to make it easier to append the type
            filename = "/".join(filename)

        asciiVersion = self._getStream(filename + '001E')
        unicodeVersion = windowsUnicode(self._getStream(filename + '001F'))
        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion

    @property
    def subject(self):
        return self._getStringStream('__substg1.0_0037')

    @property
    def header(self):
        try:
            return self._header
        except Exception:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText is not None:
                self._header = EmailParser().parsestr(headerText)
            else:
                self._header = None
            return self._header

    @property
    def date(self):
        # Get the message's header and extract the date
        if self.header is None:
            return None
        else:
            return self.header['date']

    @property
    def parsedDate(self):
        return email.utils.parsedate(self.date)

    @property
    def sender(self):
        try:
            return self._sender
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["from"]
                if headerResult is not None:
                    self._sender = headerResult
                    return headerResult

            # Extract from other fields
            text = self._getStringStream('__substg1.0_0C1A')
            email = self._getStringStream('__substg1.0_0C1F')
            result = None
            if text is None:
                result = email
            else:
                result = text
                if email is not None:
                    result = result + " <" + email + ">"

            self._sender = result
            return result

    @property
    def to(self):
        try:
            return self._to
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["to"]
                if headerResult is not None:
                    self._to = headerResult
                    return headerResult

            # Extract from other fields
            # TODO: This should really extract data from the recip folders,
            # but how do you know which is to/cc/bcc?
            display = self._getStringStream('__substg1.0_0E04')
            self._to = display
            return display

    @property
    def cc(self):
        try:
            return self._cc
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["cc"]
                if headerResult is not None:
                    self._cc = headerResult
                    return headerResult

            # Extract from other fields
            # TODO: This should really extract data from the recip folders,
            # but how do you know which is to/cc/bcc?
            display = self._getStringStream('__substg1.0_0E03')
            self._cc = display
            return display

    @property
    def body(self):
        # Get the message body
        return self._getStringStream('__substg1.0_1000')

    @property
    def attachments(self):
        try:
            return self._attachments
        except Exception:
            # Get the attachments
            attachmentDirs = []

            for dir_ in self.listdir():
                if dir_[0].startswith('__attach') and dir_[0] not in attachmentDirs:
                    attachmentDirs.append(dir_[0])

            self._attachments = []

            for attachmentDir in attachmentDirs:
                self._attachments.append(Attachment(self, attachmentDir))

            return self._attachments

    def createBaseDir(self, destDir, useFileName, filename=None):

        if useFileName:
            assert filename is not None
            dirName = os.path.splitext(os.path.basename(filename))[0]
        else:
            # Create a directory based on the date and subject of the message
            d = self.parsedDate
            if d is None:
                dirName = "UnknownDate"
            else:
                assert isinstance(d, tuple) and len(d) >= 5
                # year, month, day, hour, min = dateTup
                dirName = '{0:02d}-{1:02d}-{2:02d}_{3:02d}{4:02d}'.format(*d)

            subject = self.subject
            subject = subject if subject else "[No subject]"

            dirName = dirName + '-' + subject

        dirName = "".join(c for c in dirName if not
                       (ord(c) < 32 or c in r'\/:*?"<>|'))
        dirName = os.path.join(destDir, dirName.replace(' ', '-'))

        return createNumDirIfNotExists(dirName)

    def save(self, newDirName, toJson=False):

        assert os.path.isdir(newDirName)

        try:
            # Save the message body
            fext = 'json' if toJson else 'text'
            filePath = os.path.join(newDirName, "message." + fext)
            f = open(filePath, "wb")
            # From, to , cc, subject, date

            def xstr(s):
                return '' if s is None else str(s)

            attachmentNames = []
            # Save the attachments
            for attachment in self.attachments:
                attachName = os.path.basename(attachment.save(newDirName))
                attachmentNames.append(attachName)

            if toJson:
                from imapclient.imapclient import decode_utf7

                emailObj = {'from': xstr(self.sender),
                            'to': xstr(self.to),
                            'cc': xstr(self.cc),
                            'subject': xstr(self.subject),
                            'date': xstr(self.date),
                            'attachments': attachmentNames,
                            'body': decode_utf7(self.body)}

                f.write(json.dumps(emailObj, ensure_ascii=True))
            else:
                f.write("From: " + xstr(self.sender) + "\n")
                f.write("To: " + xstr(self.to) + "\n")
                f.write("CC: " + xstr(self.cc) + "\n")
                f.write("Subject: " + xstr(self.subject) + "\n")
                f.write("Date: " + xstr(self.date) + "\n")
                f.write("-----------------\n\n")
                f.write(self.body)

            f.close()

        except Exception, e:
            self.saveRaw(newDirName)
            raise


    def saveRaw(self, newNumDir):
        # Create a 'raw' folder
        rawDir = os.path.join(newNumDir, 'raw')
        if not os.path.isdir(rawDir):
            os.makedirs(rawDir)

        # Loop through all the directories
        for dir_ in self.listdir():
            sysdir = "/".join(dir_)
            code = dir_[-1][-8:-4]
            global properties
            if code in properties:
                sysdir = sysdir + " - " + properties[code]

            sysdirPath = os.path.join(rawDir, sysdir)
            if not os.path.isDir(sysdirPath):
                os.makedirs(sysdirPath)

            # Generate appropriate filename
            if dir_[-1].endswith("001E"):
                filename = "contents.txt"
            else:
                filename = "contents"

            filePath = os.path.join(sysdirPath, filename)

            # Save contents of directory
            f = open(filePath, 'wb')
            f.write(self._getStream(dir_))
            f.close()

    def dump(self):
        # Prints out a summary of the message
        print('Message')
        print('Subject:', self.subject)
        print('Date:', self.date)
        print('Body:')
        print(self.body)

    def debug(self):
        for dir_ in self.listdir():
            if dir_[-1].endswith('001E'):  # FIXME: Check for unicode 001F too
                print("Directory: " + str(dir))
                print("Contents: " + self._getStream(dir))


if __name__ == "__main__":

    epilogue = """
Launched from command line, this script parses Microsoft Outlook Message files
and saves their contents to the current directory. On error the script will
write out a 'raw' directory with all the details from the file, but in a
less-than-desirable format. To force this mode, use the '--raw' flag."""

    parser = argparse.ArgumentParser(
            formatter_class=argparse.RawDescriptionHelpFormatter,
             description =__doc__,
             epilog=epilogue)

    parser.add_argument('--raw',
                        dest='writeRaw_',
                        action='store_true',
                        help="See main description below.")

    parser.add_argument('--json',
                        dest='toJson_',
                        action='store_true',
                        help="Save the message body as a JSON object.")

    parser.add_argument('--dest-dir',
                        dest='destDir_',
                        metavar='DEST_DIR',
                        # default= os.path.abspath(os.path.dirname(__file__)),
                        default=os.getcwd(),
                        help='The directory (which must exist) where the message details will be saved. Defaults to "%(default)s" (which should be your current directory.)')

    parser.add_argument('--use-file-name',
                        dest='useFileName_',
                        action='store_true',
                        help="Email contents are normally saved in a sub-directory of DEST_DIR, named after the email's date and subject. This switch creates a sub-directory based on the name of the input file (excluding extension.)")

    parser.add_argument('msg_paths',
                        metavar='msg-path',
                        nargs='+',
                        help='Path to an Outlook message file.')

    args = parser.parse_args()

    for path in args.msg_paths:
        for filename_ in glob.glob(path):
            msg = Message(filename_)
            try:
                newDirName = msg.createBaseDir(args.destDir_, args.useFileName_,
                                               filename_)
                if args.writeRaw_:
                    msg.saveRaw(newDirName)
                else:
                    msg.save(newDirName, args.toJson_)
                print 'Message extracted to:', newDirName
            except Exception:
                # msg.debug()
                print("Error with file '" + filename_ + "': " +
                      traceback.format_exc())