#! /usr/bin/env python3
#
# Copyright (C) 2017 Brian Hartvigsen
#
# Permission to use, copy, modify, and/or distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH
# REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY
# AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,
# INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM
# LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR
# OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR
# PERFORMANCE OF THIS SOFTWARE.

import os
import pickle
import sqlite3
import urllib.parse
import uuid

from random import randint

import click

PATH = '~/Library/Group Containers/UBF8T346G9.Office/Outlook/' \
       'Outlook 15 Profiles/%s/Data/'


class Signature(object):
    def __init__(self, recordid, path):
        fpath = urllib.parse.unquote(path)
        click.echo("Found signature @ %s" % fpath)
        with open(Profile.getPath(fpath), 'rb') as fp:
            self.contents = fp.read()
        self.OwnedBlocks = []
        self.__getBlocks(recordid)

    def __getBlocks(self, recordid):
        cursor = Profile.getCursor()
        # Not storing the intermediatery Signature_OwnedBlocks table info
        cursor.execute('SELECT b.* FROM Blocks b LEFT JOIN'
                       ' Signatures_OwnedBlocks s ON s.BlockId = b.BlockId'
                       ' WHERE s.Record_RecordID = ?', (recordid,))
        for result in cursor:
            click.echo("Signature owns block %s [%s]" % (result['BlockId'].hex(), result['BlockTag']))
            self.OwnedBlocks.append(Block(*result))
        cursor.close()

    def write(self, cursor):
        fpath = '/'.join(['Signatures', str(randint(0, 255)), "%s.olk15Signature" %
                         str.upper(uuid.uuid4().__str__())])
        click.echo("Writing Signature to %s" % fpath)
        cursor.execute('INSERT INTO Signatures (PathToDataFile) VALUES  (?)',
                       (urllib.parse.quote(fpath),))
        recordid = cursor.lastrowid
        for block in self.OwnedBlocks:
            click.echo("Signature owns block %s [%s]" % (block.id.hex(), block.tag))
            cursor.execute('INSERT INTO Signatures_OwnedBlocks'
                           ' (Record_RecordID, BlockID, BlockTag) VALUES'
                           ' (?, ?, ?)',
                           (recordid, block.id, block.tag))
            block.write(cursor)
        fpath = Profile.getPath(fpath)
        os.makedirs(os.path.dirname(fpath), exist_ok=True)
        with open(fpath, 'wb') as fp:
            fp.write(self.contents)


class Block(object):
    def __init__(self, id, tag, path):
        # This looks to always be b'15000000' but just in case
        self.header = id[0:4]
        self.tag = tag
        fpath = urllib.parse.unquote(path)
        click.echo("Block exists at %s" % fpath)
        with open(Profile.getPath(fpath), 'rb') as fp:
            self.contents = fp.read()

    def write(self, cursor):
        fpath = '/'.join(['Signature Attachments', str(randint(0, 255)),
                          "%s.olk15SigAttachment" %
                          str.upper(self.uuid.__str__())])
        click.echo("Writing Block %s to %s" % (self.id.hex(), fpath))
        cursor.execute('INSERT INTO Blocks (BlockId, BlockTag, PathToDataFile)'
                       ' VALUES (?, ?, ?)',
                       (self.id, self.tag, urllib.parse.quote(fpath)))
        fpath = Profile.getPath(fpath)
        os.makedirs(os.path.dirname(fpath), exist_ok=True)
        with open(fpath, 'wb') as fp:
            fp.write(self.contents)

    def __setstate__(self, state):
        self.__dict__.update(state)
        if self.blockid:
            self.header = self.blockid[0:4]
            self.uuid = uuid.UUID(bytes=self.blockid[4:])
        else:
            self.uuid = uuid.uuid4()
        self.id = self.header + self.uuid.bytes


class Profile(object):
    def __init__(self, profile):
        self.profile = profile
        self.conn = sqlite3.connect(self.__getPath('Outlook.sqlite'))
        self.conn.row_factory = sqlite3.Row
        self.signatures = []

    def __getPath(self, file):
        return os.path.expanduser(PATH % self.profile + file)

    @staticmethod
    def _getProfile():
        return click.get_current_context().find_object(Profile)

    @staticmethod
    def getPath(file):
        return Profile._getProfile().__getPath(file)

    @staticmethod
    def getCursor():
        return Profile._getProfile().conn.cursor()

    def readSignatures(self):
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM Signatures')
        for result in cursor:
            self.signatures.append(Signature(result["Record_RecordID"],
                                             result["PathToDataFile"]))
        cursor.close()

    def writeSignatues(self):
        cursor = self.conn.cursor()
        for sig in self.signatures:
            sig.write(cursor)
        cursor.close()
        self.conn.commit()


@click.group()
@click.option('--profile', '-p', default='Main Profile',
              help='Which Outlook profile to operate on')
@click.pass_context
def cli(ctx, profile):
    ctx.obj = Profile(profile)


@cli.command()
@click.argument('file', type=click.File(mode='wb'))
@click.pass_obj
def export(profile, file):
    profile.readSignatures()
    pickle.dump(profile.signatures, file)


@cli.command(name='import')
@click.argument('file', type=click.File(mode='rb'))
@click.pass_obj
def import_(profile, file):
    profile.signatures = pickle.load(file)
    profile.writeSignatues()


if __name__ == '__main__':
    cli()
