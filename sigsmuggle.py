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

import click

PATH = '~/Library/Group Containers/UBF8T346G9.Office/Outlook/' \
       'Outlook 15 Profiles/%s/Data/'


class Signature(object):
    def __init__(self, recordid, path):
        profile = click.get_current_context().find_object(Profile)
        # this is an auto-increment so only used for exporting
        self.recordid = recordid
        self.path = self.path = urllib.parse.unquote(path)
        with open(profile._path(self.path), 'rb') as fp:
            self.contents = fp.read()
        self.OwnedBlocks = []
        self.__getBlocks()

    def __getBlocks(self):
        profile = click.get_current_context().find_object(Profile)
        cursor = profile.conn.cursor()
        # Not storing the intermediatery Signature_OwnedBlocks table info
        cursor.execute('SELECT b.* FROM Blocks b LEFT JOIN'
                       ' Signatures_OwnedBlocks s ON s.BlockId = b.BlockId'
                       ' WHERE s.Record_RecordID = ?', (self.recordid,))
        for result in cursor:
            self.OwnedBlocks.append(Block(*result))
        cursor.close()

    def __str__(self):
        string = '<Signature %d: %s' % (self.recordid, self.path)
        if (self.OwnedBlocks):
            string += ' Blocks: [%s]' % ', '.join(map(str, self.OwnedBlocks))
        string + '>'
        return string

    def write(self, cursor):
        profile = click.get_current_context().find_object(Profile)
        cursor.execute('INSERT OR REPLACE INTO Signatures (Record_RecordID,'
                       ' PathToDataFile) VALUES  (?, ?)',
                       (self.recordid, urllib.parse.quote(self.path)))
        for block in self.OwnedBlocks:
            cursor.execute('INSERT OR REPLACE INTO Signatures_OwnedBlocks'
                           ' (Record_RecordID, BlockID, BlockTag) VALUES'
                           ' (?, ?, ?)',
                           (self.recordid, block.blockid, block.tag))
            block.write(cursor)
        with open(profile._path(self.path), 'wb') as fp:
            fp.write(self.contents)


class Block(object):
    def __init__(self, blockid, tag, path):
        profile = click.get_current_context().find_object(Profile)
        self.blockid = blockid
        self.tag = tag
        self.path = urllib.parse.unquote(path)
        with open(profile._path(self.path), 'rb') as fp:
            self.contents = fp.read()

    def __str__(self):
        return '<Block %s [%s]: %s>' % (self.blockid, self.tag, self.path)

    def write(self, cursor):
        profile = click.get_current_context().find_object(Profile)
        cursor.execute('INSERT OR REPLACE INTO Blocks (BlockId, BlockTag,'
                       ' PathToDataFile) VALUES (?, ?, ?)',
                       (self.blockid, self.tag, urllib.parse.quote(self.path)))
        with open(profile._path(self.path), 'wb') as fp:
            fp.write(self.contents)


class Profile(object):
    def __init__(self, profile):
        self.profile = profile
        self.conn = sqlite3.connect(self._path('Outlook.sqlite'))
        self.conn.row_factory = sqlite3.Row
        self.signatures = []

    def _path(self, file):
        return os.path.expanduser(PATH % self.profile + file)

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
