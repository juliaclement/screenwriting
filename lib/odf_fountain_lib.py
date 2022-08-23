#!/usr/bin/python3
# -*- coding: utf-8 -*-
"""
Open document text (.odt) to Fountain translator
Utility functions / etc
"""
#
# Copyright (C) 2022 Julia Ingleby Clement
# 
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# 
#     http://www.apache.org/licenses/LICENSE-2.0
# 
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
# Contributor(s):
#
# 10/08/2022 Julia Clement     v 0.0.0     Initial version
#

import argparse
from ast import arg
from pathlib import Path
import sys

def coalesce( *args ):
    """
    Return first non-null value
    based on non standard MsSQL function COALESCE
    """
    for i in args:
        if not i is None:
            return i
    return None

"""
Convert measurements
Libre Office supports setting measurement units to
mm, cm, inch, pica & points.
Experimentally, all measures we use in the documents I've checked
are the same throughout the document, but some documents are inches
& some are cm. It simplifies logic to work internally on a single measurement.
As points are the most fine of these measurements, we
convert the measurements in the document to points
"""
measurement_factors = {
    'pt':   1.0,
    'pc':   12.0, # not found in my sample documents
    'in':   72.0,
    'cm':   28.3465,
    'mm':   2.83465,
    'ff':   272160 # Football fields, for our US friends
}

"""
Convert string like 123cm, 12in. 0.0002645ff etc to points
"""
def to_points( value: str ) :
    uom = value[-2:]
    factor = measurement_factors.get(uom,1.0)
    return float(value.strip('cimntpf '))*factor


class ArgOptions:
    class Arg:
        def set_default(self, default):
            self.kw['default']=default

        def __init__(self, *pargs, **kw) -> None:
            self.pargs=pargs
            self.name=pargs[0].strip('-')
            self.kw=kw

    def add_argument(self, *pargs, **kw):
        arg=self.Arg(*pargs,**kw)
        self.args[arg.name]=arg

    def set_default(self, opt, default):
        try:
            arg=self.args[opt]
            arg.set_default(default)
        except KeyError:
            print(f'Unknown option {opt}, skipped')

    def create_parser(self)->argparse.ArgumentParser:
        arg_parser=argparse.ArgumentParser(self.desc)
        if self.config and self.args.get('config') == None:
            self.add_argument('--config', type=Path,
                help="Configuration file which will be merged with command-line options. See documentation for format.")
        
        for arg in self.args.values():
            arg_parser.add_argument(*arg.pargs, **arg.kw)
        return arg_parser

    def parse_args(self, args=sys.argv, parser:argparse.ArgumentParser=None)->argparse.Namespace:
        if parser is None:
            parser = self.create_parser()
        if len(args)==0:
            args=['prog', '--help'] # throws exception on parsing
        user_options=parser.parse_args(args)
        if hasattr(user_options,'config') and user_options.config:
            storeTrues=[]
            with open(user_options.config, 'r') as configFile:
                lines=configFile.readlines()
            for line in lines:
                line=line.strip('\t\r\n -')
                if line[0]=='#':
                    pass
                elif '=' in line:
                    s=line.split('=',maxsplit=1)
                    self.set_default(s[0].strip(),s[1].strip())
                elif ' ' in line:
                    s=line.split(maxsplit=1)
                    self.set_default(s[0].strip(),s[1].strip())
                else:
                    storeTrues.append(line)
            # second pass with defaults modified
            parser=self.create_parser( )
            user_options=parser.parse_args(args)
            for opt in storeTrues:
                user_options.__setattr__(opt, True)
        return user_options

    def __init__(self, desc, config=False) -> None:
        self.desc = desc
        self.config=config
        self.args={}
