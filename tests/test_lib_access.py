#!/usr/bin/python3
# -*- coding: utf-8 -*-
#
""" Fountain to Open document text (.odt) translator
    tests setup & common """ 
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


from pathlib import Path
import sys
import argparse

try:
    from screenwriting_tests_common import add_to_path
except ModuleNotFoundError:
    self_path=Path(__file__).parent
    sys.path.append(str(self_path))
    from screenwriting_tests_common import add_to_path
add_to_path()
from odf_fountain_lib import ArgOptions, coalesce, to_points, attributes_to_str

def test_attributes_to_str():
    dict={'a':'orses', 'b':'for mutton'}
    ans=attributes_to_str('<attr',dict,'/>')
    assert ans=='<attr a="orses" b="for mutton"/>'
    ans=attributes_to_str('<attr',dict)
    assert ans=='<attr a="orses" b="for mutton"'

def test_inches_to_points():
    pt = to_points('1in')
    assert pt == 72

def test_coalesce_returns_correct_element():
    result = coalesce(None, None, 'This', 'Not This')
    assert result == 'This'

def test_coalesce_returns_None_for_empty_args():
    result = coalesce()
    assert result == None

def test_coalesce_returns_None_for_all_None():
    result = coalesce(None, None, None)
    assert result == None
 
def test_arg_parse():
    options_parser=ArgOptions("Description")
    options_parser.add_argument('prog', type=Path, help="")
    options_parser.add_argument('files', nargs='+', type=Path, help="input files space separated")
    options_parser.add_argument('--test', default='No')
    options=options_parser.parse_args(['prog', 'file',])
    assert str(options.prog)=='prog'
    assert options.files==[Path('file')]
    assert options.test == 'No'
 
def test_arg_parse_config(tmp_path:Path):
    options_parser=ArgOptions("Description")
    options_parser.add_argument('prog', type=Path, help="")
    options_parser.add_argument('files', nargs='+', type=Path, help="input files space separated")
    options_parser.add_argument('--test', default='No')
    options_parser.add_argument('--config', type=Path,
                     help="Configuration file which will be merged with command-line options. See documentation for format.")
    config = tmp_path / 'config.txt'
    with open(config, 'w') as f:
        f.write('test=Yes\n')
    options=options_parser.parse_args(['prog', 'file','--config',str(config)])
    assert str(options.prog)=='prog'
    assert options.files==[Path('file')]
    assert options.test == 'Yes'

def test_arg_parse_magic_config(tmp_path:Path):
    options_parser=ArgOptions("Description", config=True)
    options_parser.add_argument('prog', type=Path, help="")
    options_parser.add_argument('files', nargs='+', type=Path, help="input files space separated")
    options_parser.add_argument('--test', default='No')
    config = tmp_path / 'config.txt'
    with open(config, 'w') as f:
        f.write('test=Yes\n')
    options=options_parser.parse_args(['prog', 'file','--config',str(config)])
    assert str(options.prog)=='prog'
    assert options.files==[Path('file')]
    assert options.test == 'Yes'
# test_arg_parse()
# test_arg_parse_config(Path('/tmp'))