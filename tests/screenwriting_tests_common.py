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

def add_to_path():
    """Access screenwriting shared library. If we can't, assume none of the system is on
        the path, expect to find them in sub directories of our parent directory."""
    # like Pooh I know there must be a better way but can't think what it might be
    try:
        from odf_fountain_lib import to_points, coalesce
    except ModuleNotFoundError:
        base_path=Path(__file__).parent.parent

        for subdir in ['lib', 'fountain2odf', 'odf2fountain']:
            target_path=base_path / subdir
            if target_path.is_dir():
                sys.path.append(str(target_path))
        #if this still fails we have a problem
        from odf_fountain_lib import to_points, coalesce
