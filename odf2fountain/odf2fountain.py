#!/usr/bin/python3
# -*- coding: utf-8 -*-
#
# Open document text (.odt) to Fountain translator
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
#  3/08/2022   Julia Clement     v 0.0.0     Initial version
# 10/08/2022   Julia Clement     v 0.0.0     Improve usability, separate lib module
# 19/08/2022   Julia Clement     v 0.0.0     Various Clean-ups.
#

from enum import IntEnum
import argparse
from pathlib import Path
import zipfile
from xml.sax import make_parser, handler
from xml.sax.xmlreader import InputSource
import sys
from odf.namespaces import TEXTNS, STYLENS
from io import StringIO
# access our shared library.
# expect to find it on the path or in either the same directory as this module or ../lib
# like Pooh I know there must be a better way but can't think what it might be
try:
    from odf_fountain_lib import to_points, coalesce
except ModuleNotFoundError:
    self_path = Path(__file__).parent
    sys.path.append(str(self_path))
    util_path = self_path.parent / 'lib'
    if util_path.is_dir():
        sys.path.append(str(util_path))
    from odf_fountain_lib import to_points, coalesce


user_options=None

class FountainType(IntEnum):
    NULL = 0
    TITLE = 1
    ACTION = 2
    # CENTERED = special case of ACTION
    CHARACTER  = 3
    DIALOGUE = 4
    # LYRIC = special case of DIALOGUE
    NOTES = 5
    PARENTHETICAL = 6
    SCENE_HEADING = 7
    TRANSITION = 8
    PAGE_BREAK = 9

class FountainRule():
    def __init__( self, name, fountain_type=None, prefix='', suffix='', blank_before=False, \
                  blank_after=False, always_required=False, require_before=[]):
        self.name: str = name
        self.fountain_type: FountainType=fountain_type
        self.prefix: str=prefix
        self.suffix: str=suffix
        self.blank_before: bool=blank_before
        self.blank_after: bool=blank_after
        self.always_required: bool=always_required
        self.require_before: list [FountainType]= require_before
    
    def __str__(self):
        answer=f" {self.fountain_type}({self.prefix}, {self.suffix})"
        if self.blank_before:
            answer += ", Blank Before"
        if self.blank_after:
            answer += ", Blank After"
        if self.require_before:
            answer += f",require before [{self.require_before}]"
        return answer

# Lookup the primary rule for a FountainType
# Order must be by FountainType without gaps or duplicates
# TODO: This order rule is far too fragile, convert to a dict 
fountain_rules_by_type = [
    FountainRule( 'Null', FountainType.NULL),   #No action, don't change type
    FountainRule( 'Title', FountainType.TITLE, 'Title:', always_required=True),
    FountainRule( 'Action', FountainType.ACTION, '!'),
    FountainRule( 'Character', FountainType.CHARACTER, '@', blank_before=True),
    FountainRule( 'Dialogue', FountainType.DIALOGUE,
                require_before=[FountainType.CHARACTER,FountainType.PARENTHETICAL, FountainType.DIALOGUE]),
    FountainRule( 'Notes', FountainType.NOTES, '[[', ']]', always_required=True),
    FountainRule( 'Parenthetical', FountainType.PARENTHETICAL, '(', ')', always_required=True,
                require_before=[FountainType.CHARACTER, FountainType.DIALOGUE]),
    FountainRule( 'Scene', FountainType.SCENE_HEADING, '.', blank_before=True, blank_after=True),
    FountainRule( 'Transition', FountainType.TRANSITION, '>', blank_before=True, blank_after=True),
]
    

fountain_rules = {  'Subtitle': FountainRule( 'Subtitle', FountainType.TITLE ),
                    'Centred':  FountainRule( 'Centred', FountainType.ACTION, '>', '<', always_required=True),
                    'Lyrics':   FountainRule( 'Lyrics', FountainType.DIALOGUE, '~', always_required=True,
                                require_before=[FountainType.CHARACTER,FountainType.PARENTHETICAL]),
                }
for rule in fountain_rules_by_type:
    fountain_rules[ rule.name ] = rule


style_to_fountain = { 
    'Title':                fountain_rules['Title'],
    'Subtitle':             fountain_rules['Subtitle'],
    'Title_20_Line':        fountain_rules['Subtitle'],
    'Script_20_Elements':   fountain_rules['Null'],
    'Standard':             fountain_rules['Null'],
    'Action':               fountain_rules['Action'],
    'Centered':             fountain_rules['Centred'],
    'Character':            fountain_rules['Character'],
    'Dialogue':             fountain_rules['Dialogue'],
    'Lyrics':               fountain_rules['Lyrics'],
    'Notes':                fountain_rules['Notes'],
    'Parenthetical':        fountain_rules['Parenthetical'],
    'Heading':              fountain_rules['Null'],
    'Scene_20_Heading':     fountain_rules['Scene'],
    'Scene':                fountain_rules['Scene'],
    'Transition':           fountain_rules['Transition']
}

class OdtStyle():
    all_styles = {}
    styles_to_parent=[]

    # inherit properties
    def maybe_inherit_from_parent(self):
        if self.parent and self.parent.heritage_loaded:
            self.heritage_loaded = True
            self.fountain_info = coalesce( self.fountain_info, self.parent.fountain_info )
            self.italic = coalesce( self.italic, self.parent.italic )
            self.bold = coalesce( self.bold, self.parent.bold )
            self.underline = coalesce( self.underline, self.parent.underline)
            self.uppercase = coalesce( self.uppercase, self.parent.uppercase )
            self.align = coalesce( self.align, self.parent.align )
            self.border_line_width = coalesce( self.border_line_width, self.parent.border_line_width )
            self.border = coalesce( self.border, self.parent.border )
            self.margin_left = coalesce( self.margin_left, self.parent.margin_left )
            self.margin_right = coalesce( self.margin_right, self.parent.margin_right )
            self.margin_top = coalesce( self.margin_top, self.parent.margin_top )
            self.margin_bottom = coalesce( self.margin_bottom, self.parent.margin_bottom )
            self.page_break = coalesce( self.page_break, self.parent.page_break )
            self.is_title = coalesce( self.is_title, self.parent.is_title )

    def assign_parent( self, parent ):
        self.parent = parent
        self.base_parent = parent
        self.base_parent_name = parent.name

    def assign_parent_by_name( self, parent_name ):
        self.parent = OdtStyle.all_styles.get(parent_name)
        if self.parent:
            self.assign_parent( self.parent )
        else:
            self.base_parent_name = parent_name

    def __init__(self, name, parent_name, display_name = None, margin_left=None, margin_right=None):
        self.name = name
        self.display_name = coalesce( display_name, name )
        self.fountain_info = style_to_fountain.get(name)
        self.heritage_loaded = False
        self.margin_left=margin_left
        self.margin_right=margin_right
        self.margin_top = None
        self.margin_bottom = None
        self.parent_name = parent_name
        self.bold = None
        self.italic = None
        self.underline = None

        if name in ('Standard','Script_20_Elements'):
            # These two styles are special and end the parent chain
            self.parent_name = name
            self.parent = self
            self.base_parent_name = name
            self.base_parent = self
            self.is_base = True
            self.heritage_loaded = True
            if not self.fountain_info:
                self.fountain_info=fountain_rules['Null']
            self.italic = False
            self.uppercase = False
            self.align='left'
            self.border_line_width = '0cm'
            self.border = '0.1pt double #000000'
            self.page_break = False
            self.is_title = False
        else:
            self.is_base = False
            self.base_parent_name=parent_name # will be adjusted later
            self.assign_parent_by_name(parent_name)
            self.uppercase = None
            self.align = None
            self.border_line_width = None
            self.border = None
            self.page_break = None
            if 'TITLE' in name.upper():
                self.is_title = True
            else:
                self.is_title = None

        if self.parent and self.parent.base_parent:
            self.base_parent = self.parent.base_parent
            self.fountain_info = coalesce(self.fountain_info, self.parent.fountain_info)
        if self.parent and self.parent.heritage_loaded:
            self.maybe_inherit_from_parent()
        else:
            self.styles_to_parent.append( self )

    def __str__(self):
        return f" {self.name}({self.parent_name}) margins {self.margin_left}, {self.margin_right}, {self.fountain_info}"

class OdtXmlStyleHandler(handler.ContentHandler):
    """ Extract headings from content.xml of an ODT file """
    def __init__(self):
        handler.ContentHandler.__init__(self)
        self.data = []
        self.current_style = None
        self.building_style = False

    def startElementNS(self, tag, qname, attrs):
        if tag == (STYLENS, 'style'):
            if user_options.debug:
                print("Start:",tag[1], "Attrs:", attrs.items())
            try:
                name = attrs._attrs[(STYLENS,'name')]
                parent_name = attrs._attrs.get((STYLENS,'parent-style-name'), 'Standard')
                display_name = attrs._attrs.get((STYLENS,'display-name'), name)
                self.current_style=OdtStyle(name, parent_name, display_name)
                self.building_style = True
            except KeyError:
                pass
        elif tag[0] == STYLENS and tag[1] in ['paragraph-properties', 'text-properties'] and self.current_style:
            if user_options.debug:
                print("Start:",tag[1], "Attrs:", attrs.items())
            for (attr_name,value) in attrs.items():
                kw=attr_name[1]
                if kw == 'margin-left':
                    self.current_style.margin_left = to_points( value )
                elif kw == 'margin-right':
                    self.current_style.margin_right = to_points( value )
                elif kw == 'margin-top':
                    self.current_style.margin_top = to_points( value )
                elif kw == 'margin-bottom':
                    self.current_style.margin_bottom = to_points( value )
                elif kw == 'text-transform':
                    if value == 'uppercase':
                        self.current_style.uppercase = True
                elif kw == 'text-align':
                    if value == 'end':
                        self.current_style.align = 'right'
                elif kw == 'font-style':
                    if value == 'italic':
                        self.current_style.italic = True
                elif kw == 'font-weight':
                    if value == 'bold':
                        self.current_style.bold = True
                elif kw == 'text-underline-style':
                    if value == 'solid':
                        self.current_style.underline = True
                elif kw == 'border':
                    self.current_style.border = value
                elif kw == 'break-before' and value=='page':
                    self.current_style.page_break = True
                elif kw == 'page-number' and value != 'auto': # implies page break
                    self.current_style.page_break = True
                elif kw == 'border-line-width':
                    self.current_style.border_line_width = value

    def endElementNS(self, tag, qname):
        if tag == (STYLENS, 'style') and self.current_style and self.building_style:
            self.building_style = False
            OdtStyle.all_styles[self.current_style.name] = self.current_style
            # Weird, but I've encountered an ODT which used the display name 
            # in the <text:p text:style-name="..."> tag.
            if '_20_' in self.current_style.name:
                OdtStyle.all_styles[self.current_style.display_name] = self.current_style
            OdtStyle.styles_to_parent.append(self.current_style)

def post_process_styles(styles_to_assign_base_parents, styles_to_assign_heritage):
    #
    # 1) Assign base parents. This process  may also assign parents which
    #    may end up assigning inherited properties.
    #    As the order of elements in a dictionary is not defined,
    #    we (Have done in every test I've done so far) need
    #    to make multiple passes to assign all parents
    next_pass=styles_to_assign_base_parents
    redo=True
    while redo:
        this_pass = next_pass
        next_pass = {}
        for style_name in this_pass:
            style=OdtStyle.all_styles[style_name]
            if not style.parent:
                style.parent = OdtStyle.all_styles.get(style.parent_name)
                if style.parent and not style.fountain_info:
                    style.fountain_info = style.parent.fountain_info
            if not style.is_base:
                base_parent=OdtStyle.all_styles.get(style.base_parent_name,None)
                if base_parent:
                    style.base_parent = base_parent.base_parent
                    style.base_parent_name = base_parent.base_parent_name
                    if not base_parent.is_base:
                        next_pass[style_name]=style
        redo=len(next_pass)<len(this_pass)
    #
    # 2) If needed, assign parents and inherit properties
    #
    #    TBH, I don't know if this will ever be needed, but just-in-case
    #    I also can't imagine multiple passes being needed, but will support anyway
    next_pass = styles_to_assign_heritage
    redo = True
    while redo:
        this_pass = next_pass
        next_pass = []
        for style in this_pass:
            if style.parent_name and not style.heritage_loaded:
                style.assign_parent_by_name(style.parent_name)
                if not style.heritage_loaded:
                    next_pass.append(style)
        redo = len(next_pass) < len(this_pass)

def load_ODT_styles(odtfile):
    style_parser = make_parser()
    style_parser.setFeature(handler.feature_namespaces, 1)
    style_parser.setContentHandler(OdtXmlStyleHandler())
    content = getZipPart(odtfile,'styles.xml')
    if not isinstance(content, str):
        content=content.decode("utf-8")
    style_parser.setErrorHandler(handler.ErrorHandler())
    inpsrc = InputSource()
    # if not isinstance(content, line):
    #     content=content.decode("utf-8")
    inpsrc.setByteStream(StringIO(content))
    style_parser.parse(inpsrc)
    post_process_styles(OdtStyle.all_styles, OdtStyle.styles_to_parent)
    if user_options.debug:
        print( "Styles:")
        for style in OdtStyle.all_styles:
            print( OdtStyle.all_styles[style] )


class DocumentState(IntEnum):
    STARTING = 2
    TITLES = 0
    BODY = 1
#
# Extract text from content.xml
# Content contains additional styles.
# To avoid duplication of code, we subclass the style handler
#
class OdtXmlTextHandler(OdtXmlStyleHandler):
    """ Extract headings from content.xml of an ODT file """
    def __init__(self, lines: list[ list[str]]):
        OdtXmlStyleHandler.__init__(self)
        self.state: DocumentState = DocumentState.STARTING
        while len(lines) < 3:
            lines.append([])
        self.r: list[ list[str]] = lines
        self.data: list [ str ] = []
        self.last_line_blank: bool = False
        self.last_rule: FountainRule = fountain_rules_by_type[0]
        self.last_character: str = ''
        self.in_text = 0

        #information to end a span
        #convert nested emphasis back to previous level when a span ends
        self.spans=['']

    def characters(self, chars: str):
        if self.in_text > 0 and len(chars)>0:
            self.data.append(chars)

    def startElementNS(self, tag, qname, attrs):
        if tag == (TEXTNS, 'p') or tag == (TEXTNS, 'h'):
            if self.in_text == 0:
                self.data = []
            self.in_text += 1
            for (attr_name,value) in attrs.items():
                if attr_name[1] == 'style-name':
                    self.current_style = OdtStyle.all_styles.get(value, self.current_style)
            if user_options.debug:
                print("Start:",tag[1], "Attrs:", attrs._attrs)
        elif tag==(TEXTNS, 'span'):
            new_style:OdtStyle=None
            for (attr_name,value) in attrs.items():
                if attr_name==(TEXTNS, 'style-name'):
                    new_style = OdtStyle.all_styles.get(value, self.current_style)
            if new_style:
                self.spans.append('')
                open_text=''
                while len(self.data) > 0 and self.data[-1] == '':
                    self.data.pop()
                if len(self.data)> 0 and self.data[-1][-1] != ' ':
                    open_text = ' '
                close_text=''
                if new_style.underline:
                    open_text+='_'
                    close_text='_'
                if new_style.bold:
                    open_text+='**'
                    close_text='**'+close_text
                if new_style.italic:
                    open_text+='*'
                    close_text='*'+close_text
                self.spans.append(close_text)
                self.data.append(open_text)
            else:
                self.spans.append['']

        elif tag[0] == STYLENS and \
                tag[1] in ['style', 'paragraph-properties', 'text-properties']:
            OdtXmlStyleHandler.startElementNS(self, tag, qname, attrs)
        elif tag == (TEXTNS, 'tab'):
            self.data.append("\t")
        elif tag == (TEXTNS, 's'):
            for (attr_name,value) in attrs.items():
                shortName=attr_name[1]
                if shortName == 'c':
                    count=int(value)
                    self.data.append( ' ' * count)

    def endElementNS(self, tag, qname):
        if tag[0] == STYLENS:
            if tag[1] == 'style' and self.current_style and self.building_style:
                OdtXmlStyleHandler.endElementNS(self, tag, qname)
                post_process_styles({self.current_style.name:self.current_style }, [self.current_style])
            else:
                OdtXmlStyleHandler.endElementNS(self, tag, qname)
        elif tag == (TEXTNS, 'h') or tag == (TEXTNS, 'p'):
            self.in_text -= 1
            line = ''.join(self.data)
            if user_options.debug:
                print("End:", tag, line, self.current_style.name)
            self.data = []
            if self.current_style.uppercase:
                line=line.upper()
            if self.current_style.is_title:
                self.state = DocumentState.TITLES
            elif self.current_style.page_break:
                if self.state == DocumentState.STARTING:
                    self.state = DocumentState.BODY
                if self.state == DocumentState.TITLES:
                    self.state = DocumentState.BODY
                elif self.state == DocumentState.BODY and len(self.r[DocumentState.BODY]) > 0:
                    self.r[DocumentState.BODY].append('===')
            if line == '':
                if self.state != DocumentState.STARTING:
                    self.state = DocumentState.BODY
            else:
                if self.state == DocumentState.STARTING:
                    self.state = DocumentState.TITLES
                if self.state == DocumentState.TITLES: # If style is a title one the keyword may be missing
                    if line[0] in "\t "  or \
                            ( self.current_style.margin_left is not None and self.current_style.margin_left > 25.0):
                        line = '    '+line.strip(' \t')
                    elif ':' in line:
                        line = line.strip(' \t')
                    else:
                        line = 'Title: '+line.strip(' \t')
                if self.state == DocumentState.BODY:
                    self.outputFountainPart( self.current_style.fountain_info, line)
                else:
                    self.outputText(line)
        elif tag==(TEXTNS, 'span'):
            self.data.append(self.spans[-1])
            self.spans.pop()

    def outputText(self, line: str):
        self.last_line_blank = line.strip() == ''
        self.r[self.state].append(line)

    # Output lines to body of document.
    # Recursive as some types need to follow other specific types
    def outputFountainPart( self, rule: FountainRule, line: str ):
        if rule.fountain_type == FountainType.CHARACTER:
            line = line.strip()
            if line == '':
                line = self.last_character
            else:
                self.last_character = line
        # if either forcing types or the type always requires special characters, and not
        # already present insert them
        if ( user_options.forcetypes or rule.always_required ) and \
            line[0:len(rule.prefix)] != rule.prefix:
            line=rule.prefix + line.strip() + self.current_style.fountain_info.suffix
        # some types restrict which types they can follow
        if rule.require_before and self.last_rule.fountain_type not in rule.require_before:
            self.outputFountainPart( fountain_rules_by_type[ rule.require_before[0]], '')
        # some types require special conditions
        if rule.blank_before and not self.last_line_blank:
            self.outputText('')
        self.last_rule = rule
        self.outputText(line)
        if rule.blank_after:
            self.outputText('')

def getZipPart(odffile, xmlfile):
    """ Get the content out of the ODT file"""
    z = zipfile.ZipFile(odffile)
    content = z.read(xmlfile)
    z.close()
    return content

def odtFileDecode(odtfile: str):
    mimetype = getZipPart(odtfile,'mimetype')
    content = getZipPart(odtfile,'content.xml')
    lines = []
    parser = make_parser()
    parser.setFeature(handler.feature_namespaces, 1)
    if not isinstance(mimetype, str):
        mimetype=mimetype.decode("utf-8")
    if mimetype in ('application/vnd.oasis.opendocument.text',
      'application/vnd.oasis.opendocument.text-template'):
        load_ODT_styles(odtfile)
        parser.setContentHandler(OdtXmlTextHandler(lines))
    else:
        print ("Unsupported file format")
        sys.exit(2)
    parser.setErrorHandler(handler.ErrorHandler())

    inpsrc = InputSource()
    if not isinstance(content, str):
        content=content.decode("utf-8")
    inpsrc.setByteStream(StringIO(content))
    parser.parse(inpsrc)
    return lines

def odf2fountain_main():
    global user_options
    argParser = argparse.ArgumentParser(description='Open Document text to Fountain converter.')
    argParser.add_argument('files', nargs='+', type=Path, help = "input files space separated" )
    argParser.add_argument('--output', '-output', type=Path, \
            help = "output filename. Default = input filename.fountain" )
    argParser.add_argument('--forcetypes', '-forcetypes', action="store_true",
            help="Start lines with a special character rather than let the fountain translator use heuristics to determine line type." )
    argParser.add_argument('--extendedfountain', '-extendedfountain', action="store_true",
            help="Enable extra type forcing characters not supported by most fountain translators" )
    argParser.add_argument('--debug', '-debug', action="store_true", help="provide developer information" )
    user_options = argParser.parse_args( sys.argv[1:] )
    if user_options.extendedfountain:
        fountain_rules['Dialogue'].prefix = '%'
    if user_options.output and len(user_options.files) != 1:
        raise Exception("--output file is only valid when a single input file is specified")
    filename: Path = Path('.')
    for filename in user_options.files:
        groups = odtFileDecode(str(filename))

        outstring = ''
        if len(groups[DocumentState.TITLES])> 0:
            outstring = '\n'.join(groups[DocumentState.TITLES])
            outstring += '\n'
        outstring += '\n'.join(groups[DocumentState.BODY]) + '\n'
        if user_options.output:
            outfile = user_options.output
        else:
            outfile = filename.parent / (filename.stem + '.fountain')
        outfile.write_text(outstring, 'utf-8')

if __name__ == "__main__":
    odf2fountain_main()