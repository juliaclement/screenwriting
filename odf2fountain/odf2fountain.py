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
# 3/08/2022 Julia Clement     v 0.0.0     Initial version
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

# Convert measurements
# Libre Office supports setting measurement units to
# mm, cm, inch, pica & points.
# As points are the most fine of these measurements, we
# convert the measurements in the document to points 
measurementFactors = {
    'pt':   1.0,
    'pc':   12.0, # not found in my sample documents
    'in':   72.0,
    'cm':   28.3465,
    'mm':   2.83465
}

def toPoints( value: str ) :
    uom = value[-2:]
    factor = measurementFactors.get(uom,1.0)
    return float(value.strip('cimitp '))*factor

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
    def __init__( self, name, fountainType=None, prefix='', suffix='', blankBefore=False, \
                  blankAfter=False, alwaysRequired=False, requireBefore=[]):
        self.name : str = name
        self.fountainType :FountainType=fountainType
        self.prefix : str=prefix
        self.suffix : str=suffix
        self.blankBefore : bool=blankBefore
        self.blankAfter : bool=blankAfter
        self.alwaysRequired : bool=alwaysRequired
        self.requireBefore : list [FountainType]= requireBefore
    
    def __str__(self):
        answer=f" {self.fountainType}({self.prefix}, {self.suffix})"
        if self.blankBefore :
            answer += ", Blank Before"
        if self.blankAfter :
            answer += ", Blank After"
        if self.requireBefore :
            answer += f",require before [{self.requireBefore}]"
        return answer

# Lookup the primary rule for a FountainType
# Order must be by FountainType without gaps or duplicates
# TODO: This order rule is far too fragile, convert to a dict 
fountainRulesByType = [
    FountainRule( 'Null', FountainType.NULL),   #No action, don't change type
    FountainRule( 'Title', FountainType.TITLE, 'Title:', alwaysRequired=True),
    FountainRule( 'Action', FountainType.ACTION, '!'),
    FountainRule( 'Character', FountainType.CHARACTER, '@', blankBefore=True),
    FountainRule( 'Dialogue', FountainType.DIALOGUE,
                requireBefore=[FountainType.CHARACTER,FountainType.PARENTHETICAL, FountainType.DIALOGUE]),
    FountainRule( 'Notes', FountainType.NOTES, '[[', ']]', alwaysRequired=True),
    FountainRule( 'Parenthetical', FountainType.PARENTHETICAL, '(', ')', alwaysRequired=True,
                requireBefore=[FountainType.CHARACTER, FountainType.DIALOGUE]),
    FountainRule( 'Scene', FountainType.SCENE_HEADING, '.', blankBefore=True, blankAfter=True),
    FountainRule( 'Transition', FountainType.TRANSITION, '>', blankBefore=True, blankAfter=True),
]
    

fountainRules = {   'Subtitle': FountainRule( 'Subtitle', FountainType.TITLE ),
                    'Centred':  FountainRule( 'Centred', FountainType.ACTION, '>', '<', alwaysRequired=True),
                    'Lyrics':   FountainRule( 'Lyrics', FountainType.DIALOGUE, '~', alwaysRequired=True,
                                requireBefore=[FountainType.CHARACTER,FountainType.PARENTHETICAL]),
                }
for rule in fountainRulesByType:
    fountainRules[ rule.name ] = rule


StyleToFountain = { 
    'Title':                fountainRules['Title'],
    'Subtitle':             fountainRules['Subtitle'],
    'Title_20_Line':        fountainRules['Subtitle'],
    'Script_20_Elements':   fountainRules['Null'],
    'Standard':             fountainRules['Null'],
    'Action':               fountainRules['Action'],
    'Centered':             fountainRules['Centred'],
    'Character':            fountainRules['Character'],
    'Dialogue':             fountainRules['Dialogue'],
    'Lyrics':               fountainRules['Lyrics'],
    'Notes':                fountainRules['Notes'],
    'Parenthetical':        fountainRules['Parenthetical'],
    'Heading':              fountainRules['Null'],
    'Scene_20_Heading':     fountainRules['Scene'],
    'Scene':                fountainRules['Scene'],
    'Transition':           fountainRules['Transition']
}

def assignIfNull( old, new ):
    if old is None:
        return new
    return old

class OdtStyle():
    allStyles = {}
    stylesToParent=[]

    def __init__(self, name, parentName, margin_left=None, margin_right=None):
        self.name = name
        self.fountainInfo = StyleToFountain.get(name)
        self.heritageLoaded = False
        self.margin_left=margin_left
        self.margin_right=margin_right
        self.margin_top = None
        self.margin_bottom = None
        self.parentName = parentName
        self.bold = None
        self.italic = None
        self.underline = None

        if name in ('Standard','Script_20_Elements'):
            # These two styles are special and end the parent chain
            self.parentName = name
            self.parent = self
            self.baseParentName = name
            self.baseParent = self
            self.is_base = True
            self.heritageLoaded = True
            if not self.fountainInfo:
                self.fountainInfo=fountainRules['Null']
            self.italic = False
            self.uppercase = False
            self.align='left'
            self.border_line_width = '0cm'
            self.border = '0.1pt double #000000'
            self.page_break = False
            self.is_title = False
        else:
            self.is_base = False
            self.baseParentName=parentName # will be adjusted later
            self.assignParentByName(parentName)
            self.uppercase = None
            self.align = None
            self.border_line_width = None
            self.border = None
            self.page_break = None
            if 'TITLE' in name.upper():
                self.is_title = True
            else:
                self.is_title = None

        if self.parent and self.parent.baseParent:
            self.baseParent = self.parent.baseParent
            self.fountainInfo = assignIfNull(self.fountainInfo, self.parent.fountainInfo)
        if self.parent and self.parent.heritageLoaded:
            self.maybeInheritFromParent()
        else:
            self.stylesToParent.append( self )

    # inherit properties
    def maybeInheritFromParent(self):
        if self.parent and self.parent.heritageLoaded:
            self.heritageLoaded = True
            self.fountainInfo = assignIfNull( self.fountainInfo, self.parent.fountainInfo )
            self.italic = assignIfNull( self.italic, self.parent.italic )
            self.bold = assignIfNull( self.bold, self.parent.bold )
            self.underline = assignIfNull( self.underline, self.parent.underline)
            self.uppercase = assignIfNull( self.uppercase, self.parent.uppercase )
            self.align = assignIfNull( self.align, self.parent.align )
            self.border_line_width = assignIfNull( self.border_line_width, self.parent.border_line_width )
            self.border = assignIfNull( self.border, self.parent.border )
            self.margin_left = assignIfNull( self.margin_left, self.parent.margin_left )
            self.margin_right = assignIfNull( self.margin_right, self.parent.margin_right )
            self.margin_top = assignIfNull( self.margin_top, self.parent.margin_top )
            self.margin_bottom = assignIfNull( self.margin_bottom, self.parent.margin_bottom )
            self.page_break = assignIfNull( self.page_break, self.parent.page_break )
            self.is_title = assignIfNull( self.is_title, self.parent.is_title )

    def assignParent( self, parent ):
        self.parent = parent
        self.baseParent = parent
        self.baseParentName = parent.name

    def assignParentByName( self, parentName ):
        self.parent = OdtStyle.allStyles.get(parentName)
        if self.parent:
            self.assignParent( self.parent )
        else:
            self.baseParentName = parentName

    def __str__(self):
        return f" {self.name}({self.parentName}) margins {self.margin_left}, {self.margin_right}, {self.fountainInfo}"

class OdtXmlStyleHandler(handler.ContentHandler):
    """ Extract headings from content.xml of an ODT file """
    def __init__(self):
        handler.ContentHandler.__init__(self)
        self.data = []
        self.currentStyle = None
        self.buildingStyle = False

    def startElementNS(self, tag, qname, attrs):
        if tag == (STYLENS, 'style'):
            if userOptions.debug:
                print("Start:",tag[1], "Attrs:", attrs.items())
            try:
                name = attrs._attrs[(STYLENS,'name')]
                parentName = attrs._attrs.get((STYLENS,'parent-style-name'), 'Standard')
                self.currentStyle=OdtStyle(name, parentName)
                self.buildingStyle = True
            except KeyError:
                pass
        elif tag[0] == STYLENS and tag[1] in ['paragraph-properties', 'text-properties'] and self.currentStyle :
            if userOptions.debug:
                print("Start:",tag[1], "Attrs:", attrs.items())
            for (attrName,value) in attrs.items():
                kw=attrName[1]
                if kw == 'margin-left':
                    self.currentStyle.margin_left = toPoints( value )
                elif kw == 'margin-right':
                    self.currentStyle.margin_right = toPoints( value )
                elif kw == 'margin-top':
                    self.currentStyle.margin_top = toPoints( value )
                elif kw == 'margin-bottom':
                    self.currentStyle.margin_bottom = toPoints( value )
                elif kw == 'text-transform':
                    if value == 'uppercase':
                        self.currentStyle.uppercase = True
                elif kw == 'text-align':
                    if value == 'end':
                        self.currentStyle.align = 'right'
                elif kw == 'font-style':
                    if value == 'italic':
                        self.currentStyle.italic = True
                elif kw == 'font-weight':
                    if value == 'bold':
                        self.currentStyle.bold = True
                elif kw == 'text-underline-style':
                    if value == 'solid':
                        self.currentStyle.underline = True
                elif kw == 'border':
                    self.currentStyle.border = value
                elif kw == 'break-before' and value=='page':
                    self.currentStyle.page_break = True
                elif kw == 'page-number' and value != 'auto': # implies page break
                    self.currentStyle.page_break = True
                elif kw == 'border-line-width':
                    self.currentStyle.border_line_width = value

    def endElementNS(self, tag, qname):
        if tag == (STYLENS, 'style') and self.currentStyle and self.buildingStyle:
            self.buildingStyle = False
            OdtStyle.allStyles[self.currentStyle.name] = self.currentStyle
            OdtStyle.stylesToParent.append(self.currentStyle)

def postProcessStyles(stylesToAssignBaseParents, stylesToAssignHeritage):
    #
    # 1) Assign base parents. This process  may also assign parents which
    #    may end up assigning inherited properties.
    #    As the order of elements in a dictionary is not defined,
    #    we (Have done in every test I've done so far) need
    #    to make multiple passes to assign all parents
    nextPass=stylesToAssignBaseParents
    redo=True
    while redo:
        thisPass = nextPass
        nextPass = {}
        for styleName in thisPass:
            style=OdtStyle.allStyles[styleName]
            if not style.parent:
                style.parent = OdtStyle.allStyles.get(style.parentName)
                if style.parent and not style.fountainInfo:
                    style.fountainInfo = style.parent.fountainInfo
            if not style.is_base:
                base_parent=OdtStyle.allStyles.get(style.baseParentName,None)
                if base_parent:
                    style.baseParent = base_parent.baseParent
                    style.baseParentName = base_parent.baseParentName
                    if not base_parent.is_base:
                        nextPass[styleName]=style
        redo=len(nextPass)<len(thisPass)
    #
    # 2) If needed, assign parents and inherit properties
    #
    #    TBH, I don't know if this will ever be needed, but just-in-case
    #    I also can't imagine multiple passes being needed, but will support anyway
    nextPass = stylesToAssignHeritage
    redo = True
    while redo:
        thisPass = nextPass
        nextPass = []
        for style in thisPass:
            if style.parentName and not style.heritageLoaded:
                style.assignParentByName(style.parentName)
                if not style.heritageLoaded:
                    nextPass.append(style)
        redo = len(nextPass) < len(thisPass)

def loadOdtStyles(odtfile):
    styleParser = make_parser()
    styleParser.setFeature(handler.feature_namespaces, 1)
    styleParser.setContentHandler(OdtXmlStyleHandler())
    content = getZipPart(odtfile,'styles.xml')
    if not isinstance(content, str):
        content=content.decode("utf-8")
    styleParser.setErrorHandler(handler.ErrorHandler())
    inpsrc = InputSource()
    # if not isinstance(content, line):
    #     content=content.decode("utf-8")
    inpsrc.setByteStream(StringIO(content))
    styleParser.parse(inpsrc)
    postProcessStyles(OdtStyle.allStyles, OdtStyle.stylesToParent)
    if userOptions.debug:
        print( "Styles:")
        for style in OdtStyle.allStyles:
            print( OdtStyle.allStyles[style] )


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
    def __init__(self, lines : list[ list[str]]):
        OdtXmlStyleHandler.__init__(self)
        self.state :DocumentState = DocumentState.STARTING
        while len(lines) < 3:
            lines.append([])
        self.r : list[ list[str]] = lines
        self.data : list [ str ] = []
        self.lastLineBlank : bool = False
        self.lastRule : FountainRule = fountainRulesByType[0]
        self.lastCharacter : str = ''
        self.inText = 0

        #information to end a span
        #convert nested emphasis back to previous level when a span ends
        self.spans=['']

    def characters(self, chars : str):
        if self.inText > 0 and len(chars)>0:
            self.data.append(chars)

    def startElementNS(self, tag, qname, attrs):
        attrDict=attrs._attrs
        if tag == (TEXTNS, 'p') or tag == (TEXTNS, 'h'):
            if self.inText == 0:
                self.data = []
            self.inText += 1
            for (attrName,value) in attrs.items():
                if attrName[1] == 'style-name':
                    self.currentStyle = OdtStyle.allStyles.get(value, self.currentStyle)
            if userOptions.debug:
                print("Start:",tag[1], "Attrs:", attrDict)
        elif tag==(TEXTNS, 'span'):
            newStyle:OdtStyle=None
            for (attrName,value) in attrs.items():
                if attrName==(TEXTNS, 'style-name'):
                    newStyle = OdtStyle.allStyles.get(value, self.currentStyle)
            if newStyle:
                self.spans.append('')
                openText=''
                while len(self.data) > 0 and self.data[-1] == '':
                    self.data.pop()
                if len(self.data)> 0 and self.data[-1][-1] != ' ':
                    openText = ' '
                closeText=''
                if newStyle.underline:
                    openText+='_'
                    closeText='_'
                if newStyle.bold:
                    openText+='**'
                    closeText='**'+closeText
                if newStyle.italic:
                    openText+='*'
                    closeText='*'+closeText
                self.spans.append(closeText)
                self.data.append(openText)
            else:
                self.spans.append['']

        elif tag[0] == STYLENS and \
                tag[1] in ['style', 'paragraph-properties', 'text-properties'] :
            OdtXmlStyleHandler.startElementNS(self, tag, qname, attrs)
        elif tag == (TEXTNS, 'tab'):
            self.data.append("\t")
        elif tag == (TEXTNS, 's'):
            for (attrName,value) in attrs.items():
                shortName=attrName[1]
                if shortName == 'c':
                    count=int(value)
                    self.data.append( ' ' * count)

    def endElementNS(self, tag, qname):
        if tag[0] == STYLENS:
            if tag[1] == 'style' and self.currentStyle and self.buildingStyle:
                OdtXmlStyleHandler.endElementNS(self, tag, qname)
                postProcessStyles({self.currentStyle.name:self.currentStyle }, [self.currentStyle])
            else:
                OdtXmlStyleHandler.endElementNS(self, tag, qname)
        elif tag == (TEXTNS, 'h') or tag == (TEXTNS, 'p'):
            self.inText -= 1
            line = ''.join(self.data)
            if userOptions.debug:
                print("End:", tag, line, self.currentStyle.name)
            self.data = []
            if self.currentStyle.uppercase:
                line=line.upper()
            if self.currentStyle.is_title:
                self.state = DocumentState.TITLES
            elif self.currentStyle.page_break:
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
                            ( self.currentStyle.margin_left is not None and self.currentStyle.margin_left > 25.0) :
                        line = '    '+line.strip(' \t')
                    elif ':' in line:
                        line = line.strip(' \t')
                    else:
                        line = 'Title: '+line.strip(' \t')
                if self.state == DocumentState.BODY:
                    self.outputFountainPart( self.currentStyle.fountainInfo, line)
                else:
                    self.outputText(line)
        elif tag==(TEXTNS, 'span'):
            self.data.append(self.spans[-1])
            self.spans.pop()

    def outputText(self, line : str):
        self.lastLineBlank = line.strip() == ''
        self.r[self.state].append(line)

    # Output lines to body of document.
    # Recursive as some types need to follow other specific types
    def outputFountainPart( self, rule: FountainRule, line: str ) :
        if rule.fountainType == FountainType.CHARACTER :
            line = line.strip()
            if line == '' :
                line = self.lastCharacter
            else:
                self.lastCharacter = line
        # if either forcing types or the type always requires special characters, and not
        # already present insert them
        if ( userOptions.forcetypes or rule.alwaysRequired ) and \
            line[0:len(rule.prefix)] != rule.prefix:
            line=rule.prefix + line.strip() + self.currentStyle.fountainInfo.suffix
        # some types restrict which types they can follow
        if rule.requireBefore and self.lastRule.fountainType not in rule.requireBefore :
            self.outputFountainPart( fountainRulesByType[ rule.requireBefore[0]], '')
        # some types require special conditions
        if rule.blankBefore and not self.lastLineBlank:
            self.outputText('')
        self.lastRule = rule
        self.outputText(line)
        if rule.blankAfter:
            self.outputText('')

def getZipPart(odffile, xmlfile):
    """ Get the content out of the ODT file"""
    z = zipfile.ZipFile(odffile)
    content = z.read(xmlfile)
    z.close()
    return content

def odtFileDecode(odtfile : str):
    mimetype = getZipPart(odtfile,'mimetype')
    content = getZipPart(odtfile,'content.xml')
    lines = []
    parser = make_parser()
    parser.setFeature(handler.feature_namespaces, 1)
    if not isinstance(mimetype, str):
        mimetype=mimetype.decode("utf-8")
    if mimetype in ('application/vnd.oasis.opendocument.text',
      'application/vnd.oasis.opendocument.text-template'):
        loadOdtStyles(odtfile)
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

if __name__ == "__main__":
    argParser = argparse.ArgumentParser(description='Open Document text to Fountain converter.')
    argParser.add_argument('files', nargs='+', type=Path, help = "input files space separated" )
    argParser.add_argument('--output', '-output', type=Path, \
            help = "output filename. Default = input filename.fountain" )
    argParser.add_argument('--forcetypes', '-forcetypes', action="store_true",
            help="Start lines with a special character rather than let the fountain translator use heuristics to determine line type." )
    argParser.add_argument('--extendedfountain', '-extendedfountain', action="store_true",
            help="Enable extra type forcing characters not supported by most fountain translators" )
    argParser.add_argument('--debug', '-debug', action="store_true", help="provide developer information" )
    userOptions = argParser.parse_args( sys.argv[1:] )
    if userOptions.extendedfountain:
        fountainRules['Dialogue'].prefix = '%'
    filler = "          "
    if userOptions.output and len(userOptions.files) != 1:
        raise Exception("--output file is only valid when a single input file is specified")
    filename : Path = Path('.')
    for filename in userOptions.files:
        groups = odtFileDecode(filename.name)

        outstring = ''
        if len(groups[DocumentState.TITLES])> 0:
            outstring = '\n'.join(groups[DocumentState.TITLES])
            outstring += '\n'
        outstring += '\n'.join(groups[DocumentState.BODY]) + '\n'
        if userOptions.output :
            outfile = userOptions.output
        else:
            outfile = filename.parent / (filename.stem + '.fountain')
        outfile.write_text(outstring, 'utf-8')
