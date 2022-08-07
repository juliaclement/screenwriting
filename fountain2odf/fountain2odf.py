#!/usr/bin/python3
# -*- coding: utf-8 -*-
#
# Fountain to Open document text (.odt) translator
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
# 7/08/2022   Julia Clement     v 0.0.0     Initial version
#

from enum import IntFlag
import argparse
from pathlib import Path
from odfdo import Document, Paragraph, Element
import sys

class StyleFlags(IntFlag):
    # None
    NONE = 0
    # Primary
    UNDERLINE = 1
    ITALIC = 2
    BOLD = 4
    # Combinations
    ITALIC_UNDERLINE = 3
    BOLD_UNDERLINE = 5
    BOLD_ITALIC = 6
    BOLD_ITALIC_UNDERLINE = 7

class fountainDecoder():
    # 
    styleReplacement = {
        # (Previous & new styles): (Use this style & blank lines following tracker)
        # Style names = base style + Ax for "After X"
        ('Title Line', 'Action'): ('Action ATi', True),
        ('Title Line', 'Centered'): ('Action ATi', True),
        ('Title Line', 'Scene Heading'): ('Scene Heading ATi', True),
        ('Transition', 'Scene Heading'): ('Scene Heading ATr', True),
        ('Scene Heading', 'Character'): ('Character AS', False),
    }
    
    emphasiseSubstrs=[
        # start   end     style,    code(biu = bold, italics, underline), 
        ['_***',  '***_', None,     StyleFlags.BOLD_ITALIC_UNDERLINE],
        ['***',   '***',  None,     StyleFlags.BOLD_ITALIC            ],
        ['_**',   '**_',  None,     StyleFlags.BOLD_UNDERLINE],
        ['**',    '**',   None,     StyleFlags.BOLD],
        ['*',     '*',    None,     StyleFlags.ITALIC],
        ['_',     '_',    None,     StyleFlags.UNDERLINE],
    ]
    def __init__(self, document:Document) :
        self.doc = document
        self.docbody=document.body
        self.inTitles = False
        self.linenr = 0
        self.maxline = 0
        self.style=''
        self.lastStyle=''
        self.Blank = True
        self.lastBlank = True
        self.pageBreakRequired = False
        self.maxAutoStyle = 0
        for s in document.get_styles(family='text', automatic=True):
            autostyle=int(s.name[1:])
            self.maxAutoStyle = autostyle if autostyle > self.maxAutoStyle else self.maxAutoStyle
            settings=['','','']
            for c in s.children:
                if c.tag=='style:text-properties':
                    for name in c.attributes:
                        value=c.attributes[name]
                        if name == 'style:text-underline-style' and value=='solid':
                            settings[2]='u'
                        elif name == 'fo:font-style' and value=='italic':
                            settings[1]='i'
                        elif name == 'fo:font-weight' and value=='bold':
                            settings[0]='b'
                        else:
                            if userOptions.debug:
                                print( 'Unused', s.name, name,value)
            name='T'+''.join(settings)
            for e in self.emphasiseSubstrs:
                if e[3] == name:
                    e[2] = s.name
        self.start2Method = {
            '!':    lambda l: self.addLine(l[1:], 'Action' ),
            '@':    lambda l: self.addLine(l[1:], 'Character' ),
            '%':    lambda l: self.addLine(l[1:], 'Dialogue'),  # non standard extension
            '~':    lambda l: self.addLine(l[1:], 'Lyrics', 'Dialogue' ),
            '(':    lambda l: self.addLine(l, 'Parenthetical'), # expects the ()
            '.':    lambda l: self.addLine(l[1:], 'Scene Heading'),
            '>':    lambda l: self.transitionOrCentred(l), # expects the >
        }

    def process_titles( self ):
        self.inTitles = True
        firstTitle = True
        while self.linenr < self.maxline and \
              (line := self.lines[ self.linenr].rstrip('\n\r')) != '':
            if firstTitle:
                # first title becomes document title
                if len(line) > 6 and line[0:6] == 'Title:':
                    line=line[6:].strip()
                docline=Paragraph(line.strip(),style="Title")
                self.docbody.append(docline)
                firstTitle = False
            elif ':' in line:
                docline=Paragraph(line.strip(),style="Title Line")
                self.docbody.append(docline)
            elif len( line) > 3 and (line[0:3] == '   ' or line[0]=='\t'):
                docline=Paragraph(style="Title Line")
                docline.append_plain_text('    '+(line.strip()))
                self.docbody.append(docline)
            else:
                break
            self.linenr += 1
        if line == '':
            self.lastBlank = True
            self.linenr += 1
        self.inTitles = False
        self.lastStyle='Title Line'
        self.pageBreakRequired = True
        
    def getOrCreateStyle(self, ParentRule: StyleFlags):
        for rule in self.emphasiseSubstrs:
            if rule[3] == ParentRule:
                if rule[2]:
                    return rule[2]
                break
        self.maxAutoStyle += 1
        styleName = 'T'+str(self.maxAutoStyle)
        newstyle= f'<style:style style:name="{styleName}" style:family="text">'\
                  +'<style:text-properties'
        if rule[3] & StyleFlags.ITALIC:
            newstyle+=' fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"'
        if rule[3] & StyleFlags.BOLD:
            newstyle+=' fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"'
        if rule[3] & StyleFlags.UNDERLINE:
            newstyle+=' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"'
        newstyle+='/></style:style>'
        self.doc.insert_style(newstyle, automatic=True)
        return styleName

    def emphasise( self, line, localStyle, parentSettings=StyleFlags.NONE ):
        newline = line
        for rule in self.emphasiseSubstrs:
            matchStr, endStr, style, settings = rule
            line=newline 
            newline=''
            while line != '':
                if (start:=line.find(matchStr)) >=0:
                    if start == 0 or line[start-1] in ' \t':
                        newline+=line[:start]
                        line = line[start+len(matchStr):]
                        endMatch=line.find(endStr)
                        if endMatch >= 0:
                            bit=line[:endMatch]
                            line=line[endMatch+len(endStr):]
                            newSettings=parentSettings | settings
                            rule[2] = self.getOrCreateStyle(newSettings)
                            style=rule[2]
                            bit=self.emphasise(bit,localStyle, newSettings)
                            newline+=f'<text:span text:style-name="{style}">{bit}</text:span>'
                        else:
                            newline+=line
                            line=''
                    elif line[start-1] == '\\':
                        start += 1
                        newline += line[:start]
                        line = line[start:]
                    else:
                        split=start+len(matchStr)
                        newline+=line[:split]
                        line=line[split:]
                else:
                    newline+=line
                    line=''
        return newline

    def addLine(self, line, style, fakeStyle=None, blankAfter=False):
        if (replacement:=self.styleReplacement.get((self.lastStyle, style))):
            localStyle, blankAfter = replacement
        elif self.pageBreakRequired:
            localStyle = style + ' PB'
        else:
            localStyle = style
        if '_' in line or '*' in line:
            line=self.emphasise( line, localStyle )
            docline=Element.from_tag( '<text:p text:style-name="'+\
                    localStyle.replace(' ','_20_')+\
                    '">'+line+'</text:p>' )
        else:
            docline=Paragraph(line, style=localStyle)
        self.docbody.append(docline)
        self.style= fakeStyle if fakeStyle else style
        if blankAfter:
            self.Blank = True
        else:
            self.Blank = line.strip() == ''
        self.pageBreakRequired = False

    def pageBreak(self):
        self.pageBreakRequired = True
        self.Blank = True

    def blank(self, line):
        if not self.lastBlank:
            docline=Paragraph(' ')
            self.docbody.append(docline)
        self.Blank = True

    def noteBlock(self, line):
        line=line[2:]
        if ']]' in line:
            # single line note
            split=line.split(']]')
            style='Notes'
            for part in split:
                if part != '':
                    self.addLine(part, style)
                style='Dialogue'
        else:
            #multi-line notes. Not yet implemented
            self.addLine(line + '...  Not yet implemented', 'Notes')
        self.Blank = False
    
    def transitionOrCentred(self, line):
        if line[0] == '>':
            line=line[1:]
        if line[-1]==':':
            self.addLine( line, 'Transition', blankAfter=True)
        else:
            self.addLine( line.rstrip('<'), 'Centered', 'Action')

    def process(self, lines):
        if lines is str:
            lines = lines.replace(b'\r\n',b'\n')
            lines = lines.split(b'\n')
        
        if len(lines) < 2:
            print( "Emptyish file")
            return
        line : str = ''
        self.lines = lines
        self.linenr = 0
        self.maxline = len(lines)
        if lines[self.maxline-1].rstrip('\n\r ')=='':
            self.maxline -= 1
        else:
            lines.append('')
        # skip blank lines at start
        while self.linenr < self.maxline and lines[ self.linenr ].rstrip(' \r\n') == '':
            self.linenr += 1
        # optional titles
        if ':' in lines[ self.linenr] :
            self.process_titles()
        while self.linenr < self.maxline:
            line = lines[self.linenr].rstrip('\n\r')
            if len(line.strip()) == 0:
                self.blank(line)
            elif (fun:= self.start2Method.get(line[0])):
                fun(line)
            elif len(line) > 1 and line[0:2] == '[[':
                self.noteBlock( line )
            elif (l := line.strip()[0]) == '(':
                self.addLine(l, 'Parenthetical')
            elif line[0:4] in ['INT', 'EXT'] and \
                    lines[self.linenr+1].strip(' \t\r\n') == '':
                self.addLine(line, 'Scene Heading')
            elif self.lastBlank and line.upper() == line and \
                    lines[self.linenr+1].strip(' \t\r\n') != '':
                if line[-3:] == 'TO:':
                    self.addLine(line, 'Transition')
                else:
                    self.addLine(line, 'Character')
            elif line[0:3] == '===':
                self.pageBreak()
            elif self.lastStyle in ['Character', 'Parenthetical', 'Notes']:
                self.addLine(line, 'Dialogue')
            else:
                self.addLine(line, 'Action')
            self.linenr += 1
            self.lastBlank = self.Blank
            self.lastStyle = self.style


if __name__ == "__main__":
    argParser = argparse.ArgumentParser(description='Fountain to Open Document text converter.')
    argParser.add_argument('files', nargs='+', type=Path, help = "input files space separated" )
    argParser.add_argument('--output', '-output', type=Path, \
            help = "output filename. Default = input filename.odt" )
    argParser.add_argument('--template', '-template', type=Path, \
            default = 'Screenplay.odt', \
            help = "File with the Screenplay styles. Default = \"Screenplay.odt\". Template files are supported." )
    argParser.add_argument('--debug', '-debug', action="store_true", help="provide developer information" )
    userOptions = argParser.parse_args( sys.argv[1:] )
    filler = "          "
    filename : Path = Path('.')
    new_document = Document(userOptions.template.name)
    #
    # Delete existing text
    body = new_document.body
    paragraphs = body.get_paragraphs()
    for p in paragraphs:
        body.delete(p)
    decoder = fountainDecoder( new_document )
    for sourceFile in userOptions.files:
        with open(sourceFile.name, 'r') as file:
            lines = file.readlines()
            if userOptions.debug:
                print( lines )
            decoder.process(lines)

    if userOptions.output:
        outputFile = userOptions.output.name
    else:
        sourceFile = userOptions.files[0]
        outputFile = sourceFile.parent / (sourceFile.stem + '.odt')
    new_document.save(outputFile, pretty=True)

pass
