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
#  7/08/2022    Julia Clement     v 0.0.0     Initial version
# 10-12/08/2022 Julia Clement     v 0.0.0     Improve usability, separate lib module
# 19/08/2022    Julia Clement     v 0.0.0     Various Clean-ups. Add .pdf output
#

from enum import IntFlag
import argparse
from pathlib import Path
from unicodedata import name
from odfdo import Document, Paragraph, Element, Style
from odfdo.xmlpart import XmlPart
import sys

# access our shared library.
# expect to find it on the path or in either the same directory as this module or ../lib
# like Pooh I know there must be a better way but can't think what it might be
try:
    from odf_fountain_lib import toPoints, ifNull
except ModuleNotFoundError:
    selfPath = Path(__file__).parent
    sys.path.append(str(selfPath))
    utilPath = selfPath.parent / 'lib'
    if utilPath.is_dir():
        sys.path.append(str(utilPath))
    from odf_fountain_lib import toPoints, ifNull

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

class OdtStyle:
    """
    Information we need to know about styles
    TODO: Very similar to the class of the same name in odf2fountain. Merge?
    """
    allStyles = {}
    stylesToParent=[]
    def __init__(self, style:Style ):
        self.margin_left = None
        self.margin_right = None
        self.margin_top = None
        self.margin_bottom = None
        self.break_after = None
        self.break_before = None
        # Fixme:
        parentName = style.parent_style
        name = style.name
        self.fountainInfo=None
        for child in style.children:
            if (value := child.attributes.get('fo:break-after')) and value == 'page':
                self.break_after=True
            if (value := child.attributes.get('fo:break-before')) and value == 'page':
                self.break_before=True
            if (value := child.attributes.get('fo:margin-top')):
                self.margin_top=toPoints(value)
            if (value := child.attributes.get('fo:margin-bottom')):
                self.margin_bottom=toPoints(value)
            if (value := child.attributes.get('fo:margin-left')):
                self.margin_left=toPoints(value)
            if (value := child.attributes.get('fo:margin-right')):
                self.margin_right=toPoints(value)                
        self.name = name
        self.heritageLoaded = False
        self.parentName = parentName
        self.bold = None
        self.italic = None
        self.underline = None

        if name in ('Standard','Script_20_Elements', 'Heading'):
            # These two styles are special and end the parent chain
            self.parentName = name
            self.parent = self
            self.baseParentName = name
            self.baseParent = self
            self.is_base = True
            self.heritageLoaded = True
            #if not self.fountainInfo:
            #    self.fountainInfo=fountainRules['Null']
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
            self.fountainInfo = ifNull(self.fountainInfo, self.parent.fountainInfo)
        if self.parent and self.parent.heritageLoaded:
            self.maybeInheritFromParent()
        self.allStyles[self.name] = self
        self.stylesToParent.append(self)

    # inherit properties
    def maybeInheritFromParent(self):
        if self.parent and self.parent.heritageLoaded:
            self.heritageLoaded = True
            self.fountainInfo = ifNull( self.fountainInfo, self.parent.fountainInfo )
            self.italic = ifNull( self.italic, self.parent.italic )
            self.bold = ifNull( self.bold, self.parent.bold )
            self.underline = ifNull( self.underline, self.parent.underline)
            self.uppercase = ifNull( self.uppercase, self.parent.uppercase )
            self.align = ifNull( self.align, self.parent.align )
            self.border_line_width = ifNull( self.border_line_width, self.parent.border_line_width )
            self.border = ifNull( self.border, self.parent.border )
            self.margin_left = ifNull( self.margin_left, self.parent.margin_left )
            self.margin_right = ifNull( self.margin_right, self.parent.margin_right )
            self.margin_top = ifNull( self.margin_top, self.parent.margin_top )
            self.margin_bottom = ifNull( self.margin_bottom, self.parent.margin_bottom )
            self.page_break = ifNull( self.page_break, self.parent.page_break )
            self.break_before = ifNull( self.break_before, self.parent.break_before )
            self.break_after = ifNull( self.break_after, self.parent.break_after )
            self.is_title = ifNull( self.is_title, self.parent.is_title )

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
    
    def isSpaceBefore(self):
        if self.break_before or \
            (self.margin_top is not None and self.margin_top > 5 ):
            return True
        return False
    
    def isSpaceAfter(self):
        if self.break_after or \
            (self.margin_bottom is not None and self.margin_bottom > 5 ):
            return True
        return False

    def __str__(self):
        return f" {self.name}({self.parentName}) margins {self.margin_left}, {self.margin_right}, {self.fountainInfo}"

class FountainProcessor():
    # 
    styleReplacement = {
        # (Previous & new styles): (Use this style & blank lines following tracker)
        # Style names = base style + Ax for "After X"
        ('Title Line', 'Action'): ('Action ATi', True),
        ('Title Line', 'Centered'): ('Action ATi', True),
        ('Title Line', 'Scene Heading'): ('Scene Heading ATi', True),
        ('Transition', 'Scene Heading'): ('Scene Heading', True),
        ('Scene Heading', 'Character'): ('Character', False)
    }
    
    emphasiseSubstrs=[
        # start   end     style,    code(biu = bold, italics, underline), 
        ['_***',  '***_', None,     StyleFlags.BOLD_ITALIC_UNDERLINE],
        ['***',   '***',  None,     StyleFlags.BOLD_ITALIC],
        ['_**',   '**_',  None,     StyleFlags.BOLD_UNDERLINE],
        ['**',    '**',   None,     StyleFlags.BOLD],
        ['*',     '*',    None,     StyleFlags.ITALIC],
        ['_',     '_',    None,     StyleFlags.UNDERLINE],
    ]

    # Style Name, Parent Style, text to create
    needed_styles_list=[
        ['Script_20_Elements', 'Standard', '<style:style style:name="Script_20_Elements" style:display-name="Script Elements" style:family="paragraph" style:parent-style-name="Standard" style:class="text">'+
            '<style:paragraph-properties><style:tab-stops/></style:paragraph-properties>'+
            '<style:text-properties style:font-name="Liberation Mono" fo:font-family="&apos;Liberation Mono&apos;" style:font-style-name="Regular" style:font-family-generic="modern" style:font-pitch="fixed"/>'+
            '</style:style>'],
        ['Character', 'Script_20_Elements', '<style:style style:name="Character" style:family="paragraph" style:parent-style-name="Script_20_Elements" style:next-style-name="Dialogue" style:master-page-name="">'+
            '<style:paragraph-properties fo:margin-left="5.59cm" fo:margin-right="0cm" fo:margin-top="0.3528cm" fo:text-indent="0cm" style:auto-text-indent="false" style:page-number="auto" fo:keep-with-next="always"/>'+
            '<style:text-properties fo:text-transform="uppercase"/></style:style>'],
        ['Dialogue', 'Script_20_Elements', '<style:style style:name="Dialogue" style:family="paragraph" style:parent-style-name="Script_20_Elements" style:master-page-name="">'+
            '<style:paragraph-properties fo:margin-left="2.54cm" fo:margin-right="0cm" fo:text-indent="0cm" style:auto-text-indent="false" style:page-number="auto"/>'+
            '</style:style>'],
        ['Scene_20_Heading', 'Script_20_Elements', '<style:style style:name="Scene_20_Heading" style:display-name="Scene Heading" style:family="paragraph" style:parent-style-name="Script_20_Elements" style:next-style-name="Character" style:master-page-name="">'+
            '<style:paragraph-properties fo:margin-top="0.3528cm" fo:margin-bottom="0.3528cm" fo:orphans="4" fo:widows="4" style:page-number="auto" fo:keep-with-next="always"/>'+
            '<style:text-properties fo:text-transform="uppercase"/>'+
            '</style:style>'],
        ['Transition', 'Script_20_Elements', '<style:style style:name="Transition" style:family="paragraph" style:parent-style-name="Script_20_Elements" style:next-style-name="Scene_20_Heading" style:master-page-name="">'+
            '<style:paragraph-properties  fo:margin-top="0.3528cm" fo:margin-bottom="0.3528cm" fo:text-align="end" style:justify-single-word="false" style:page-number="auto" fo:keep-with-next="always">'+
            '<style:tab-stops/>'+
            '</style:paragraph-properties>'+
            '<style:text-properties fo:text-transform="uppercase" officeooo:rsid="0005d2eb"/>'+
            '</style:style>'],
        ['Action', 'Script_20_Elements', '<style:style style:name="Action" style:family="paragraph" style:parent-style-name="Script_20_Elements">'+
            '<style:paragraph-properties>'+
            '<style:tab-stops/>'+
            '</style:paragraph-properties>'+
            '<style:text-properties officeooo:rsid="0005d2eb"/>'+
            '</style:style>'],
        ['Parenthetical', 'Script_20_Elements', '<style:style style:name="Parenthetical" style:family="paragraph" style:parent-style-name="Script_20_Elements" style:next-style-name="Dialogue">'+
            '<style:paragraph-properties fo:margin-left="3.81cm" fo:margin-right="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>'+
            '<style:text-properties officeooo:rsid="0005d2eb"/></style:style>'],
        ['Lyrics', 'Dialogue', '<style:style style:name="Lyrics" style:family="paragraph" style:parent-style-name="Dialogue">'+
            '<style:text-properties fo:font-style="italic" officeooo:rsid="0008211f"/></style:style>'],
        ['Centered', 'Action', '<style:style style:name="Centered" style:family="paragraph" style:parent-style-name="Action">'+
            '<style:paragraph-properties fo:text-align="center" style:justify-single-word="false"/>'+
            '</style:style>'],
        ['Notes', 'Script_20_Elements', '<style:style style:name="Notes" style:family="paragraph" style:parent-style-name="Script_20_Elements" style:next-style-name="Dialogue">'+
            '<style:paragraph-properties fo:margin-left="1.27cm" fo:margin-right="0cm" fo:text-indent="0cm" style:auto-text-indent="false" style:border-line-width="0cm 0.026cm 0.026cm" fo:padding="0.049cm" fo:border="1.5pt double #000000">'+
            '<style:tab-stops/>'+
            '</style:paragraph-properties>'+
            '<style:text-properties fo:font-style="italic" fo:background-color="#fff5ce"/>'+
            '</style:style>'],
        ['Title_20_Line','Script_20_Elements','''<style:style style:name="Title_20_Line" style:display-name="Title Line" style:family="paragraph" style:parent-style-name="Script_20_Elements">
<style:text-properties style:font-name="Liberation Sans1" fo:font-family="&apos;Liberation Sans&apos;" style:font-style-name="Regular" style:font-family-generic="swiss" style:font-pitch="variable" fo:font-size="14pt"/>
</style:style>'''],
        ['Title_20_Line_20_Centered','Title_20_Line','''<style:style style:name="Title_20_Line_20_Centered" style:display-name="Title Line Centered" style:family="paragraph" style:parent-style-name="Title_20_Line">
<style:paragraph-properties fo:margin-left="5.08cm" style:justify-single-word="false"/>
</style:style>'''],
        ['Title_20_Ends','Title_20_Line','''<style:style style:name="Title_20_Ends" style:display-name="Title Ends" style:family="paragraph" style:parent-style-name="Title_20_Line" style:master-page-name="">
<style:paragraph-properties style:page-number="auto" fo:break-after="page" style:border-line-width-bottom="0.018cm 0.004cm 0.018cm" fo:padding="0.049cm" fo:border-left="none" fo:border-right="none" fo:border-top="none" fo:border-bottom="1.11pt double-thin #808080"/>
<style:text-properties officeooo:rsid="001983ec"/>
</style:style>'''],
        ['Character_20_AS', 'Character', '''<style:style style:name="Character_20_AS" style:display-name="Character AS" style:family="paragraph" style:parent-style-name="Character" style:next-style-name="Dialogue">
<style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm" style:contextual-spacing="false"/>
<style:text-properties officeooo:rsid="0024c7a2"/>
</style:style>'''],
        ['Scene_20_Heading_20_ATi', 'Scene_20_Heading', '''<style:style style:name="Scene_20_Heading_20_ATi" style:display-name="Scene Heading ATi" style:family="paragraph" style:parent-style-name="Scene_20_Heading" style:next-style-name="Action" style:master-page-name="Standard">
<style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.3528cm" style:contextual-spacing="false" style:page-number="1"/>
</style:style>'''],
        ['Scene_20_Heading_20_PB', 'Scene_20_Heading', '''<style:style style:name="Scene_20_Heading_20_PB" style:display-name="Scene Heading PB" style:family="paragraph" style:parent-style-name="Scene_20_Heading" style:next-style-name="Action" style:master-page-name="">
<style:paragraph-properties style:page-number="auto" fo:break-before="page"/>
<style:text-properties officeooo:rsid="00273222"/>
</style:style>'''],
        ['Character_20_PB', 'Character', '''<style:style style:name="Character_20_PB" style:display-name="Character PB" style:family="paragraph" style:parent-style-name="Character" style:next-style-name="Action" style:master-page-name="">
<style:paragraph-properties style:page-number="auto" fo:break-before="page"/>
</style:style>'''],
        ['Action_20_PB', 'Action', '''<style:style style:name="Action_20_PB" style:display-name="Action PB" style:family="paragraph" style:parent-style-name="Action" style:next-style-name="Action" style:master-page-name="">
<style:paragraph-properties style:page-number="auto" fo:break-before="page"/>
<style:text-properties officeooo:rsid="00278ad8"/>
</style:style>'''],
        ['Notes_20_PB', 'Notes', '''<style:style style:name="Notes_20_PB" style:display-name="Notes PB" style:family="paragraph" style:parent-style-name="Notes" style:next-style-name="Dialogue" style:master-page-name="">
<style:paragraph-properties style:page-number="auto" fo:break-before="page"/>
</style:style>'''],
        ['Centered_20_PB', 'Centered', '''<style:style style:name="Centered_20_PB" style:display-name="Centered PB" style:family="paragraph" style:parent-style-name="Centered" style:next-style-name="Centered" style:master-page-name="">
<style:paragraph-properties style:page-number="auto" fo:break-before="page"/>
</style:style>'''],
        ['Parenthetical_20_PB', 'Parenthetical', '''<style:style style:name="Parenthetical_20_PB" style:display-name="Parenthetical PB" style:family="paragraph" style:parent-style-name="Parenthetical" style:next-style-name="Dialogue" style:master-page-name="">
<style:paragraph-properties style:page-number="auto" fo:break-before="page"/>
</style:style>'''],
        ['Scene_20_Heading_20_ATr', 'Scene_20_Heading', '''<style:style style:name="Scene_20_Heading_20_ATr" style:display-name="Scene Heading ATr" style:family="paragraph" style:parent-style-name="Scene_20_Heading" style:next-style-name="Action" style:master-page-name="">
<style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.3528cm" style:contextual-spacing="false" style:page-number="auto"/>
</style:style>'''],
        ['Transition_20_PB', 'Transition', '''<style:style style:name="Transition_20_PB" style:display-name="Transition PB" style:family="paragraph" style:parent-style-name="Transition" style:next-style-name="Scene_20_Heading_20_ATr" style:master-page-name="">
<style:paragraph-properties style:page-number="auto" fo:break-before="page"/>
</style:style>'''],
        ['Dialogue_20_PB', 'Dialogue', '''<style:style style:name="Dialogue_20_PB" style:display-name="Dialogue PB" style:family="paragraph" style:parent-style-name="Dialogue" style:next-style-name="Dialogue" style:master-page-name="">
<style:paragraph-properties style:page-number="auto" fo:break-before="page"/>
</style:style>'''],
        ['Lyrics_20_PB', 'Lyrics', '''<style:style style:name="Lyrics_20_PB" style:display-name="Lyrics PB" style:family="paragraph" style:parent-style-name="Lyrics" style:next-style-name="Lyrics" style:master-page-name="">
<style:paragraph-properties style:page-number="auto" fo:break-before="page"/>
</style:style>'''],
        ['Action_20_ATi', 'Action', '''<style:style style:name="Action_20_ATi" style:display-name="Action ATi" style:family="paragraph" style:parent-style-name="Action" style:next-style-name="Action" style:master-page-name="Standard">
<style:paragraph-properties style:page-number="1"/>
</style:style>''']
    ]

    style_templates = {}
    known_styles={}

    def load_a_style( self, template ):
        if None==self.known_styles.get(template[1]):
            self.load_a_style( self.style_templates.get(template[1]))
        styleName=self.document.insert_style(template[2])
        style:Style=self.document.get_style('paragraph', styleName)
        self.known_styles[template[0]]=OdtStyle( style )

    def insert_style_templates(self):
        #Pass 1 load in the styles from the template
        for child in self.document.get_styles(family="paragraph"):
            if child.name:
                self.known_styles[child.name]=OdtStyle( child )
        #Pass 2 load the list of needed styles
        for child in self.needed_styles_list:
            self.style_templates[child[0]]=child
        #Pass 3 create any styles that don't exist or replace existing ones
        if self.userOptions.forcestyles:
            for child in self.needed_styles_list:
                styleName=self.document.insert_style(child[2])
                style:Style=self.document.get_style('paragraph', styleName)
                self.known_styles[child[0]]=OdtStyle( style )
        else:
            for child in self.needed_styles_list:
                if None==self.known_styles.get(child[0]):
                    self.load_a_style( child )

    def attributesToStr( self, prefix, collection, suffix='' ):
        """
            From a dictionary {a:1, b:2, ...} create the xml entity
            <ent a="1" b="2" c="3" />
        """
        answer=""
        for item,value in collection.items():
            answer+=f' {item}="{value}"'
        return prefix+answer+suffix

    def globalOptions( self, userOptions) :
        # Set papersize & margins
        # This is really horrible, odfdo doesn't seem to provide a way of replacing 
        # the attribute values in an entity definition :(
        # As a work around, we pull the original style apart, rebuild it and replace
        # the original. YUCK!!!
        oldStyle:Style = None
        if oldStyle:=self.document.get_style('page-layout', 'Mpm1'):
            output:str=self.attributesToStr('<style:page-layout', oldStyle.attributes, '>')
            for child in oldStyle.children:
                if child.tag == 'style:page-layout-properties':
                    attrs=child.attributes
                    if userOptions.papersize and userOptions.papersize != 'asis':
                        width, height = {'A4' : ('21cm', '29.7cm'),
                                        'US' : ('8.5in', '11in'),
                                        'LE' : ('8.5in', '11in')
                                        }.get( userOptions.papersize[:2].upper() )
                        attrs['fo:page-width'] = width
                        attrs['fo:page-height'] = height
                    if userOptions.margins.upper() != 'ASIS':
                        attrs['fo:margin-left'] = '1.5in'
                        attrs['fo:margin-right'] = '1in'
                        attrs['fo:margin-top'] = '0.7874in'
                        attrs['fo:margin-bottom'] = '1in'
                    output+=self.attributesToStr('<style:page-layout-properties',attrs,'/>')
                else:
                    output+=child.serialize()
            output+='</style:page-layout>'
            oldStyle.delete()
            self.document.insert_style(output,'Mpm1',automatic=True)

        # Turn off compatability option 'Add spacing between paragraphs and tabs'
        # Fortunately we can replace the text of a config-item, so much easier
        docSettings:XmlPart=self.document.get_part("settings")
        setting:Element=None
        for setting in docSettings.xpath("//config:config-item[@config:name='AddParaTableSpacing']"):
            setting.text = 'false'
        for setting in docSettings.xpath("//config:config-item[@config:name='AddParaTableSpacingAtStart']"):
            setting.text = 'false'

    def __init__(self, userOptions) :
        self.userOptions = userOptions
        if self.userOptions.template:
            self.document = Document(self.userOptions.template.name)
        else:
            self.document = Document("text")
            
        # New / empty .odt documents contain a single blank paragraph.
        # Delete if present
        self.docbody = self.document.body
        paragraphs =  self.docbody.get_paragraphs()
        if len(paragraphs) == 1:
             self.docbody.delete(paragraphs[0])
        self.inTitles = False
        self.linenr = 0
        self.maxline = 0
        self.style=''
        self.lastStyle=''
        self.BlankPending = False
        self.Blank = True
        self.lastBlank = True
        self.pageBreakRequired = False
        self.maxAutoStyle = 0
        self.globalOptions( userOptions)
        self.insert_style_templates()
        self.lastTitleStyle='Title Line'
        for s in self.document.get_styles(family='text', automatic=True):
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
            '%':    lambda l: self.addLine(l[1:], 'Dialogue', mergeLines=True), # non standard extension
            '\'':   lambda l: self.addLine(l, 'Dialogue', mergeLines=True), # Romeo & Juliet is riddled with lines starting 'Tis
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
                docline=Paragraph(line.strip('_* \t'),style="Title")
                self.docbody.append(docline)
                self.lastTitleStyle= 'Title Line Centered'
                firstTitle = False
            elif ':' in line:
                parts=line.strip().split(':',1)
                self.lastTitleStyle= 'Title Line Centered' if parts[0].lower() in ['title', 'credit','author','authors','source'] else 'Title Line'
                if '_' in line or '*' in line:
                    line=self.emphasise( line, self.lastTitleStyle, required_before=' :\t' )
                    docline=Element.from_tag( '<text:p text:style-name="'+\
                            self.lastTitleStyle+'">'+line+'</text:p>' )
                else:
                    docline=Paragraph(line.strip(), style=self.lastTitleStyle)
                self.docbody.append(docline)
            elif len( line) > 3 and (line[0:3] == '   ' or line[0]=='\t'):
                line=line.strip()
                if '_' in line or '*' in line:
                    line=self.emphasise( line, self.lastTitleStyle, required_before=' :\t' )
                docline=Element.from_tag( '<text:p text:style-name="'+\
                        self.lastTitleStyle+'"><text:tab/>'+line+'</text:p>' )
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
        self.document.insert_style(newstyle, automatic=True)
        return styleName

    def emphasise( self, line, localStyle, parentSettings=StyleFlags.NONE, required_before=' \t' ):
        newline = line
        for rule in self.emphasiseSubstrs:
            matchStr, endStr, style, settings = rule
            line=newline 
            newline=''
            while line != '':
                if (start:=line.find(matchStr)) >=0:
                    if start == 0 or line[start-1] in required_before:
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

    def addLine(self, line, style, fakeStyle=None, mergeLines = False ):
        if (replacement:=self.styleReplacement.get((self.lastStyle, style))):
            localStyle = replacement[0]
        elif self.pageBreakRequired:
            localStyle = style + ' PB'
        else:
            localStyle = style
        self.pageBreakRequired = False
        internalStyleName=localStyle.replace(' ','_20_')
        styleInfo : OdtStyle = self.known_styles.get(internalStyleName) or \
                               self.known_styles.get('Action')
        if self.BlankPending:
            if not styleInfo.isSpaceBefore():
                docline=Paragraph(' ')
                self.docbody.append(docline)
            self.BlankPending = False
        if '_' in line or '*' in line:
            line=self.emphasise( line, localStyle )
            docline=Element.from_tag( '<text:p text:style-name="'+\
                internalStyleName+'">'+line+'</text:p>' )
        else:
            docline=Paragraph(line, style=localStyle)
        self.docbody.append(docline)
        self.style= fakeStyle if fakeStyle else style
        if styleInfo.isSpaceAfter():
            self.Blank = True
        else:
            self.Blank = line.strip() == ''
        while mergeLines and self.linenr < len(self.lines) and \
              self.lines[self.linenr+1][0].isalnum():
            self.linenr += 1
            self.addLine( self.lines[self.linenr], style)

    def pageBreak(self):
        self.pageBreakRequired = True
        self.Blank = True
        self.BlankPending = False

    def blank(self):
        self.Blank = True
        self.BlankPending = not self.lastBlank

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
            self.addLine( line, 'Transition' )
        else:
            self.addLine( line.rstrip('<'), 'Centered', 'Action')

    def processFile(self, lines):
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
                self.blank()
            elif (fun:= self.start2Method.get(line[0])):
                fun(line)
            elif len(line) > 1 and line[0:2] == '[[':
                self.noteBlock( line )
            elif (l := line.strip()[0]) == '(':
                self.addLine(l, 'Parenthetical')
            elif line[0:5] in ['INT. ', 'EXT. '] and \
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
                self.addLine(line, 'Dialogue', mergeLines=True)
            else:
                self.addLine(line, 'Action')
            self.linenr += 1
            self.lastBlank = self.Blank
            self.lastStyle = self.style

    def processFiles(self):
        for fileName in self.userOptions.files:
            with open(fileName, 'r') as file:
                lines = file.readlines()
                if self.userOptions.debug:
                    print( lines )
                self.processFile(lines)

    def saveOdt(self):
        from os import chdir, getcwd, system
        if self.userOptions.output:
            outputFile = self.userOptions.output
        else:
            sourceFile = self.userOptions.files[0]
            outputFile = sourceFile.parent / (sourceFile.stem + '.odt')
        self.document.save(outputFile, pretty=True)
        if self.userOptions.pdf or self.userOptions.docx:
            olddir=getcwd()
            chdir(outputFile.parent)
            if self.userOptions.pdf:
                system("soffice  --headless --convert-to pdf "+str(outputFile))
            if self.userOptions.docx:
                system("soffice  --headless --convert-to docx "+str(outputFile))
            chdir( olddir )

    def run(self):
        self.processFiles()
        self.saveOdt()

class ArgOptions:
    class Arg:
        def __init__(self, *pargs, **kw) -> None:
            self.pargs = pargs
            self.name=pargs[0].strip('-')
            self.kw=kw
        def set_default(self, default ):
            self.kw['default']=default
    def __init__(self) -> None:
        self.args={}
    def add_argument(self, *pargs, **kw ):
        arg=self.Arg(*pargs,**kw)
        self.args[arg.name]=arg
    def set_default(self, opt, default):
        try:
            arg=self.args[opt]
            arg.set_default(default)
        except KeyError:
            print(f'Unknown option {opt}, skipped')
    def export( self, argParser:argparse.ArgumentParser) :
        for arg in self.args.values():
            argParser.add_argument( *arg.pargs, **arg.kw)

class Fountain2odf():
    def __init__( self ):
        self.argOptions=ArgOptions()
        self.argOptions.add_argument('--output', '-o', type=Path, \
                help = "output filename. Default = an empty odt file." )
        self.argOptions.add_argument('--template', '-t', type=Path, \
                help = "File with the Screenplay styles pre-loaded. Template files are supported." )
        self.argOptions.add_argument('--forcestyles', '-fs', action="store_true", \
                help = "Replace existing styles of the same name in the template with the current versions" )
        self.argOptions.add_argument('--pdf', action="store_true", \
                help = "use LibreOffice or Apache OpenOffice to create a PDF file in the same directory as the output file. "+ \
                       "Requires LibreOffice or Apache OpenOffice installed and in the current path" )
        self.argOptions.add_argument('--docx', action="store_true", \
                help = "use LibreOffice or Apache OpenOffice to create a MS Word file in the same directory as the output file. "+ \
                       "Requires LibreOffice or Apache OpenOffice installed and in the current path" )
        self.argOptions.add_argument('--papersize','-p', choices=['a4', 'A4', 'asis', 'US', 'Letter', 'US Letter'], default='asis',\
                help = "Document's page size. Default = the current setting of the template file, if any, or your LibreOffice default")
        self.argOptions.add_argument('--margins','-m', choices=['Standard', 'standard', 'asis', 'STD', 'std'], default='standard', 
                help="Page margins. Asis = use whatever the template or LibreOffice uses as default. Standard = 1/1.5 inches all around")
        self.argOptions.add_argument('--config', type=Path,
                help="Configuration file which will be merged with command-line options. See documentation for format.")
        self.argOptions.add_argument('--debug', action="store_true", help="provide developer information" )

        self.argParser=self.createParser()
        self.userOptions = argparse.Namespace()

    def createParser(self)->argparse.ArgumentParser:
        argParser = argparse.ArgumentParser(description='Fountain to Open Document text converter.')
        argParser.add_argument('prog', type=Path, help = "" )
        argParser.add_argument('files', nargs='+', type=Path, help = "input files space separated" )
        self.argOptions.export(argParser)
        return argParser

    def parseArgs( self, args = sys.argv ):
        if len(args) == 0:
            args = ['prog', '--help']
        self.userOptions = self.argParser.parse_args( args )
        if self.userOptions.config:
            storeTrues = []
            with open(self.userOptions.config, 'r') as configFile:
                lines = configFile.readlines()
            for line in lines:
                line=line.strip('\t\r\n -')
                if line[0] == '#':
                    pass
                elif '=' in line:
                    s=line.split('=',maxsplit=1)
                    self.argOptions.set_default(s[0].strip(),s[1].strip())
                elif ' ' in line:
                    s=line.split(maxsplit=1)
                    self.argOptions.set_default(s[0],s[1])
                else:
                    storeTrues.append(line)
        # second pass with defaults modified
        self.argParser = self.createParser()
        self.userOptions = self.argParser.parse_args( args )
        for opt in storeTrues:
            self.userOptions.__setattr__(opt, True)
        return self.userOptions

if __name__ == "__main__":
    prog=Fountain2odf()
    prog.parseArgs()
    processor = FountainProcessor( prog.userOptions )
    processor.run()
