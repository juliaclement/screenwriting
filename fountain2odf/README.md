# Fountain2odf – Fountain to Open Document text converter

## Purpose:

Converts a .fountain (a screenplay / teleplay script interchange format)
file into a word processing file.

## Licences:

This program depends heavily on the [odfdo
library](https://github.com/jdum/odfdo) which is licenced under the
[Apache 2.0 licence](https://www.apache.org/licenses/LICENSE-2.0).
Although this does not create an obligation, I have decided to use the
Apache 2.0 licence for this program. You may use it only under the terms
of the Apache 2.0 licence. This licence is recognised as GPL compatible
so shouldn’t create problems for people wanting to reuse or extend the
utilities.

Documentation: All original non-program files are licenced under the
[Creative Commons Attribution-ShareAlike 4.0 International (CC
BY-SA 4.0)](https://creativecommons.org/licenses/by-sa/4.0/) licence.

## Usage:

After installing, this is used from the command line:

python fountain2odf.py \[-h\] \[--output OUTPUT\] \[--template
TEMPLATE\] \[--forcestyles\] \[--pdf\] \[--docx\] \[--papersize
{A4,asis,US Letter}\]

\[--margins {Standard,asis,std}\]

\[--debug\]

prog files \[files ...\]

positional arguments:

files                input files space separated

options:

\-h, --help           show help message and exit

\--output OUTPUT, -o OUTPUT

output filename. Default = an empty odt file.

\--template TEMPLATE, -t TEMPLATE

File with the Screenplay styles. Default = "Screenplay.odt". Template
files are supported.

\--forcestyles, -fs   Replace existing styles of the same name in the in
memory template with the current versions. Does not change the saved
copy

\--pdf                use LibreOffice or Apache OpenOffice to create a
PDF file in the same directory as the output file. Requires LibreOffice
or Apache

OpenOffice installed and in the current path

\--docx               use LibreOffice or Apache OpenOffice to create a
MS Word file in the same directory as the output file. Requires
LibreOffice or Apache

OpenOffice installed and in the current path

\--papersize {a4,A4,asis,US,Letter,US Letter}, also -p. Document's page
size. Default = the current setting of the template file, if any, or
your LibreOffice default

\--margins {Standard,standard,asis,STD,std}, -m
{Standard,standard,asis,STD,std}

Page margins. Asis = use whatever the template or LibreOffice uses as
default. Standard = 1/1.5 inches all around

\--debug              provide developer information

### Setup

1.  Install dependencies, see below,
2.  Download the program folder from github,
3.  Download the file odf\_fountain\_lib.py from the lib directory, and
    place in the same directory as fountain2odf.py

### Making PDF files

As its use case is to allow editing of files in a wordprocessor,
fountain2odf does **not**, by design, directly create PDF files, but
LibreOffice does and can run in a “headless” (batch) mode. 

As a convenience, at the end of processing, fountain2odf can run a
headless LibreOffice to copy the created Open Document file to .PDF &
even .DOCX formats. For this to work you will need to install
LibreOffice and make sure it’s in your path. It might also work with an
install of Apache OpenOffice but this is currently untested.

## Dependencies

  - The latest version of python3
  - The odfdo library. This can be installed with pip:  
    pip install odfdo

## Language

These documents and programs are in New Zealand English which does not
use American spelling. If you don’t like reading correctly spelled New
Zealand English, please don’t read them.

## Manifest:

Fountain2odf.py -- the program

README.md -- this file

ScreenplayA4.odt – A template file with the required styles and page
size set to A4

ScreenplayUS.odt – A template file with the required styles and page
size set to US Letter

## To do:

1.  fountain2odf relies on having a template file with the styles we use
    pre-loaded. These should be generated on the fly when not found in
    the template, allowing the template to be used for page size,
    header, footer, and boiler plate text.
2.  Writing some tests is required.

## Experimental

None of these tools are intended to do the complete job, but I am
prepared to consider requests for small changes that improve usability
and accept pull requests for enhancements.

I have no current plans to write installers, or Pip packages. For the
two most popular operating systems I lack the means to test created
installers. I would welcome requests from interested and technically
capable people who would like to join the project to create installers
for some or all platforms on an ongoing basis.

Some changes that can be justified, I am unwilling to do. examples:

1.  While it handles simple emphasis, fountain2odf.py does not handle
    complex emphasis correctly.  
      
    For example ‘This is some text in \*italics with embedded
    \*\*bold\*\* sweet\!\*’ would be interpreted as italics terminating
    in the first \* of \*\*bold, the \*\*of bold would be ignored as the
    first ‘\* would be consumed by the italics, ...  
      
    Reason: I’ve looked at the tradeoffs of the complexity of fixing
    this vs benefits to me and as I am unlikely to generate scripts
    containing these constructs have decided to live with it. There are
    work arounds, e.g. ‘This is some text in \*italics with embedded\*
    \*\*\*bold\*\*\* \*sweet\!\*’ should work
