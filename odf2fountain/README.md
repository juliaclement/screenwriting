# **Odf2fountain – Open Document to Fountain text converter**
## **Purpose:**
Converts a word processing file in Oasis Open Document format (LibreOffice, etc) into a .fountain file unto a word processing file.
## **Licences:**
This program is licenced under the [Apache 2.0 licence](https://www.apache.org/licenses/LICENSE-2.0).  You may use it only under the terms of the Apache 2.0 licence. This licence is recognised as GPL compatible so shouldn’t create problems for people wanting to reuse or extend the utilities.

Documentation:  All original non-program files are licenced under the [Creative Commons Attribution-ShareAlike 4.0 International (CC BY-SA 4.0)](https://creativecommons.org/licenses/by-sa/4.0/) licence.
## **Input:**
Input files are standard LibreOffice writer documents formatted as a screenplay script. If you base your document on the documents in the [fountain2odf](https://github.com/juliaclement/screenwriting/tree/main/fountain2odf) you have styles named Action, Character, Dialogue, etc matching the fountain types with appropriate formatting. If you choose not to use these styles, you need to indicate the style to use by setting the capitalisation / number of leading spaces to let odt2fountain and the receiving translator guess what it is working on.
## **Usage:**
After installing, this is used from the command line:

`	`python3 odf2fountain.py [-h] [--output OUTPUT] [--forcetypes] [--extendedfountain] [--debug] file [files ...]

positional arguments:

`  `file/files	input files space separated

options:

`  `-h, --help	show help message and exit

`  `--output OUTPUT

`		`output filename. Default = input filename.fountain

`  `--forcetypes	Start lines with a special character from the [fountain specification](https://fountain.io/syntax) rather than let the fountain translator use heuristics to determine line type.

`  `--extendedfountain

`		`Enable extra type forcing characters not supported by most fountain translators. Currently just % for Dialogue

`  `--debug, -debug       provide developer information
### **Setup**
1. Install dependencies, see below
1. Download the program folder from github
## **Dependencies**
- The latest version of python3
## **Language**
These documents and programs are in New Zealand English which does not use American spelling. If you don’t like reading correctly spelled New Zealand English, please don’t read them.
## **Manifest:**
odf2fountain.py -- the program

README.md -- this file
## **To do:**
1. Writing some tests is required.
## **Experimental**
None of these tools are intended to do the complete job, but I am prepared to consider requests for small changes that improve usability and accept pull requests for enhancements.

I have no current plans to write installers, or Pip packages.  For the two most popular operating systems I lack the means to test created installers. I would welcome requests from interested and technically capable people who would like to join the project to create installers for some or all platforms on an ongoing basis.


Copyright © 2022 Julia Ingleby Clement.

You may use, modify and republish this document only under the terms of  the [Creative Commons Attribution-ShareAlike 4.0 International (CC BY-SA 4.0)](https://creativecommons.org/licenses/by-sa/4.0/) licence
