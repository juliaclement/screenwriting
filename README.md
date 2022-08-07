# **Script writing utilities**
1. ## **Scope and intents**
These are a number of small command line Python 3 utilities to carry out simple tasks to help reduce the impedance mismatch between my way of working and the requirements of scripts for stage and screen. In particular allowing scripts to be written in a full featured word processor (e.g. LibreOffice, Calligra Words, OpenOffice). I understand that Microsoft Office can export & import the popular ODT format so it would also be a candidate, but I run the GNOME desktop on Debian so an unable to run Microsoft Office to verify this. Trademarks respected.

They are placed here in case they are useful to others.
## **Language**
These documents and programs are in New Zealand English which does not use American spelling. If you don’t like reading correctly spelled New Zealand English, please don’t read them.
## **Licences**
Software: SAX is open source. The fountain2odf program depends heavily on the odfdo library which is licenced under the [Apache 2.0 licence](https://www.apache.org/licenses/LICENSE-2.0). Although this does not create an obligation, I have decided to use the Apache 2.0 licence for these tools programs . This licence is recognised as GPL compatible so shouldn’t create problems for people wanting to reuse or extend the utilities.

Documentation:  All original non-program files are licenced under the [Creative Commons Attribution-ShareAlike 4.0 International (CC BY-SA 4.0)](https://creativecommons.org/licenses/by-sa/4.0/) licence.
## **Contents**
Each utility has its own subdirectory with the utility, install instructions, and usage instructions in its README.

[Fountain2odf](http://fountain2odf/): Convert .fountain files to .odt.

[Odf2fountain](http://odf2fountain/): Convert .odt files formatted acording t our rules to .fountain files
## **To do:**
1. fountain2odf relies on having a template file with the styles we use pre-loaded. These should be generated on the fly when not in the template, allowing the template to be used for page size, header, footer, and boiler plate text.
1. Writing some tests is required.
## **Experimental**
None of these tools are intended to do the complete job, but I am prepared to consider requests for small changes that improve usability and accept pull requests for enhancements.

I have no current plans to write installers, or Pip packages.  For the two most popular operating systems I lack the means to test created installers. I would welcome requests from interested and technically capable people who would like to join the project to create installers for some or all platforms on an ongoing basis.

Some changes that can be justified, I am unwilling to do. examples:

1. While it handles simple emphasis, fountain2odf.py does not handle complex emphasis correctly. 

   For example ‘This is some text in \*italics with embedded \*\*bold\*\* sweet!\*’ would be interpreted as italics terminating in the first \* of \*\*bold, the \*\*of bold would be ignored as the first ‘\* would be consumed by the italics, ...

   Reason: I’ve looked at the tradeoffs of the complexity of fixing this vs benefits to me and as I am unlikely to generate scripts containing these constructs have decided to live with it. There are work arounds, e.g. ‘This is some text in \*italics with embedded\* \*\*\*bold\*\*\* \*sweet!\*’ should work
1. I’m not going to write a Fountain to PDF converter.

   Reason: Libreoffice comes with a commandline PDF writer. My LibreOffice originals can be directly exported by that software as can ODF files created by fountain2odf.py. To turn .fountain files into PDFs:
   `	`python3 fountain2odf.py my.fountain –template pageinfo.odt -o my\_temp.odt 
   `	`libreoffice --headless --convert-to pdf my\_temp.odt


Copyright © 2022 Julia Ingleby Clement.

You may use, modify and republish this document only under the terms of  the [Creative Commons Attribution-ShareAlike 4.0 International (CC BY-SA 4.0)](https://creativecommons.org/licenses/by-sa/4.0/) licence
