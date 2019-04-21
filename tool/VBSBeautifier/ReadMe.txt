------------------------------------
        VBScript Beautifier
 
          For Win9x/NT/2k

           By Niek Albers

     (C)2001-2007 DaanSystems

Homepage: http://www.daansystems.com
Email:    nieka@daansystems.com
------------------------------------

VBScript beautifier beautifies VBScript
files.

Features:

- Works on serverside and clientside VBScript.
- Skips HTML.
- Properizes keywords.
- Splits Dim statements.
- Places spaces around operators.
- Indents blocks.
- Lines out assignment statements.
- Removes redundant endlines.
- Makes backups.


Instructions:
-------------
It's all very straight forward. You can open any 
file that holds VBScript, or paste it into the editwindow,
then click on the Beautify button and see what happens!

IMPORTANT: Make sure the VBSCript code works before you
try to beautify it.

You can also use the commandline utility vbsbeaut.exe.
Enter vbsbeaut.exe without any arguments to see the commandline
options.

You can add vbscript keywords to the keywords.txt file.

26.Nov.2007 v1.10

  + VBSBeautifier is now freeware.
  + Fixed indenting of multip if then's on one line.

25.Jun.2002 v1.09

  Bugfixes:
  + Added a separated keywords_indent.txt file with keywords and their
    indentation behaviour.

31.Oct.2001 v1.07

  Bugfixes:
  + Forgot about the keywords.txt in the last path bug....
  + VBSBeautifier recognizes => and >= as comparing operators.
    (It was on >=).

31.Oct.2001 v1.06

  Bugfixes:
  + Fixed a path bug that caused registration information to
    get lost.

20.Sep.2001 v1.05

  Improvements:
  + Keywords are separated in another file: keyword.txt which
    can be modified.
  + Better handling of dos/unix type textfiles.
 
  Bugfixes:
  + Fixed insertion of extra newlines in certain loops.

27.Aug.2001 v1.04

  Bugfixes:
  + Indent problems with 'if then else' on one line.
  + No extra enters around if then else statements.
  + enters after <% are contained.
   
  Improvements:
  + Added an example asp file.
  + Some taint stuff.

7.Aug.2001 v1.03 Update

  Bugfixes:
  + Function header comments are not separated with an extra newline
    anymore bug fixed.

6.Aug.2001 v1.02 Update

  Improvements:
  + Function header comments are not separated with an extra newline
    anymore.

  Bugfixes:
  + clientside vbscript balance error when mixed with javascript.
  + 'On Error Resume Next' indendation balance error.
  + GUI stalled on very large files.

5.Aug.2001 v1.01 Update
 
  Improvements:
  + Option added to prevent splitting of Dim statements.

  Bugfixes:
  + Splitting out dim statement sometimes forgot to place a space between
    Dim and variable name.


4.Aug.2001 v1.0 release


25.Jul.2001 First Beta release

This is a beta version, and surely has some serious bugs.
Please send them to: nieka@daansystems.com


Known Bugs:

- VBScript beautifier will choke when you Place <% and %>
  between "" f.e.:
  Response.Write("%> %> <% <% %>")
 

