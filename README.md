<div align="center">

## Callers Add\-in \[FINAL 2\.24 \- Sept 12, 2011\]


</div>

### Description

UPDATE 2.21: fixed fatal typo bug introduced in v2.2 ...

UPDATE 2.22: fixed 'exclude addin designer' introduced in v1.9 ...

UPDATE 2.23: fixed 'underscore' incorrect behavior introduced

in v1.8 (Initialize matched Class_Initialize) - thanks heaps

to Kenneth Ives (kenaso) for your feedback ...

UPDATE 2.24: fixed yet another bug - discovered the VBE Find

function searches up to but not including the specified last

line. Now correctly handles API declares with line continuation ...

DETAILS:

I borrowed code from Kamilche, Ulli and Darryl Hasieber to

create this quite simple VB add-in that adds something I've

sometimes needed as projects get too large and complex ...

It adds a couple of entries to VB's code pane context menu

so that when you right-click within a procedure (or on a

declare) and select 'Callers' you get a popup menu listing

all other routines in the project that call this code member ...

It displays which routines reference any particular code member

in the project, and allows you to select one to go to it - and

a second entry called 'Callee' to go back if you wish ...

FEATURES:

Callee's include all procedures, API declares, Enums and Types,

module level constants and variables, Implemented classes and

Raised Events. Also identifies parent Type or Enum when right-

clicking on one of their members. Finds public Enums of classes

without being qualified by class name. Also adds a 'Clear' menu

item to the Immediate Window ...

This will probably be the last update as it seems to work fairly

well now. However, bug reports or suggestions for improvement may

see some further updates ...

Just compile the project to automatically add the add-in to VB

and get two new menu items 'Callers' and 'Callee' on the code

context menu ... If you have compiled a previous version compile

again with no VB IDE's open by selecting 'Make' on the

Caller.vbp's context menu ... 16 kb zip.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2011-09-12 17:28:24
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Intermediate
**User Rating**    |5.0 (85 globes from 17 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Callers\_Ad2211429122011\.zip](https://github.com/Planet-Source-Code/rde-callers-add-in-final-2-24-sept-12-2011__1-73617/archive/master.zip)








