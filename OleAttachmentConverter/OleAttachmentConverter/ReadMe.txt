========================================================================
    CONSOLE APPLICATION : OleAttachmentConverter Project Overview
========================================================================

AppWizard has created this OleAttachmentConverter application for you.

This file contains a summary of what you will find in each of the files that
make up your OleAttachmentConverter application.


OleAttachmentConverter.vcxproj
    This is the main project file for VC++ projects generated using an Application Wizard.
    It contains information about the version of Visual C++ that generated the file, and
    information about the platforms, configurations, and project features selected with the
    Application Wizard.

OleAttachmentConverter.vcxproj.filters
    This is the filters file for VC++ projects generated using an Application Wizard. 
    It contains information about the association between the files in your project 
    and the filters. This association is used in the IDE to show grouping of files with
    similar extensions under a specific node (for e.g. ".cpp" files are associated with the
    "Source Files" filter).

OleAttachmentConverter.cpp
    This is the main application source file.

/////////////////////////////////////////////////////////////////////////////
Other standard files:

StdAfx.h, StdAfx.cpp
    These files are used to build a precompiled header (PCH) file
    named OleAttachmentConverter.pch and a precompiled types file named StdAfx.obj.

/////////////////////////////////////////////////////////////////////////////
Other notes:

AppWizard uses "TODO:" comments to indicate parts of the source code you
should add to or customize.

/////////////////////////////////////////////////////////////////////////////


*** NOTE *** - This sample assumes that there is a message in the Inbox that is in RTF Format and has at least 1 image embedded in it.

Repro Steps for it's use:

1. Create a message in RTF format, and embed an IMAGE in it.  Either through Copy and Paste or Insert > Picture
2. Note the Subject
3. Send it to yourself
4. Run the code passing the arguments it needs. It should find the message and utilize it.
