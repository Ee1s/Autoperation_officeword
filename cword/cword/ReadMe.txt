========================================================================
       MICROSOFT FOUNDATION CLASS LIBRARY : cword
========================================================================


AppWizard has created this cword application for you.  This application
not only demonstrates the basics of using the Microsoft Foundation classes
but is also a starting point for writing your application.

This file contains a summary of what you will find in each of the files that
make up your cword application.

cword.dsp
    This file (the project file) contains information at the project level and
    is used to build a single project or subproject. Other users can share the
    project (.dsp) file, but they should export the makefiles locally.

cword.h
    This is the main header file for the application.  It includes other
    project specific headers (including Resource.h) and declares the
    CCwordApp application class.

cword.cpp
    This is the main application source file that contains the application
    class CCwordApp.

cword.rc
    This is a listing of all of the Microsoft Windows resources that the
    program uses.  It includes the icons, bitmaps, and cursors that are stored
    in the RES subdirectory.  This file can be directly edited in Microsoft
	Visual C++.

cword.clw
    This file contains information used by ClassWizard to edit existing
    classes or add new classes.  ClassWizard also uses this file to store
    information needed to create and edit message maps and dialog data
    maps and to create prototype member functions.

res\cword.ico
    This is an icon file, which is used as the application's icon.  This
    icon is included by the main resource file cword.rc.

res\cword.rc2
    This file contains resources that are not edited by Microsoft 
	Visual C++.  You should place all resources not editable by
	the resource editor in this file.




/////////////////////////////////////////////////////////////////////////////

AppWizard creates one dialog class:

cwordDlg.h, cwordDlg.cpp - the dialog
    These files contain your CCwordDlg class.  This class defines
    the behavior of your application's main dialog.  The dialog's
    template is in cword.rc, which can be edited in Microsoft
	Visual C++.


/////////////////////////////////////////////////////////////////////////////
Other standard files:

StdAfx.h, StdAfx.cpp
    These files are used to build a precompiled header (PCH) file
    named cword.pch and a precompiled types file named StdAfx.obj.

Resource.h
    This is the standard header file, which defines new resource IDs.
    Microsoft Visual C++ reads and updates this file.

/////////////////////////////////////////////////////////////////////////////
Other notes:

AppWizard uses "TODO:" to indicate parts of the source code you
should add to or customize.

If your application uses MFC in a shared DLL, and your application is 
in a language other than the operating system's current language, you
will need to copy the corresponding localized resources MFC42XXX.DLL
from the Microsoft Visual C++ CD-ROM onto the system or system32 directory,
and rename it to be MFCLOC.DLL.  ("XXX" stands for the language abbreviation.
For example, MFC42DEU.DLL contains resources translated to German.)  If you
don't do this, some of the UI elements of your application will remain in the
language of the operating system.

/////////////////////////////////////////////////////////////////////////////
//////////////////////输入文字到另存为录制的宏///////////////////////////////


//    Selection.TypeText Text:="hello,test"
//    Selection.TypeParagraph
//    Selection.TypeText Text:="by wanglz;"
//    ChangeFileOpenDirectory "D:\"
//    ActiveDocument.SaveAs FileName:="hello.docx", FileFormat:= _
//        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
//        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
//        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
//        SaveAsAOCELetter:=False
//////////////////////插入图片的宏///////////////////////////////////////////	
	Sub picture()
'
' picture 宏
'
'
    Selection.InlineShapes.AddPicture FileName:= _
        "C:\Users\姜还涛\Pictures\doge.PNG", LinkToFile:=False, SaveWithDocument:= _
        True
End Sub	
//////////////////////插入表格2*3的宏///////////////////////////////////////////
Sub bg()
'
' bg 宏
'
'
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:= _
        3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "网格型" Then
            .Style = "网格型"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
    Selection.TypeText Text:="姓名"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="专业"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="学号"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="王露芝"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="计算机中美"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="2013012527"
End Sub
	Table tables = saveDoc.GetTables();
	Range rg =m_Sel.GetRange();	
