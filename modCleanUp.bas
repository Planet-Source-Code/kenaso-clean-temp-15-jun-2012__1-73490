Attribute VB_Name = "modCleanUp"
' ***************************************************************************
' Module:        modCleanUp
'
' Description:   This module will clean the Windows default TEMP folder
'                of *.TMP files that are not in use by another process.
'                The Recent Documents list will also be cleared and the
'                Recycle Bin emptied.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 03-Oct-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated EmptyWindowsTemp() routine to find and empty eligible
'              Windows TEMP folders
' 09-Apr-2012  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for pointing out a potential bug
'              by using a long name to designate a path instead of a short
'              name that may be misidentified by the system.
'              See EmptyWindowsTemp() routine.
' ***************************************************************************
Option Explicit

' ********************************************************************
' Constants
' ********************************************************************
  Private Const MAX_SIZE             As Long = 260
  Private Const S_OK                 As Long = 0
  Private Const SHERB_NOCONFIRMATION As Long = &H1    ' Recycle Bin flags
  Private Const SHERB_NOPROGRESSUI   As Long = &H2
  Private Const SHERB_NOSOUND        As Long = &H4

' ********************************************************************
' Type Structures
' ********************************************************************
  Private Type ULARGE_INTEGER
      LowPart  As Long
      HighPart As Long
  End Type

  Private Type SHQUERYRBINFO
      cbSize      As Long
      i64Size     As ULARGE_INTEGER
      i64NumItems As ULARGE_INTEGER
  End Type
   
' ********************************************************************
' Enumerations
' ********************************************************************
  Public Enum enumDriveType
      eUnknown           ' 0 Unknown drive type
      eBadRoot           ' 1 No root directory
      eRemovable         ' 2 Floppy, Flash, etc
      eFixed             ' 3 Local hard drive
      eNetwork           ' 4 Shared Network drive
      eCDRom             ' 5 CD-Rom drive (CD or DVD)
      eRamdisk           ' 6 Virtual memory disk
  End Enum

' ********************************************************************
' API Declares
' ********************************************************************
  ' The GetTempPath function retrieves the path of the directory designated
  ' for temporary files.  The GetTempPath function gets the temporary file
  ' path as follows:
  '   1.  The path specified by the TMP environment variable.
  '   2.  The path specified by the TEMP environment variable, if TMP
  '       is not defined.
  '   3.  The current directory, if both TMP and TEMP are not defined.
  Private Declare Function GetTempPath Lib "Kernel32.dll" _
          Alias "GetTempPathA" (ByVal nBufferLength As Long, _
          ByVal lpBuffer As String) As Long
            
  ' Empties the Recycle Bin on the specified drive.
  Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" _
          Alias "SHEmptyRecycleBinA" _
          (ByVal hwnd As Long, ByVal pszRootPath As String, _
          ByVal dwFlags As Long) As Long
  
  ' SHUpdateRecycleBinIcon updates the Recycle Bin icon on the
  ' desktop to reflect the state of the systemwide Recycle Bin.
  Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long

  ' SHQueryRecycleBin retrieves information about how many files
  ' (or other items) are currently in the Recycle Bin as well as
  ' how much disk space they consume.
  Private Declare Function SHQueryRecycleBin Lib "shell32.dll" _
          Alias "SHQueryRecycleBinA" _
          (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long

  ' Adds a document to the Shell's list of recently used documents
  ' or clears all documents from the list.
  Private Declare Function SHAddToRecentDocs Lib "shell32.dll" _
          (ByVal lFlags As Long, ByVal lPv As Long) As Long

  ' The GetLogicalDriveStrings function fills a buffer with strings that
  ' specify valid drives in the system.
  Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
          Alias "GetLogicalDriveStringsA" _
          (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
          
  ' The GetDriveType function determines whether a disk drive is a removable,
  ' fixed, CD-ROM, RAM disk, or network drive.
  Private Declare Function GetDriveType Lib "kernel32" _
          Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
          


' *************************************************************************** 
' ****                      Methods                                      **** 
' *************************************************************************** 

' ***************************************************************************
'  Routine:       EmptyRecycleBin
'
'  Description:   This routine will prompt the user to make sure all active
'                 applications have been stopped.  Then will verify the
'                 request before removing the TEMP folder.
'
'  Parameters:    None
'
'  Returns:       TRUE - Successful; FALSE - Unsuccessful
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Created routine
' 03/29/2006   Kenneth Ives  kenaso@tx.rr.com
'              Set the flag to produce no confirmation, sound or progress
' 03/30/2006   Kenneth Ives  kenaso@tx.rr.com
'              The Win32 help file is wrong.  You must supply the drive
'              letter to be able to capture any required information from
'              the Recycle Bin folder
' ***************************************************************************
Public Function EmptyRecycleBin() As Boolean

    Dim strDrives    As String
    Dim astrDrv()    As String
    Dim lngIndex     As Long
    Dim hwnd         As Long
    Dim lngFlag      As Long
    Dim lngDriveType As enumDriveType
    Dim typSHQBI     As SHQUERYRBINFO
    
    On Error GoTo EmptyRecycleBin_CleanUp
    
    ' set the option flags based on the
    lngFlag = SHERB_NOCONFIRMATION Or SHERB_NOPROGRESSUI Or SHERB_NOSOUND
    hwnd = Screen.ActiveForm.hwnd
    
    ' capture list of all logical drives
    strDrives = Space$(MAX_SIZE)             ' pad buffer to hold the results
    lngIndex = GetLogicalDriveStrings(MAX_SIZE, strDrives)
    strDrives = Left$(strDrives, lngIndex)   ' remove any trailing garbage
    astrDrv() = Split(strDrives, Chr$(0))    ' store data in an array
    
    ' Cycle thru all the fixed drive letters to
    ' make sure the Recycle Bin has been emptied.
    For lngIndex = 0 To UBound(astrDrv) - 1
    
        ' Get the drive number constant
        lngDriveType = GetDriveType(astrDrv(lngIndex))
    
        ' we only want the local fixed drives
        Select Case lngDriveType
               
               Case eFixed
                    ' set the length of the structure
                    typSHQBI.cbSize = LenB(typSHQBI)
                    
                    ' must contain a valid drive letter. Ex:  "C:\"
                    SHQueryRecycleBin astrDrv(lngIndex), typSHQBI
                    
                    ' If the recycle bin is empty and you attempt to
                    ' empty it, you will get a warning prompt that
                    ' something cannot be accessed.  This is internal
                    ' to Windows.  Make sure there is something to
                    ' remove first.
                    If (typSHQBI.i64NumItems.LowPart + _
                        typSHQBI.i64NumItems.HighPart) > 0 Then
                        
                        ' You do not need a drive letter here to
                        ' empty all the recycle bins
                        SHEmptyRecycleBin hwnd, vbNullString, lngFlag
                        
                        ' Update the recycle Bin icon on the desktop
                        SHUpdateRecycleBinIcon
                        EmptyRecycleBin = True
                    
                    Else
                        EmptyRecycleBin = True
                    End If
        End Select
    
    Next lngIndex
    
EmptyRecycleBin_CleanUp:
    Err.Clear
    Erase astrDrv()
    On Error GoTo 0
    
End Function

' ***************************************************************************
'  Routine:       EmptyMostRecent
'
'  Description:   This routine will prompt the user to make sure all active
'                 applications have been stopped.  Then will verify the
'                 request before emptying the TEMP folder.
'
'  Parameters:    None
'
'  Returns:       TRUE - Successful; FALSE - Unsuccessful
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Created routine
' ***************************************************************************
Public Function EmptyMostRecent() As Boolean

    On Error Resume Next   ' ignore any errors
    
    If SHAddToRecentDocs(0, 0) = S_OK Then
        EmptyMostRecent = True
    Else
        EmptyMostRecent = False
    End If
    
    On Error GoTo 0        ' nullify error routine
    
End Function

' ***************************************************************************
'  Routine:       EmptyWindowsTemp
'
'  Description:   This routine will prompt the user to make sure all active
'                 applications have been stopped.  Then will verify the
'                 request before emptying the TEMP folder.
'
'  Returns:       TRUE  - Successful
'                 FALSE - Unsuccessful
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Created routine
' 09-Apr-2012  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for pointing out a potential bug
'              by using a long name to designate a path instead of a short
'              name that may be misidentified by the system.
' ***************************************************************************
Public Function EmptyWindowsTemp() As Boolean

    Dim lngLength   As Long
    Dim strPath     As String
    Dim strLongName As String
    
    On Error Resume Next        ' ignore any errors
    
    EmptyWindowsTemp = False    ' Preset to FALSE
    strPath = Space$(MAX_SIZE)  ' Preload with spaces
    
    ' Determine the Windows temporary work path.
    ' Will sometimes return data in DOS 8.3 format.
    ' Ex:  C:\DOCUME~1\OWNER~1.KEN\LOCALS~1\Temp\
    lngLength = GetTempPath(MAX_SIZE, strPath)   ' Capture Temp folder location
    strPath = Left$(strPath, lngLength)          ' Capture path
    strPath = UnQualifyPath(Trim$(strPath))      ' Remove traiiing backslash
    
    ' Is this a TEMP folder?
    If StrComp(Right$(strPath, 5), "\temp", vbTextCompare) = 0 Then
        
        strPath = QualifyPath(strPath)   ' Add trailing backslash
        
        ' Convert 8.3 path name to a legible name
        ' ex:  C:\Documents And Settings\Owner.Kenaso\Local Settings\Temp\
        strLongName = QualifyPath(GetLongName(strPath))
        
        ' Locate first backslash after drive letter (C:\)
        lngLength = InStr(4, strLongName, "\")
        
        If lngLength > 0 Then
        
            ' 09-Apr-2012 Use complete long name instead of a short
            ' Verify this path exist
            If IsPathValid(strLongName) Then
                EmptyFolder strLongName   ' Empty temp folder
                EmptyWindowsTemp = True   ' Good finish
            End If
        End If
    End If

EmptyWindowsTemp_CleanUp:
    On Error GoTo 0       ' nullify error routine
    
End Function



' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

' ***************************************************************************
'  Routine:       EmptyFolder
'
'  Description:   Deletes an entire directory tree reguardless of the file
'                 attributes of the specified directory's contents.  It
'                 effectively deletes an entire directory tree regardless
'                 of the attributes of the files and directories it contains.
'                 In this case, the Temp folder is not removed because some
'                 of its contents may still be active.
'
'  Parameters:    strPathToDelete - Full pathname of the directory tree to
'                      be removed.
'
'  Returns:       None.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Created routine
' 05-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug.  While in the VB IDE, temp files created by VB
'              were being deleted and causing a system fault.
' ***************************************************************************
Private Sub EmptyFolder(ByVal strSearchPath As String)

    Dim strPath   As String                      ' current folder or filename
    Dim strTemp   As String                      ' temp holding area
    Dim lngIndex  As Long                        ' loop index
    Dim lngCount  As Long                        ' number of items in collection
    Dim colFiles  As Collection                  ' collection of folders and files
    Dim objFSO    As Scripting.FileSystemObject
    
    ' There are some files that are being accessed
    ' by the Windows enviroment and cannot be deleted.
    ' These will be ignored.  We will move on to the
    ' next file.
    On Error Resume Next

    Set colFiles = New Collection                ' Instantiate collection
    Set objFSO = New Scripting.FileSystemObject  ' Instantiate scripting object
    lngCount = 0                                 ' zero counter
    
    ' Add trailing backslash if missing
    strSearchPath = QualifyPath(strSearchPath)
    
    ' Get a list of path & filenames from this folder
    strPath = Dir$(strSearchPath & "*.*", vbNormal Or vbReadOnly Or _
                                          vbHidden Or vbSystem Or _
                                          vbArchive Or vbDirectory)
    ' Loop thru the directory structure
    ' and add the path & filename to
    ' the collection.
    Do While Len(strPath) > 0
        
        If (strPath <> ".") And (strPath <> "..") Then
            
            ' Are we in the VB development enviroment?
            If gblnIDE_Environment Then
            
                ' While in the IDE do not delete any file
                ' beginning with "VB" as this may cause
                ' a system fault and force VB to close
                ' and you will lose any changes you have
                ' not saved.
                
                strTemp = objFSO.GetFileName(strPath)   ' capture file name
                
                ' If file begins with with "vb" and
                ' ends with "tmp" then ignore it
                If LCase$(strTemp) Like "vb*tmp" Then
                    
                    ' This may be one of the files that will
                    ' give us problems.  Get the next file.
                    DoEvents
                    
                Else
                    colFiles.Add strSearchPath & strPath  ' Add to collection
                End If
            Else
                ' add path/filename to collection
                colFiles.Add strSearchPath & strPath
            End If
            
        End If
        
        ' Is there anything left?
        ' Parameters from above are retained.
        strPath = Dir$()
        DoEvents
              
    Loop
    
    lngCount = colFiles.Count
    
    ' Parse backwards thru collection and delete data.
    ' Backwards parsing prevents a collection from
    ' having to reindex itself after each data removal.
    If lngCount > 0 Then
        For lngIndex = lngCount To 1 Step -1
            
            ' move the path & filename to a variable
            strPath = colFiles(lngIndex)
            
            ' See if it is a directory.
            If GetAttr(strPath) And vbDirectory Then
                SetAttr strPath, vbNormal          ' reset the attributes to normal
                EmptyFolder strPath                ' make a cursive call on this folder
                objFSO.DeleteFolder strPath, True  ' delete the folder
            Else
                ' It's a file. Delete it.
                SetAttr strPath, vbNormal          ' reset file attributes to normal
                objFSO.DeleteFile strPath, True    ' delete file
            End If
            
        Next lngIndex
    End If
    
    Set colFiles = Nothing   ' Always free objects from memory when not needed
    Set objFSO = Nothing
    On Error GoTo 0          ' Nullify error trap in this routine

End Sub

