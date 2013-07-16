 Dim strFilter As String
   Dim lngflags As Long
   Dim varFileName As Variant
   Dim strFileName As String
   Dim dirDivider As Long
   Dim varFileNameR As String
   Dim Uid As String
   Dim CurDate
   Dim CurTime
   
     CurTime = Time
    CurDate = Date
  
 Uid = Format(CurTime, "HHMMSS") & Format(CurDate, "MMDDYYYY")
   

   strFilter = "All Files (*.*)" & vbNullChar & "*.*" _
    & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*"
    
   lngflags = tscFNPathMustExist Or tscFNFileMustExist _
    Or tscFNHideReadOnly
   
   varFileName = tsGetFileFromUser( _
   fOpenFile:=True, _
   strFilter:=strFilter, _
   rlngflags:=lngflags, _
   strDialogTitle:="Please choose a file...")
   
   If IsNull(varFileName) Then
    Else
     'Determine filename
 varFileNameR = StrReverse(varFileName)
 dirDivider = (InStr(varFileNameR, "\")) - 1
 strFileName = Right(varFileName, dirDivider)
 Uid = Me![pID]
 
 
 'copy file to network
 Dim fs As Object
Dim oldPath As String, newPath As String
oldPath = varFileName
newPath = "\\rome\data\HANDS\NSP\images\"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.CopyFile oldPath, newPath & Uid & strFileName
Set fs = Nothing
      
      Me![PhotoLoc] = newPath & Uid & strFileName
      
     
      
   End If
   
   

cmdAddLarge_End:
   On Error GoTo 0
   Me.Requery
       DoCmd.GoToRecord , , acLast
       
   Exit Sub

cmdAddLarge_Err:
   Beep
   MsgBox Err.Description, , "Error: " & Err.Number _
    & " in file"
   Resume cmdAddLarge_End

End Sub
