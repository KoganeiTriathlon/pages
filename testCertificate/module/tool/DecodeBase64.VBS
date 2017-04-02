' DecodeBase64.VBS V1.00 (C) yoshioka.teruo@nifty.com 2005-12-03
' DecodeBase64.VBS decodes file by Base64.
' Usage: Start DecodeBase64.VBS file [outfile]

Option Explicit
Dim fso
Dim sFile
Dim dFile
Dim ParentFolderName
Dim FileName
Dim BaseName
Dim ExtensionName

Set fso=CreateObject("Scripting.FileSystemObject") 

Select Case WScript.Arguments.Count
Case 2
  sFile=WScript.Arguments.Item(0)
  dFile=WScript.Arguments.Item(1)
Case 1
  sFile=WScript.Arguments.Item(0)
  ParentFolderName=fso.GetParentFolderName(sFile)
  FileName=fso.GetFileName(sFile)
  BaseName=fso.GetBaseName(FileName)
  ExtensionName=fso.GetExtensionName(FileName)
  If ExtensionName="" Then
    FileName=BaseName&".BIN"
  Else
    FileName=BaseName
  End If
  dFile=fso.BuildPath(ParentFolderName,FileName)
Case Else
  WScript.Echo "Usage: Start DecodeBase64.VBS file [outfile]"
  WScript.Quit
End Select

Const adTypeBinary=1 
Const adSaveCreateNotExist=1
Const adSaveCreateOverWrite=2
Dim Dom
Dim Tmp 
Dim Bin 
Dim Dst

Set Dom=CreateObject("Microsoft.XMLDOM") 
Set Tmp=Dom.createElement("tmp") 
Tmp.DataType="bin.base64" 
Tmp.Text=fso.OpenTextFile(sFile).ReadAll()
Bin=Tmp.NodeTypedValue 

Set Dst=CreateObject("ADODB.Stream") 
Dst.Open 
Dst.Type=adTypeBinary 
Dst.Write Bin 
Dst.SaveToFile dFile,adSaveCreateNotExist
Dst.Close 
