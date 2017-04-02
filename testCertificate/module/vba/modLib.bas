Attribute VB_Name = "modLib"
Private Type BROWSEINFO                     '�t�H���_�̑I���Ŏg�p����\����
   hwndOwner            As Long             '�eWindow�̃n���h��
   pidlRoot             As Long             '���[�g�t�H���_
   pszDisplayName       As Long             '
   lpszTitle            As String           '�_�C�A���O�ɕ\�����郁�b�Z�[�W
   ulFlags              As Long             '�I�v�V����
   lpfn                 As Long             '
   lParam               As Long             '
   iImage               As Long             '
End Type

Private Const MAX_PATH                          As Long = 260

'���[�g�t�H���_�萔
Private Const CSIDL_DESKTOP                     As Long = &H0
Private Const CSIDL_PROGRAMS                    As Long = &H2
Private Const CSIDL_CONTROLS                    As Long = &H3
Private Const CSIDL_PRINTERS                    As Long = &H4
Private Const CSIDL_PERSONAL                    As Long = &H5
Private Const CSIDL_FAVORITES                   As Long = &H6
Private Const CSIDL_STARTUP                     As Long = &H7
Private Const CSIDL_RECENT                      As Long = &H8
Private Const CSIDL_SENDTO                      As Long = &H9
Private Const CSIDL_BITBUCKET                   As Long = &HA
Private Const CSIDL_STARTMENU                   As Long = &HB
Private Const CSIDL_DESKTOPDIRECTORY            As Long = &H10
Private Const CSIDL_DRIVES                      As Long = &H11
Private Const CSIDL_NETWORK                     As Long = &H12
Private Const CSIDL_NETHOOD                     As Long = &H13
Private Const CSIDL_FONTS                       As Long = &H14
Private Const CSIDL_TEMPLATES                   As Long = &H15
Private Const CSIDL_COMMON_STARTMENU            As Long = &H16
Private Const CSIDL_COMMON_PROGRAMS             As Long = &H17
Private Const CSIDL_COMMON_STARTUP              As Long = &H18
Private Const CSIDL_COMMON_DESKTOPDIRECTORY     As Long = &H19
Private Const CSIDL_APPDATA                     As Long = &H1A
Private Const CSIDL_PRINTHOOD                   As Long = &H1B

Public Enum genmBrowsRootFolder                                 '���[�g�t�H���_�񋓒萔
    brfDesktop = CSIDL_DESKTOP                                  '�f�X�N�g�b�v
    brfPrograms = CSIDL_PROGRAMS                                '�v���O����
    brfControls = CSIDL_CONTROLS                                '�R���g���[���p�l��
    brfPrinters = CSIDL_PRINTERS                                '�v�����^�[
    brfPersonal = CSIDL_PERSONAL                                'My Documents
    brffavorites = CSIDL_FAVORITES                              '���C�ɓ���
    brfStartUp = CSIDL_STARTUP                                  '�X�^�[�g�A�b�v
    brfRecent = CSIDL_RECENT                                    '�ŋߎg�����t�@�C��
    brfSendTo = CSIDL_SENDTO                                    '����
    brfBucket = CSIDL_BITBUCKET                                 '���ݔ�
    brfStartMenu = CSIDL_STARTMENU                              '�X�^�[�g���j���[
    brfDesktopDir = CSIDL_DESKTOPDIRECTORY                      '�f�X�N�g�b�v�f�B���N�g��
    brfDrives = CSIDL_DRIVES                                    '�h���C�o
    brfNetWork = CSIDL_NETWORK                                  '�l�b�g���[�N
    brfNetHoood = CSIDL_NETHOOD                                 'NetHood
    brfFonts = CSIDL_FONTS                                      '�t�H���g
    brfTemplates = CSIDL_TEMPLATES                              'ShellNew
    brfCommonStartMenu = CSIDL_COMMON_STARTMENU                 '���� - �X�^�[�g���j���[
    brfCommonPrograms = CSIDL_COMMON_PROGRAMS                   '���� - �v���O����
    brfCommonStartup = CSIDL_COMMON_STARTUP                     '���� - �X�^�[�g�A�b�v
    brfCommonDesktopDirectory = CSIDL_COMMON_DESKTOPDIRECTORY   '���� - �f�X�N�g�b�v�f�B���N�g��
    brfAppData = CSIDL_APPDATA                                  'Application Data
    brfPrintHood = CSIDL_PRINTHOOD                              'PrintHood
End Enum

Private Const BIF_RETURNONLYFSDIRS   As Long = &H1              'For finding a folder to start document searching
Private Const BIF_DONTGOBELOWDOMAIN  As Long = &H2              'For starting the Find Computer
Private Const BIF_STATUSTEXT         As Long = &H4
Private Const BIF_RETURNFSANCESTORS  As Long = &H8
Private Const BIF_EDITBOX            As Long = &H10
Private Const BIF_VALIDATE           As Long = &H20             'insist on valid result (or CANCEL)
Private Const BIF_BROWSEFORCOMPUTER  As Long = &H1000           'Browsing for Computers.
Private Const BIF_BROWSEFORPRINTER   As Long = &H2000           'Browsing for Printers
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000           'Browsing for Everything

Public Enum genmBrowsFlags                                      '�u���E�Y�t���O�񋓒萔
    bffSysFolderOnly = BIF_RETURNONLYFSDIRS                     '�V�X�e���t�H���_�̂ݑI���\
    bffDomainHide = BIF_DONTGOBELOWDOMAIN                       '�h���C����\��
    bffSysFileFolderOnly = BIF_RETURNFSANCESTORS                '�V�X�e���t�@�C���t�H���_�̂ݑI���\
    bffForComputer = BIF_BROWSEFORCOMPUTER                      '�R���s���[�^�̂ݑI���\
    bffForPrinter = BIF_BROWSEFORPRINTER                        '�v�����^�̂ݑI���\
    bffIncludeFiles = BIF_BROWSEINCLUDEFILES                    '�t�@�C�����\������
End Enum

'�t�H���_�̑I���_�C�A���O��\������API
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                (ByRef lpBROWSEINFO As BROWSEINFO) As Long

'SHBrowseForFolder �Ŏ擾�����l����p�X�����擾����API
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                (ByVal pidl As Long, ByVal pszPath As String) As Long

'SHBrowseForFolder �Ŏ擾�����l�̃��������J������API
Private Declare Function SHFree Lib "shell32" Alias "#195" _
                (ByVal pidl As Long) As Long


Public strOutputFolder�o�͐�t�H���_ As String
Public objGetFolder As Object



' �@�\          : �t�H���_�̑I���_�C�A���O��\�����ăp�X�����擾����
'
' �Ԃ�l        : �I�����ꂽ�p�X���i�L�����Z���������͖����ȃt�H���_���I�����ꂽ�ꍇ�͋󕶎��j
'
' ����          : ARG1 - �e�E�B���h�E�̃n���h��
'                 ARG2 - �\�����b�Z�[�W
'                 ARG3 - ���[�g�f�B���N�g��
'                 ARG4 - �i�ȗ��j�u���E�Y�t���O
'
Public Function getFolderName(ByVal vlngHWnd As Long, Optional ByVal vstrMessage As String = "�t�H���_��I�����Ă��������B", _
                                 Optional ByVal vlngRootDir As genmBrowsRootFolder = genmBrowsRootFolder.brfDesktop, _
                                 Optional ByVal vlngBrowsFlags As genmBrowsFlags = genmBrowsFlags.bffSysFolderOnly) As String

    Dim typBrowseInfo       As BROWSEINFO
    Dim lngFP               As Long
    Dim strPath             As String

    
    fBrowseForFolder = ""
    
    With typBrowseInfo
        .hwndOwner = vlngHWnd       '�e�E�B���h�E
        .lpszTitle = vstrMessage    '�_�C�A���O�ɕ\�����郁�b�Z�[�W
        .pidlRoot = vlngRootDir     '���[�g�t�H���_
        .ulFlags = vlngBrowsFlags   '�t���O
    End With

    lngFP = SHBrowseForFolder(typBrowseInfo)        '�t�H���_�̎Q�ƃ_�C�A���O��\��
    strPath = String$(MAX_PATH, vbNullChar)
    Call SHGetPathFromIDList(lngFP, strPath)        '�t�H���_�̃p�X���擾
    Call SHFree(lngFP)                              '���������J��

    strOutputFolder = Left$(strPath, InStr(strPath, vbNullChar) - 1)
  getFolderName = strOutputFolder & "\"

End Function


Public Function main(strTargetSheetName, strEventId, dicTargetColumn) As String
'JSON�o�̓t�H���_�쐬�B
strJsonFolderName = modLib.getFolderName(0, "DXF�t�@�C�����������ރt�H���_�����߂Ă��������B")
strJsonFolderName = strJsonFolderName & Format(Date, "yy_mmdd") & Format(Time, "_hhmm_ss") '�t�H���_�쐬�p
MkDir strJsonFolderName

'StreamTypeEnum
Const adTypeBinary = 1
Const adTypeText = 2

'LineSeparatorsEnum
Const adCR = 13
Const adCRLF = -1
Const adLF = 10

'StreamWriteEnum
Const adWriteChar = 0
Const adWriteLine = 1

'SaveOptionsEnum
Const adSaveCreateNotExist = 1
Const adSaveCreateOverWrite = 2


'Dim objRange As Variant
Dim objRange As Range
'Set objRange = Range(Cells(1, 1)).CurrentRegion
Set objRange = Range("A1").CurrentRegion
'Set Rng = ActiveSheet.Range("A1").CurrentRegion

' MsgBox (UBound(objRange, 1) & "---" & UBound(objRange, 2))
'MsgBox (objRange.Rows.Count)


'UTF-8�ɂăt�@�C�����쐬����B
Dim outStream As Object
Set outStream = CreateObject("ADODB.Stream")

With outStream
    .Type = adTypeText
    .Charset = "UTF-8"
    .LineSeparator = adLF
End With

outStream.Open

outStream.WriteText strEventId & "_" & strTargetSheetName & "=[", adWriteLine
Select Case strTargetSheetName
Case "A", "B"
    For intR = 2 To objRange.Rows.Count - 1
        outStream.WriteText "    {", adWriteChar
        outStream.WriteText Chr(34) & "intRow" & Chr(34) & ":" & CInt(intR - 1) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlace" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 1)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intResult" & Chr(34) & ":" & CInt(IsNumeric(objRange(intR, 1))) * (-1) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intRaceNo" & Chr(34) & ":" & CInt(objRange(intR, 2)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlayerFullName" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 3)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlayerAge" & Chr(34) & ":" & CInt(objRange(intR, 4)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlayerGender" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 5)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strResidence" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 6)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strTotalRecord" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 7).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strLapSwim" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 8).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceSwim" & Chr(34) & ":" & CInt(objRange(intR, 9)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strLapRun" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 10).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceRun" & Chr(34) & ":" & CInt(objRange(intR, 11)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceMale" & Chr(34) & ":" & Val(objRange(intR, 12)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceFemale" & Chr(34) & ":" & Val(objRange(intR, 13)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceGender" & Chr(34) & ":" & CInt(WorksheetFunction.Max(objRange(intR, 12), objRange(intR, 13))), adWriteChar
        outStream.WriteText "},", adWriteLine
    Next
    outStream.WriteText "    {", adWriteChar
    outStream.WriteText Chr(34) & "intRow" & Chr(34) & ":" & CInt(intR - 1) & ",", adWriteChar
    outStream.WriteText Chr(34) & "strPlace" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 1)) & Chr(34) & ",", adWriteChar
    outStream.WriteText Chr(34) & "intResult" & Chr(34) & ":" & CInt(IsNumeric(objRange(intR, 1))) * (-1) & ",", adWriteChar
    outStream.WriteText Chr(34) & "intRaceNo" & Chr(34) & ":" & CInt(objRange(intR, 2)) & ",", adWriteChar
    outStream.WriteText Chr(34) & "strPlayerFullName" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 3)) & Chr(34) & ",", adWriteChar
    outStream.WriteText Chr(34) & "intPlayerAge" & Chr(34) & ":" & CInt(objRange(intR, 4)) & ",", adWriteChar
    outStream.WriteText Chr(34) & "strPlayerGender" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 5)) & Chr(34) & ",", adWriteChar
    outStream.WriteText Chr(34) & "strResidence" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 6)) & Chr(34) & ",", adWriteChar
    outStream.WriteText Chr(34) & "strTotalRecord" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 7).Text) & Chr(34) & ",", adWriteChar
    outStream.WriteText Chr(34) & "strLapSwim" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 8).Text) & Chr(34) & ",", adWriteChar
    outStream.WriteText Chr(34) & "intPlaceSwim" & Chr(34) & ":" & CInt(objRange(intR, 9)) & ",", adWriteChar
    outStream.WriteText Chr(34) & "strLapRun" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 10).Text) & Chr(34) & ",", adWriteChar
    outStream.WriteText Chr(34) & "intPlaceRun" & Chr(34) & ":" & CInt(objRange(intR, 11)) & ",", adWriteChar
    outStream.WriteText Chr(34) & "intPlaceMale" & Chr(34) & ":" & Val(objRange(intR, 12)) & ",", adWriteChar
    outStream.WriteText Chr(34) & "intPlaceFemale" & Chr(34) & ":" & Val(objRange(intR, 13)) & ",", adWriteChar
    outStream.WriteText Chr(34) & "intPlaceGender" & Chr(34) & ":" & CInt(WorksheetFunction.Max(objRange(intR, 12), objRange(intR, 13))), adWriteChar
    outStream.WriteText "}", adWriteLine

Case "CH", "CL"
    For intR = 2 To objRange.Rows.Count - 1
        outStream.WriteText "    {", adWriteChar
        outStream.WriteText Chr(34) & "intRow" & Chr(34) & ":" & CInt(intR - 1) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlace" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 1)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intResult" & Chr(34) & ":" & CInt(IsNumeric(objRange(intR, 1))) * (-1) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intRaceNo" & Chr(34) & ":" & CInt(objRange(intR, 2)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlayerFullName" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 3)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlayerAge" & Chr(34) & ":" & CInt(objRange(intR, 4)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlayerGrade" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 5)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlayerGender" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 6)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strResidence" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 7)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strTotalRecord" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 8).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strLapSwim" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 9).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceSwim" & Chr(34) & ":" & CInt(objRange(intR, 10)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strLapRun" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 11).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceRun" & Chr(34) & ":" & CInt(objRange(intR, 12)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceMale" & Chr(34) & ":" & Val(objRange(intR, 13)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceFemale" & Chr(34) & ":" & Val(objRange(intR, 14)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceGender" & Chr(34) & ":" & CInt(WorksheetFunction.Max(objRange(intR, 13), objRange(intR, 14))), adWriteChar
        outStream.WriteText "},", adWriteLine
    Next
    outStream.WriteText "    {", adWriteChar
        outStream.WriteText Chr(34) & "intRow" & Chr(34) & ":" & CInt(intR - 1) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlace" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 1)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intResult" & Chr(34) & ":" & CInt(IsNumeric(objRange(intR, 1))) * (-1) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intRaceNo" & Chr(34) & ":" & CInt(objRange(intR, 2)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlayerFullName" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 3)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlayerAge" & Chr(34) & ":" & CInt(objRange(intR, 4)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlayerGrade" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 5)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlayerGender" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 6)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strResidence" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 7)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strTotalRecord" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 8).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strLapSwim" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 9).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceSwim" & Chr(34) & ":" & CInt(objRange(intR, 10)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strLapRun" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 11).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceRun" & Chr(34) & ":" & CInt(objRange(intR, 12)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceMale" & Chr(34) & ":" & Val(objRange(intR, 13)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceFemale" & Chr(34) & ":" & Val(objRange(intR, 14)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceGender" & Chr(34) & ":" & CInt(WorksheetFunction.Max(objRange(intR, 13), objRange(intR, 14))), adWriteChar
    outStream.WriteText "}", adWriteLine

Case "PH", "PL"
    For intR = 2 To objRange.Rows.Count - 1
        outStream.WriteText "    {", adWriteChar
        outStream.WriteText Chr(34) & "intRow" & Chr(34) & ":" & CInt(intR - 1) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlace" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 1)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intResult" & Chr(34) & ":" & CInt(IsNumeric(objRange(intR, 1))) * (-1) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intRaceNo" & Chr(34) & ":" & CStr(objRange(intR, 2)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlayerFullName" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 3)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlayerAge" & Chr(34) & ":" & CInt(objRange(intR, 4)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlayerGender" & Chr(34) & ":" & CStr(objRange(intR, 5)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strChildFullName" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 6)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strChildAge" & Chr(34) & ":" & Chr(34) & CInt(objRange(intR, 7)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strGrade" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 8).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strChildGender" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 9).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strResidence" & Chr(34) & ":" & CStr(objRange(intR, 10)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strRecord" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 11).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strLapSwim" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 12).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceSwim" & Chr(34) & ":" & Val(objRange(intR, 13)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strLapRun" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 14).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceRun" & Chr(34) & ":" & Val(objRange(intR, 15)) & "", adWriteChar
        outStream.WriteText "},", adWriteLine
    Next
    outStream.WriteText "    {", adWriteChar
        outStream.WriteText Chr(34) & "intRow" & Chr(34) & ":" & CInt(intR - 1) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlace" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 1)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intResult" & Chr(34) & ":" & CInt(IsNumeric(objRange(intR, 1))) * (-1) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intRaceNo" & Chr(34) & ":" & CStr(objRange(intR, 2)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strPlayerFullName" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 3)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlayerAge" & Chr(34) & ":" & CInt(objRange(intR, 4)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlayerGender" & Chr(34) & ":" & CStr(objRange(intR, 5)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strChildFullName" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 6)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strChildAge" & Chr(34) & ":" & Chr(34) & CInt(objRange(intR, 7)) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strGrade" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 8).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strChildGender" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 9).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strResidence" & Chr(34) & ":" & CStr(objRange(intR, 10)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strRecord" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 11).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strLapSwim" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 12).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceSwim" & Chr(34) & ":" & Val(objRange(intR, 13)) & ",", adWriteChar
        outStream.WriteText Chr(34) & "strLapRun" & Chr(34) & ":" & Chr(34) & CStr(objRange(intR, 14).Text) & Chr(34) & ",", adWriteChar
        outStream.WriteText Chr(34) & "intPlaceRun" & Chr(34) & ":" & Val(objRange(intR, 15)) & "", adWriteChar
    outStream.WriteText "}", adWriteLine

End Select

outStream.WriteText "];", adWriteLine

outStream.Position = 0
outStream.Type = adTypeBinary
outStream.Position = 3


Dim csvStream As Object
Set csvStream = CreateObject("ADODB.Stream")

'�o�C�i�����[�h�ŊJ��
csvStream.Type = adTypeBinary
csvStream.Open


'BOM �̌ォ��f�[�^���R�s�[
outStream.CopyTo csvStream

Dim fileName As String
fileName = strJsonFolderName & "\" & strEventId & "_" & strTargetSheetName & ".json"
csvStream.SaveToFile fileName, adSaveCreateOverWrite






csvStream.Close
outStream.Close

'�������ʂ�Ԃ�
main = strJsonFolderName
End Function


Public Function s() As String
'JSON�o�̓t�H���_�쐬�B
strJsonFolderName = modLib.getFolderName(0, "DXF�t�@�C�����������ރt�H���_�����߂Ă��������B")
strJsonFolderName = strJsonFolderName & Format(Date, "yy_mmdd") & Format(Time, "_hhmm_ss") '̫��ލ쐬�p
'MkDir strCpgFolderName

MsgBox (strJsonFolderName)
End Function





