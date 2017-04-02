Attribute VB_Name = "modLib"
Private Type BROWSEINFO                     'フォルダの選択で使用する構造体
   hwndOwner            As Long             '親Windowのハンドル
   pidlRoot             As Long             'ルートフォルダ
   pszDisplayName       As Long             '
   lpszTitle            As String           'ダイアログに表示するメッセージ
   ulFlags              As Long             'オプション
   lpfn                 As Long             '
   lParam               As Long             '
   iImage               As Long             '
End Type

Private Const MAX_PATH                          As Long = 260

'ルートフォルダ定数
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

Public Enum genmBrowsRootFolder                                 'ルートフォルダ列挙定数
    brfDesktop = CSIDL_DESKTOP                                  'デスクトップ
    brfPrograms = CSIDL_PROGRAMS                                'プログラム
    brfControls = CSIDL_CONTROLS                                'コントロールパネル
    brfPrinters = CSIDL_PRINTERS                                'プリンター
    brfPersonal = CSIDL_PERSONAL                                'My Documents
    brffavorites = CSIDL_FAVORITES                              'お気に入り
    brfStartUp = CSIDL_STARTUP                                  'スタートアップ
    brfRecent = CSIDL_RECENT                                    '最近使ったファイル
    brfSendTo = CSIDL_SENDTO                                    '送る
    brfBucket = CSIDL_BITBUCKET                                 'ごみ箱
    brfStartMenu = CSIDL_STARTMENU                              'スタートメニュー
    brfDesktopDir = CSIDL_DESKTOPDIRECTORY                      'デスクトップディレクトリ
    brfDrives = CSIDL_DRIVES                                    'ドライバ
    brfNetWork = CSIDL_NETWORK                                  'ネットワーク
    brfNetHoood = CSIDL_NETHOOD                                 'NetHood
    brfFonts = CSIDL_FONTS                                      'フォント
    brfTemplates = CSIDL_TEMPLATES                              'ShellNew
    brfCommonStartMenu = CSIDL_COMMON_STARTMENU                 '共通 - スタートメニュー
    brfCommonPrograms = CSIDL_COMMON_PROGRAMS                   '共通 - プログラム
    brfCommonStartup = CSIDL_COMMON_STARTUP                     '共通 - スタートアップ
    brfCommonDesktopDirectory = CSIDL_COMMON_DESKTOPDIRECTORY   '共通 - デスクトップディレクトリ
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

Public Enum genmBrowsFlags                                      'ブラウズフラグ列挙定数
    bffSysFolderOnly = BIF_RETURNONLYFSDIRS                     'システムフォルダのみ選択可能
    bffDomainHide = BIF_DONTGOBELOWDOMAIN                       'ドメイン非表示
    bffSysFileFolderOnly = BIF_RETURNFSANCESTORS                'システムファイルフォルダのみ選択可能
    bffForComputer = BIF_BROWSEFORCOMPUTER                      'コンピュータのみ選択可能
    bffForPrinter = BIF_BROWSEFORPRINTER                        'プリンタのみ選択可能
    bffIncludeFiles = BIF_BROWSEINCLUDEFILES                    'ファイルも表示する
End Enum

'フォルダの選択ダイアログを表示するAPI
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                (ByRef lpBROWSEINFO As BROWSEINFO) As Long

'SHBrowseForFolder で取得した値からパス名を取得するAPI
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                (ByVal pidl As Long, ByVal pszPath As String) As Long

'SHBrowseForFolder で取得した値のメモリを開放するAPI
Private Declare Function SHFree Lib "shell32" Alias "#195" _
                (ByVal pidl As Long) As Long


Public strOutputFolder出力先フォルダ As String
Public objGetFolder As Object



' 機能          : フォルダの選択ダイアログを表示してパス名を取得する
'
' 返り値        : 選択されたパス名（キャンセルもしくは無効なフォルダが選択された場合は空文字）
'
' 引数          : ARG1 - 親ウィンドウのハンドル
'                 ARG2 - 表示メッセージ
'                 ARG3 - ルートディレクトリ
'                 ARG4 - （省略可）ブラウズフラグ
'
Public Function getFolderName(ByVal vlngHWnd As Long, Optional ByVal vstrMessage As String = "フォルダを選択してください。", _
                                 Optional ByVal vlngRootDir As genmBrowsRootFolder = genmBrowsRootFolder.brfDesktop, _
                                 Optional ByVal vlngBrowsFlags As genmBrowsFlags = genmBrowsFlags.bffSysFolderOnly) As String

    Dim typBrowseInfo       As BROWSEINFO
    Dim lngFP               As Long
    Dim strPath             As String

    
    fBrowseForFolder = ""
    
    With typBrowseInfo
        .hwndOwner = vlngHWnd       '親ウィンドウ
        .lpszTitle = vstrMessage    'ダイアログに表示するメッセージ
        .pidlRoot = vlngRootDir     'ルートフォルダ
        .ulFlags = vlngBrowsFlags   'フラグ
    End With

    lngFP = SHBrowseForFolder(typBrowseInfo)        'フォルダの参照ダイアログを表示
    strPath = String$(MAX_PATH, vbNullChar)
    Call SHGetPathFromIDList(lngFP, strPath)        'フォルダのパスを取得
    Call SHFree(lngFP)                              'メモリを開放

    strOutputFolder = Left$(strPath, InStr(strPath, vbNullChar) - 1)
  getFolderName = strOutputFolder & "\"

End Function


Public Function main(strTargetSheetName, strEventId, dicTargetColumn) As String
'JSON出力フォルダ作成。
strJsonFolderName = modLib.getFolderName(0, "DXFファイルを書き込むフォルダを決めてください。")
strJsonFolderName = strJsonFolderName & Format(Date, "yy_mmdd") & Format(Time, "_hhmm_ss") 'フォルダ作成用
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


'UTF-8にてファイルを作成する。
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

'バイナリモードで開く
csvStream.Type = adTypeBinary
csvStream.Open


'BOM の後からデータをコピー
outStream.CopyTo csvStream

Dim fileName As String
fileName = strJsonFolderName & "\" & strEventId & "_" & strTargetSheetName & ".json"
csvStream.SaveToFile fileName, adSaveCreateOverWrite






csvStream.Close
outStream.Close

'処理結果を返す
main = strJsonFolderName
End Function


Public Function s() As String
'JSON出力フォルダ作成。
strJsonFolderName = modLib.getFolderName(0, "DXFファイルを書き込むフォルダを決めてください。")
strJsonFolderName = strJsonFolderName & Format(Date, "yy_mmdd") & Format(Time, "_hhmm_ss") 'ﾌｫﾙﾀﾞ作成用
'MkDir strCpgFolderName

MsgBox (strJsonFolderName)
End Function





