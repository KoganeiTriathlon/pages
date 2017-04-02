VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Dim strThisSheetName As String
Dim dicColumn As Object
Dim srtResponse As String
Dim strEvent As String


strEvent = "koganeiaqua2016"
strThisSheetName = ActiveWorkbook.ActiveSheet.Name
Set dicColumn = CreateObject("Scripting.Dictionary")

'出力カラム情報を定義する。
'  1) シート列番号が m を hogehoge と出力する場合、
'    dicColumn.Add "col_m", "hogehoge"
'  と追記する。
'
'  2) シート列番号が n を dummy として計算値として利用する場合、
'    dicColumn.Add "col_-n", "dummy"
'  と追記する。
'
'  3) 内部計算等により出力したいデータを dataX と定義し出力たい場合、
'    dicColumn.Add "calc_0", "dataX"
'  と追記する。

dicColumn.Add "calc_0", "intRow"
dicColumn.Add "col_1", "strPlace"
dicColumn.Add "col_2", "intRaceNo"
dicColumn.Add "col_3", "strPlayerFullName"
dicColumn.Add "col_4", "intPlayerAge"
dicColumn.Add "col_5", "intPlayerGender"
dicColumn.Add "col_6", "strResidence"
dicColumn.Add "col_7", "strRecord"
dicColumn.Add "col_8", "strLapSwim"
dicColumn.Add "col_9", "intPlaceSwim"
dicColumn.Add "col_10", "strLapRun"
dicColumn.Add "col_11", "intPlaceRun"
dicColumn.Add "col_-12", "intPlaceMale"
dicColumn.Add "col_-13", "intPlaceFemale"
dicColumn.Add "calc_1", "intPlace"
dicColumn.Add "calc_2", "intPlaceGender"

srtResponse = modLib.main(strThisSheetName, strEvent, dicColumn)

MsgBox ("Fin!!" & Chr(13) & srtResponse)

End Sub
