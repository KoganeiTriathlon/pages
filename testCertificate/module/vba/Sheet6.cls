VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
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

'�o�̓J���������`����B
'  1) �V�[�g��ԍ��� m �� hogehoge �Əo�͂���ꍇ�A
'    dicColumn.Add "col_m", "hogehoge"
'  �ƒǋL����B
'
'  2) �V�[�g��ԍ��� n �� dummy �Ƃ��Čv�Z�l�Ƃ��ė��p����ꍇ�A
'    dicColumn.Add "col_-n", "dummy"
'  �ƒǋL����B
'
'  3) �����v�Z���ɂ��o�͂������f�[�^�� dataX �ƒ�`���o�͂����ꍇ�A
'    dicColumn.Add "calc_0", "dataX"
'  �ƒǋL����B

dicColumn.Add "calc_0", "intRow"
dicColumn.Add "col_1", "strPlace"
dicColumn.Add "col_2", "intRaceNo"
dicColumn.Add "col_3", "strPlayerFullName"
dicColumn.Add "col_4", "intPlayerAge"
dicColumn.Add "col_5", "intPlayerGender"
dicColumn.Add "col_6", "strChildFullName"
dicColumn.Add "col_7", "strChildAge"
dicColumn.Add "col_8", "strGrade"
dicColumn.Add "col_9", "strChildGender"
dicColumn.Add "col_10", "strResidence"
dicColumn.Add "col_11", "strRecord"
dicColumn.Add "col_12", "strLapSwim"
dicColumn.Add "col_13", "intPlaceSwim"
dicColumn.Add "col_14", "strLapRun"
dicColumn.Add "col_15", "intPlaceRun"
dicColumn.Add "calc_1", "intPlace"
dicColumn.Add "calc_2", "intPlaceGender"

srtResponse = modLib.main(strThisSheetName, strEvent, dicColumn)

MsgBox ("Fin!!" & Chr(13) & srtResponse)

End Sub