#include <MsgBoxConstants.au3>
#include <Word.au3>
#include <Array.au3>
#include <Excel.au3>
#include <File.au3>

; define input parameter
$write_excel=0                         ; 0=no, 1=yes
$new_excel_list="\new_paperlist.xlsx"  ; list of paper in excel to write
$excel_list="\paperlist.xlsx"          ; list of paper in excel to reread
$sheet_name="Sheet1"
$column_name="A1:B49"


; -----------get filelist in ms word format------------------

; get filelist
Local $aFileList = _FileListToArray(@ScriptDir, "*")

; display result
;_ArrayDisplay($aFileList, "$aFileList")
If @error = 1 Then
        MsgBox($MB_SYSTEMMODAL, "", "Path was invalid.")
        Exit
 EndIf
 If @error = 4 Then
        MsgBox($MB_SYSTEMMODAL, "", "No file(s) were found.")
        Exit
	 EndIf

; ---------------write filelist to excel--------------------

If $write_excel = 1 Then

  ; Create application object and create a new workbook
  Local $oExcel = _Excel_Open()
  If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
  Local $oWorkbook = _Excel_BookNew($oExcel)
  If @error Then
    MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite", "Error creating the new workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    _Excel_Close($oExcel)
    Exit
  EndIf

  ; Write a 1D array to the active sheet in the active workbook
  _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $aFileList)
  If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite", "Error writing to worksheet." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
  Local $sWorkbook = @ScriptDir & $new_excel_list
  _Excel_BookSaveAs($oWorkbook, $sWorkbook)
  MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite", "excel successfully written.")

EndIf


; ---------read file conversion list from excel--------------

; Create application object and open an example workbook
Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & $excel_list)
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error opening workbook '" & $excel_file & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    _Excel_Close($oExcel)
    Exit
 EndIf

 ; Read the formulas of a cell range
Local $aResult = _Excel_RangeRead($oWorkbook, $sheet_name, $column_name)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 3", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
; MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 3", "Data successfully read." & @CRLF & "Please click 'OK' to display all formulas in column A.")
_ArrayDisplay($aResult, "Excel UDF: _Excel_RangeRead Example 3 - Formulas in column A")