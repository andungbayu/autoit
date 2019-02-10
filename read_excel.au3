#include <Array.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>

; define input parameter
$excel_file="C:\Users\justin\Dropbox\form_filling\testfile.xlsx"
$sheet_name="Sheet1"
$column_name="A1:A5"

; Create application object and open an example workbook
Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Local $oWorkbook = _Excel_BookOpen($oExcel, $excel_file)
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error opening workbook '" & $excel_file & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    _Excel_Close($oExcel)
    Exit
EndIf

; Read data from a single cell on the active sheet of the specified workbook
Local $sResult = _Excel_RangeRead($oWorkbook, $sheet_name, "A1")
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Data successfully read." & @CRLF & "Value of cell A1: " & $sResult)

; Read the formulas of a cell range
Local $aResult = _Excel_RangeRead($oWorkbook, $sheet_name, $column_name)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 3", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
; MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 3", "Data successfully read." & @CRLF & "Please click 'OK' to display all formulas in column A.")
; _ArrayDisplay($aResult, "Excel UDF: _Excel_RangeRead Example 3 - Formulas in column A")

; display data dimension
$iRows = UBound($aResult, $UBOUND_ROWS) ; Total number of rows. In this example it will be 10.
$iCols = UBound($aResult, $UBOUND_COLUMNS) ; Total number of columns. In this example it will be 20.
$iDimension = UBound($aResult, $UBOUND_DIMENSIONS) ; The dimension of the array e.g. 1/2/3 dimensional.
MsgBox($MB_SYSTEMMODAL, "", "The array is a " & $iDimension & " dimensional array with " & _
$iRows & " row(s) & " & $iCols & " column(s).")

; copy result to show display
$result=$aResult[1];
MsgBox($MB_SYSTEMMODAL, "","value of cell A2: " & $result)
