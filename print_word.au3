#include <MsgBoxConstants.au3>
#include <Word.au3>

; Create application object
Local $oWord = _Word_Create()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocPrint Example", _
        "Error creating a new Word application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

; Open the test document
Local $sDocument = @ScriptDir & "\Test2pdf withspace.docx"
Local $oDoc = _Word_DocOpen($oWord,$sDocument, Default, Default, True)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocPrint Example", _
        "Error opening file:" & @CRLF & "@error = " & @error & ", @extended = " & @extended)

; Export the complete document with default values
Local $sFileName = @ScriptDir & "\Test2pdf withspace.pdf"
;_Word_DocExport($oDoc, $sFileName)
_Word_DocExport($oDoc, $sFileName,Default,Default,Default,Default,False,Default,False,Default)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocExport Example", _
        "Error exporting the document." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
;ShellExecute($sFileName)
;MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocExport Example", _
;        "The whole document has successfully been exported to: " & $sFileName)

; close word
_Word_DocClose($oDoc)