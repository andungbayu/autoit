#include <ie.au3>

$URL="https://hris.simaster.ugm.ac.id/sdm/fo/index.php?mod=login_default&sub=login&act=view&typ=html"
$form="form-login"
$element_name="username"
$element_value="andung.geo"
$pass_name="password"
$pass_value="Geo12345"
$logbook_dir="https://hris.simaster.ugm.ac.id/sdm/fo/index.php?mod=logbook&sub=Logbook&act=View&typ=html"

; create web browser instance
$oIE = _IECreate($URL,0,1,1,1)

; selecting form
$oForm = _IEGetObjById($oIE, $form)

; get object by name and set value
$oCode = _IEFormElementGetObjByName($oForm, $element_name)
_IEFormElementSetValue($oCode, $element_value)

; get password by name and set value
$oCode = _IEFormElementGetObjByName($oForm, $pass_name)
_IEFormElementSetValue($oCode, $pass_value)

; click form submit button
_IEFormSubmit($oForm)
_IELoadWait($oIE)

; navigate to logbook directory
_IENavigate($oIE,$logbook_dir)
_IELoadWait($oIE)

; test filling variable
$oForm_lg = _IEGetObjById($oIE, "frmInput")
$oCode_lg = _IEFormElementGetObjByName($oForm_lg, "kualitas")
_IEFormElementSetValue($oCode_lg, "77")