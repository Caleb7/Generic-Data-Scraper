#cs ----------------------------------------------------------------------------
 AutoIt Version: 3.3.14.2
 Author: Me
 Script Function:
#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ProgressConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <IE.au3>
#include <ColorConstants.au3>
#include <Array.au3>
#Include <GUIEdit.au3>
#Include <ScrollBarConstants.au3>
#Region ### START Koda GUI section ### Form=
$Form1_1 = GUICreate("Generic Data Scraper", 636, 497, 184, 123)
$BTN_START = GUICtrlCreateButton("Start", 8, 8, 75, 25)
$BTN_STOP = GUICtrlCreateButton("Stop", 8, 40, 75, 25)
$EDIT_LOG = GUICtrlCreateEdit("", 88, 8, 537, 433, $ES_AUTOVSCROLL + $WS_VSCROLL + $ES_READONLY + $ES_MULTILINE)
$GROUP_INFO = GUICtrlCreateGroup("Info", 8, 72, 73, 121)
$Label1 = GUICtrlCreateLabel("URL:", 16, 88, 60, 17)
$LABEL_COMPLETED = GUICtrlCreateLabel("0", 16, 104, 60, 17)
$Label3 = GUICtrlCreateLabel("Found:", 16, 120, 60, 17)
$LABEL_FOUND = GUICtrlCreateLabel("0", 16, 136, 60, 17)
$Label5 = GUICtrlCreateLabel("Status:", 16, 152, 60, 17)
$LABEL_STATUS = GUICtrlCreateLabel("Off", 16, 168, 60, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Progress1 = GUICtrlCreateProgress(8, 448, 622, 17)
$Progress2 = GUICtrlCreateProgress(8, 472, 622, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global $URLS = LoadURLFile( @ScriptDir & "/urls.txt" )

Global $PROGRESS_STEP = 0
Global $PROGRESS_STEPS = 10
Global $URLS_TOTAL = UBound($URLS)
Global $SCRAPER_ON = 0

Global $URLS_COMPLETED = 0
Global $EMAILS_FOUND = 0

SetCompletedNumber(0, $COLOR_RED)
SetFoundNumber(0, $COLOR_RED)

While 1

	$GUIMSG = GUIGetMsg()
	Switch $GUIMSG
		Case $GUI_EVENT_CLOSE
			dfShutdown()
			Exit
		Case $BTN_START
			dfStart()
		Case $BTN_STOP
			dfStop()
	EndSwitch

	dfRun()
WEnd

Func dfGetSource( $sURL )
	;dfLogIt("GET: " & $sURL)
	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
	$oHTTP.Open("GET", $sURL, False)
	If @error Then
		;dfLogIt("Failed to fetch URL: " & $sURL)
		Return "-1"
	EndIf
	$oHTTP.Send()
	If @error Then
		If $oHTTP.Status <> 200 Then
			;dfLogIt("Failed to fetch URL: " & $sURL)
			Return "-1"
		EndIf
	EndIf
	Return $oHTTP.ResponseText

EndFunc

Func dfRun()
	If $SCRAPER_ON == 1 Then
		If UBound($URLS) >= 1 Then
			Local $Data = dfGetSource($URLS[0])
			If $Data <> "-1" Then
				; Do stuff here

				Local $Found = StringRegExp( $Data, "[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?", 3 )

				; Local $FoundString = _ArrayToString( $Found, " " )
				; Perform additional regex to remove stupid@1.2.3 crap and //logurl/blah@1.2.3

				If UBound($Found) >= 1 Then
					$EMAILS_FOUND += UBound($Found)
					SetFoundNumber($EMAILS_FOUND, $COLOR_GREEN)
					dfLogIt( _ArrayToString($Found, @CRLF) )
				Else
					dfLogIt("No emails found for this URL")
				EndIf
				$URLS_COMPLETED += 1
				SetCompletedNumber($URLS_COMPLETED, $COLOR_RED)
				_ArrayDelete($URLS, 0)
			Else
				dfLogIt("Waiting...")
				Sleep(5000)
			EndIf

			GUICtrlSetData($Progress2, ($URLS_COMPLETED / $URLS_TOTAL) * 100)
		EndIf
	EndIf
EndFunc

Func dfStart()
	$SCRAPER_ON = 1
	GUICtrlSetState( $BTN_START, $GUI_DISABLE )
	GUICtrlSetState( $BTN_STOP, $GUI_ENABLE )
	SetStatus("on", $COLOR_GREEN)
EndFunc

Func dfStop()
	$SCRAPER_ON = 0
	GUICtrlSetState( $BTN_START, $GUI_ENABLE )
	GUICtrlSetState( $BTN_STOP, $GUI_DISABLE )
	SetStatus("off", $COLOR_RED)
EndFunc

Func dfShutdown()

EndFunc

Func SetCompletedNumber($NUMBER, $COLOR)
	GUICtrlSetData($LABEL_COMPLETED, $NUMBER & " / " & $URLS_TOTAL)
	GUICtrlSetColor($LABEL_COMPLETED, $COLOR)
EndFunc

Func SetFoundNumber($NUMBER, $COLOR)
	GUICtrlSetData($LABEL_FOUND, $NUMBER)
	GUICtrlSetColor($LABEL_FOUND, $COLOR)
EndFunc

Func SetStatus($STATUS, $COLOR)
	GUICtrlSetData($LABEL_STATUS, $STATUS)
	GUICtrlSetColor($LABEL_STATUS, $COLOR)
EndFunc

Func dfLogIt($MSG, $NewLine = 1)
	If $NewLine Then
		GUICtrlSetData( $EDIT_LOG, GUICtrlRead( $EDIT_LOG ) & $MSG & @CRLF )
	Else
		GUICtrlSetData( $EDIT_LOG, GUICtrlRead( $EDIT_LOG ) & $MSG )
	EndIf
	$iEnd = StringLen(GUICtrlRead($EDIT_LOG))
	_GUICtrlEdit_SetSel($EDIT_LOG, $iEnd, $iEnd)
	_GUICtrlEdit_Scroll($EDIT_LOG, $SB_SCROLLCARET)
EndFunc

Func LoadURLFile($PATH)
	Local $URLFile = FileOpen($PATH)
	Return FileReadToArray($URLFile)
EndFunc



