#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.
;#AutoIt3Wrapper_Change2CUI=y
#include-once
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
Global $__use_convert_au3 = 1
#include <GuiEdit.au3>

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiComboBox.au3>
#include "COMMs.au3"



TCPStartUp()

Global Enum $connectNone = 0, $connectCOM, $connectLAN, $connectVISA
Global Enum $useCOM = 0, $useLAN, $useVISA
Global $cMaxEditLen = 25000
Global $Terminal, $inputTX, $termMain

Global $COMlist[200], $COMcount = 0, $curCOM = 0
Global $hCOM
Global $history[200],$histFirst = 0, $histLast = 0, $histCur = 0, $noHist = true
Global $histLastString = ""
Global $ConOpen = $connectNone
Global $editTXcount = 0, $editRXcount = 0

Global $RXhead[2], $RXheadEcho, $RXtail
Global $RXheadStr, $RXTailStr

Global $hLog, $logEnabled = 0, $logFileSelected = 0

Global $macroDirectory = @WorkingDir, $logDirectory = @WorkingDir

Global Const $REG_ROOT = "HKEY_CURRENT_USER\Software\SRLabs\Terminal"

Global $hVISAdll = 0
;~ $hVISAdll = "visa32.dll"
;~ $hVISAdll = DllOpen("visa32.dll")
#include "visa_functions.au3"
;~ #include <Visa.au3>

Global Const $MACRO_NUMBER = 27
Global $iMcr[$MACRO_NUMBER], $iMcrP[2][$MACRO_NUMBER], $bMcrSend[$MACRO_NUMBER], $iMcrRT[$MACRO_NUMBER], $checkMcrRsend[$MACRO_NUMBER]

Global $mcrRepeat = False, $mcrRptCur = 0

Global $ShowMacros = 0

Global $VISA32_AVAILABLE = 0
$hVISAdll = DllOpen("visa32.dll")

If $hVISAdll <> -1 Then
    DLLclose($hVISAdll)
    $VISA32_AVAILABLE = 1
EndIf

$DUMMY_FILL = 0

Opt("GUIResizeMode",802)
;~ HotKeySet("{UP}", "captureUP")
Func captureUP()
    If WinGetHandle("") <> $Terminal Then
	HotKeySet("{UP}")
	Send("{UP}")
	HotKeySet("{UP}", "captureUP")
	Return
    EndIf
    Local $a
    $a = ControlGetHandle ($Terminal, "", ControlGetFocus($Terminal))
    $a = _WinAPI_GetDlgCtrlID($a)
    If $a <> $inputTX Then
	HotKeySet("{UP}")
	Send("{UP}")
	HotKeySet("{UP}", "captureUP")
	Return
    EndIf
;~     ConsoleWrite("UP!!!" & @CRLF)
    historyPrev()
EndFunc



;~ HotKeySet("{DOWN}", "captureDOWN")
Func captureDOWN()
    If WinGetHandle("") <> $Terminal Then
	HotKeySet("{DOWN}")
	Send("{DOWN}")
	HotKeySet("{DOWN}", "captureDOWN")
	Return
    EndIf
    Local $a
    $a = ControlGetHandle ($Terminal, "", ControlGetFocus($Terminal))
    $a = _WinAPI_GetDlgCtrlID($a)
    If $a <> $inputTX Then
	HotKeySet("{DOWN}")
	Send("{DOWN}")
	HotKeySet("{DOWN}", "captureDOWN")
	Return
    EndIf
;~     ConsoleWrite("DOWN!!!" & @CRLF)
    historyNext()
EndFunc


;~ HotKeySet("^{DEL}", "clearRXbuffer") ; CTRL+DEL clears RX input box


;~ HotKeySet("{ENTER}", "captureENTER")
Func captureENTER()
    If WinGetHandle("") <> $Terminal Then
	HotKeySet("{ENTER}")
	Send("{ENTER}")
	HotKeySet("{ENTER}", "captureENTER")
	Return
    EndIf
    Local $a
    $a = ControlGetHandle ($Terminal, "", ControlGetFocus($Terminal))
    $a = _WinAPI_GetDlgCtrlID($a)
    If $a = $inputTX Then
;~     ConsoleWrite("ENTER!!!" & @CRLF)
	sendInputData()
	Return
    EndIf
    For $i = 0 To $MACRO_NUMBER - 1
	If $a = $iMcr[$i] Then
	    macroSend($i)
	    Return
	EndIf
    Next
    HotKeySet("{ENTER}")
    Send("{ENTER}")
    HotKeySet("{ENTER}", "captureENTER")
    Return

EndFunc

#Region ### START Koda GUI section ### Form=D:\scripts\Terminal\Terminal2.kxf
$Terminal = GUICreate("SRLabs Terminal - " & FileGetVersion(@ScriptFullPath) , 654, 666, -1, -1, -1, -1)
$Connection = GUICtrlCreateGroup("", 0, 0, 125, 65)
$rCOM = GUICtrlCreateRadio("COM", 4, 12, 41, 17)
$rLAN = GUICtrlCreateRadio("LAN", 4, 28, 45, 17)
$rVISA = GUICtrlCreateRadio("VXI", 4, 44, 45, 17)
$bConnect = GUICtrlCreateButton("Connect", 52, 12, 65, 17)
$bScanCOM = GUICtrlCreateButton("Scan", 52, 40, 65, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gCOMset = GUICtrlCreateGroup("", 128, 0, 497, 65)
$gCOMPort = GUICtrlCreateGroup("COM port", 132, 8, 89, 53)
$cCOM = GUICtrlCreateCombo("", 140, 32, 73, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gBaudrate = GUICtrlCreateGroup("Baudrate", 224, 8, 77, 53)
$cBaud = GUICtrlCreateCombo("", 232, 32, 65, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "600|1200|2400|4800|9600|14400|19200|28800|38400|56000|57600|115200|128000|256000", "115200")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gData = GUICtrlCreateGroup("Data bits", 304, 8, 57, 53)
$cData = GUICtrlCreateCombo("", 312, 32, 41, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "5|6|7|8", "8")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gParity = GUICtrlCreateGroup("Parity", 364, 8, 65, 53)
$cParity = GUICtrlCreateCombo("", 372, 32, 49, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "none|odd|even|mark|space", "none")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gStopB = GUICtrlCreateGroup("Stop bits", 432, 8, 57, 53)
$cStopbits = GUICtrlCreateCombo("", 440, 32, 41, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "1|1.5|2", "1")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gHandshake = GUICtrlCreateGroup("Handshake", 492, 8, 129, 53)
$cHandShake = GUICtrlCreateCombo("", 500, 32, 113, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "none|RTS/CTS|XON/XOFF|RTS/CTS+XON/XOFF", "none")
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlSetState(-1, $GUI_DISABLE)
$gLANsettings = GUICtrlCreateGroup("", 128, 0, 309, 65)
$gIP = GUICtrlCreateGroup("IP address", 132, 8, 201, 53)
$iIP = GUICtrlCreateInput("", 140, 32, 185, 21)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gLANBport = GUICtrlCreateGroup("Port", 336, 8, 97, 53)
$iPort = GUICtrlCreateInput("", 344, 32, 81, 21)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlSetState(-1, $GUI_DISABLE)
$gVISAset = GUICtrlCreateGroup("", 128, 0, 349, 65)
$gVISAaddr = GUICtrlCreateGroup("VXI address", 132, 8, 341, 53)
$iVISAaddr = GUICtrlCreateInput("", 140, 32, 329, 21)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlSetState(-1, $GUI_DISABLE)
$gMacros = GUICtrlCreateGroup("Macros", 548, 64, 105, 69)
$bMacroWindow = GUICtrlCreateButton("Macro Window ->", 556, 84, 89, 17)
$bLoadMacro = GUICtrlCreateButton("Load", 556, 108, 41, 17)
$bSaveMacro = GUICtrlCreateButton("Save", 604, 108, 41, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$editTX = GUICtrlCreateEdit("", 0, 552, 653, 113, BitOR($ES_AUTOVSCROLL,$ES_AUTOHSCROLL,$ES_READONLY,$WS_VSCROLL))
GUICtrlSetFont(-1, 9, 400, 0, "Courier New")
GUICtrlSetColor(-1, 0x000000)
$editRX = GUICtrlCreateEdit("", 0, 136, 653, 389, BitOR($ES_AUTOVSCROLL,$ES_AUTOHSCROLL,$ES_READONLY,$WS_VSCROLL))
GUICtrlSetFont(-1, 9, 400, 0, "Courier New")
$checkTX_CR = GUICtrlCreateCheckbox("+CR", 536, 530, 41, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$checkCRLF = GUICtrlCreateCheckbox("CR=CR+LF", 580, 530, 73, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$inputTX = GUICtrlCreateInput("", 0, 528, 529, 21)
$gRXfilter = GUICtrlCreateGroup("Receive head / tail filter", 0, 64, 185, 69)
$InHead = GUICtrlCreateInput("", 40, 80, 137, 21)
$InTail = GUICtrlCreateInput("", 40, 106, 137, 21)
$checkEnableRXfilter = GUICtrlCreateCheckbox("On", 8, 90, 29, 29, BitOR($GUI_SS_DEFAULT_CHECKBOX,$BS_MULTILINE))
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gLog = GUICtrlCreateGroup("Logging", 464, 64, 81, 69)
$bSelectLog = GUICtrlCreateButton("Select file", 472, 84, 65, 17)
$bStartLog = GUICtrlCreateButton("Start", 472, 108, 29, 17)
$bStopLog = GUICtrlCreateButton("Stop", 508, 108, 29, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group1 = GUICtrlCreateGroup("", 188, 64, 273, 69)
$cClearRX = GUICtrlCreateButton("Clear RX buffer", 196, 84, 81, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

; Hide VISA radio buttons if VISA32 library is not found on system
If $VISA32_AVAILABLE = 0 Then GUICtrlSetState($rVISA,$GUI_HIDE)


Global Const $MCR_GRP_TOP = 0, $MCR_GRP_LEFT = 654+3, $MCR_GRP_WIDTH = 340, $MCR_ROW_HEIGHT = 24,$MACRO_WIN_WIDTH = $MCR_GRP_WIDTH
Global $MACRO_INPUT_W = 241, $MACRO_INPUT_DIFF = 36

Global $macroParVisible[$MACRO_NUMBER][2]
Global $macroString[$MACRO_NUMBER]
Global $macroStrCat[$MACRO_NUMBER][3], $macroStrPar[$MACRO_NUMBER][2]

;~ $TerminalMacros = GUICreate("Terminal macros", 663, 683, 193, 130)
$gMacros = GUICtrlCreateGroup("Macros", $MCR_GRP_LEFT, $MCR_GRP_TOP, $MCR_GRP_WIDTH, 17 + ($MCR_ROW_HEIGHT * $MACRO_NUMBER))
For $i = 0 To $MACRO_NUMBER - 1
    Local $top = $MCR_GRP_TOP+16, $rowHeight = 24
    $iMcr[$i] = GUICtrlCreateInput("", $MCR_GRP_LEFT+4, $top + ($MCR_ROW_HEIGHT *$i), $MACRO_INPUT_W, 21)
    GUICtrlSetState(-1,$GUI_HIDE)
    $iMcrP[0][$i] = GUICtrlCreateInput("", $MCR_GRP_LEFT+176, $top + ($MCR_ROW_HEIGHT *$i), 33, 21)
    GUICtrlSetState(-1,$GUI_HIDE)
    $iMcrP[1][$i] = GUICtrlCreateInput("", $MCR_GRP_LEFT+212, $top + ($MCR_ROW_HEIGHT *$i), 33, 21)
    GUICtrlSetState(-1,$GUI_HIDE)
    $bMcrSend[$i] = GUICtrlCreateButton("M" & $i+1, $MCR_GRP_LEFT+248, $top + ($MCR_ROW_HEIGHT *$i), 33, 21)
    GUICtrlSetState(-1,$GUI_HIDE)
    $iMcrRT[$i] = GUICtrlCreateInput("1000", $MCR_GRP_LEFT+284, $top + ($MCR_ROW_HEIGHT *$i), 33, 21)
    GUICtrlSetState(-1,$GUI_HIDE)
    $checkMcrRsend[$i] = GUICtrlCreateCheckbox("", $MCR_GRP_LEFT+321, $top+2 + ($MCR_ROW_HEIGHT *$i), 17, 17)
    GUICtrlSetState(-1,$GUI_HIDE)
    $macroParVisible[$i][0] = 0
    $macroParVisible[$i][1] = 0
    GUICtrlSetOnEvent($bMcrSend[$i],"macroEventSend")
    GUICtrlSetOnEvent($iMcr[$i],"macroEventSend")

    GUICtrlSetOnEvent($checkMcrRsend[$i],"macroRepeatSend")
Next

For $i = 0 To $MACRO_NUMBER - 1
    $macroStrCat[$i][0] = ""
    $macroStrCat[$i][1] = ""
    $macroStrCat[$i][2] = ""
    $macroStrPar[$i][0] = ""
    $macroStrPar[$i][1] = ""
Next


;~ $macroParVisible[1][0] = 1
;~ $macroParVisible[2][1] = 1
;~ $macroParVisible[3][0] = 1
;~ $macroParVisible[3][1] = 1

GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlSetState($gMacros,$GUI_HIDE)
;~ GUISetState(@SW_SHOW)

GUICtrlSetOnEvent($checkTX_CR,"regStoreGUI")
GUICtrlSetOnEvent($checkCRLF,"regStoreGUI")

GUISetOnEvent($GUI_EVENT_CLOSE, "SpecialEvents")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "SpecialEvents")
GUISetOnEvent($GUI_EVENT_RESTORE, "SpecialEvents")

GUICtrlSetOnEvent($rCOM,"ShowCommdialog")
GUICtrlSetOnEvent($rLAN,"ShowCommdialog")
GUICtrlSetOnEvent($rVISA,"ShowCommdialog")
GUICtrlSetOnEvent($bScanCOM,"scanCOMports")

GUICtrlSetOnEvent($cCOM,"changeCOMport")

GUICtrlSetOnEvent($bConnect,"toggleConnection")
GUICtrlSetOnEvent($bMacroWindow,"toggleMacrosView")
GUICtrlSetOnEvent($cClearRX,"clearRXbuffer")

GUICtrlSetOnEvent($bLoadMacro,"readMacroFile")
GUICtrlSetOnEvent($bSaveMacro,"writeMacroFile")


Opt("GUIOnEventMode", 1)

Global $GUI_width, $GUI_height
Global $GUIdimensions = WinGetPos("SRLabs Terminal")
If Not @error Then
    $GUI_width = $GUIdimensions[2]
    $GUI_height = $GUIdimensions[3]
Else
    ConsoleWrite (" Can't get window dimensions, setting random defaults!" & @CRLF)
    $GUI_width = 654
    $GUI_height = 666
EndIf
regLoadDirectories()
storeAccelerators()
scanCOMports()
regLoadCommType()
regLoadLAN()
If $VISA32_AVAILABLE Then regLoadVISA()
regLoadRXheadTail()
regReadMacros()
regLoadGUI()
macrosVisible($ShowMacros)

Local $temp
While 1
    If $mcrRepeat = true and BitAND(GUICtrlRead ($checkMcrRsend[$mcrRptCur]), $GUI_CHECKED) = $GUI_CHECKED then
	macroSend($mcrRptCur)
	Local $tInit = TimerInit(), $tLimit = GUICtrlRead ($iMcrRT[$mcrRptCur])
	While TimerDiff($tInit) < $tLimit
	    MainLoop() ; process input during wait
	Wend
    else
	MainLoop()
    EndIf
WEnd

#include "term_comms.au3"
#include "term_reg.au3"
#include "term_macros.au3"

Func Mainloop()
    Local $i
    ; determine which control has focus
    If WinGetHandle("") == $Terminal Then
	Local $a, $i, $match = 0
	HotKeySet("^{DEL}", "clearRXbuffer")
	$a = ControlGetHandle ($Terminal, "", ControlGetFocus($Terminal))
	$a = _WinAPI_GetDlgCtrlID($a)
	If $a = $inputTX Then
	    HotKeySet("{ENTER}", "captureENTER")
	    HotKeySet("{UP}", "captureUP")
	    HotKeySet("{DOWN}", "captureDOWN")
	Else
	    For $i = 0 To $MACRO_NUMBER - 1
		If $a = $iMcr[$i] Then
		    HotKeySet("{ENTER}", "captureENTER")
		    HotKeySet("{DOWN}")
		    HotKeySet("{UP}")
		    $match = 1
		    ExitLoop
		EndIf
	    Next
	    if $match = 0 Then
		HotKeySet("{ENTER}")
		HotKeySet("{DOWN}")
		HotKeySet("{UP}")
	    EndIf
	EndIf
    Else
	HotKeySet("^{DEL}")
	HotKeySet("{ENTER}")
	HotKeySet("{DOWN}")
	HotKeySet("{UP}")
    EndIf
    If $ConOpen = $connectCOM Or $ConOpen = $connectLAN Then
	readRXbuffer()
    EndIf
    If GUICtrlRead($inHead) <> $RXheadStr Then parseRXhead()
    If GUICtrlRead($InTail) <> $RXtailStr Then parseRXtail()
    If $ShowMacros Then
	For $i = 0 To $MACRO_NUMBER - 1
	    If $macroString[$i] <> GUICtrlRead($iMcr[$i]) Then
		parseMacroEntry($i)
	    EndIf
	Next
    EndIf
EndFunc


Func SpecialEvents()
    Switch @GUI_CtrlId
        Case $GUI_EVENT_CLOSE
            Terminate()
        Case $GUI_EVENT_MINIMIZE

        Case $GUI_EVENT_RESTORE
    EndSwitch
EndFunc   ;==>SpecialEvents


Func Terminate ()
    If $logFileSelected Then FileClose($hLog)
    closeSocket()
    TCPShutdown ( )
    Exit
EndFunc


Func storeAccelerators ()
    Local $accelListSize = $MACRO_NUMBER, $i
    If $MACRO_NUMBER > 24 Then  $accelListSize = 24
    Local $mod = "", $accelList[$accelListSize][2]
    For $i = 0 To $accelListSize - 1
	If $i > 11 Then
	    $accelList[$i][0] = "^{F" & $i-11 & "}"
	Else
	    $accelList[$i][0] = "{F" & $i+1 & "}"
	EndIf
	$accelList[$i][1] = $bMcrSend[$i]
;~ 	ConsoleWrite("Accel " & $i & ": [" & $accelList[$i][0] & ","&$accelList[$i][1] &"]" & @CRLF)
    Next
    GUISetAccelerators($accelList)
EndFunc


Func toggleConnection()
    If BitAND(GUICtrlRead ($rCOM), $GUI_CHECKED) = $GUI_CHECKED Then
;~ 	ConsoleWrite("Do COM" & @CRLF)
	toggleCOMconnection()
    Elseif BitAND(GUICtrlRead ($rLAN), $GUI_CHECKED) = $GUI_CHECKED Then
;~ 	ConsoleWrite("Do LAN" & @CRLF)
	toggleLANconnection()
    Elseif BitAND(GUICtrlRead ($rVISA), $GUI_CHECKED) = $GUI_CHECKED Then
;~ 	ConsoleWrite("Do VISA" & @CRLF)
	toggleVISAconnection()
    EndIf
EndFunc


Func clearRXbuffer()
    GUICtrlSetData($editRX, "")
EndFunc


Func changeCOMport ()
    regLoadCOMport(0)
EndFunc


Func ShowCommDialog()
    Local $type = 0
    If @GUI_CTRLID = $rVISA Then $type = 2
    If @GUI_CTRLID = $rLAN Then $type = 1
    showDialog($type)
EndFunc


Func showDialog($_type, $_init = 0)
    if $_type = $useLAN Then
	GUICtrlSetState($rCOM,$GUI_UNCHECKED)
	GUICtrlSetState($gCOMset,$GUI_HIDE)
	GUICtrlSetState($bScanCOM,$GUI_HIDE)
	GUICtrlSetState($gCOMPort,$GUI_HIDE)
	GUICtrlSetState($cCOM,$GUI_HIDE)
	GUICtrlSetState($gBaudrate,$GUI_HIDE)
	GUICtrlSetState($cBaud,$GUI_HIDE)
	GUICtrlSetState($gData,$GUI_HIDE)
	GUICtrlSetState($cData,$GUI_HIDE)
	GUICtrlSetState($gParity,$GUI_HIDE)
	GUICtrlSetState($cParity,$GUI_HIDE)
	GUICtrlSetState($gStopB,$GUI_HIDE)
	GUICtrlSetState($cStopbits,$GUI_HIDE)
	GUICtrlSetState($gHandshake,$GUI_HIDE)
	GUICtrlSetState($cHandShake,$GUI_HIDE)

	GUICtrlSetState($rLAN,$GUI_CHECKED)
	GUICtrlSetState($gLANsettings,$GUI_SHOW)
	GUICtrlSetState($gIP,$GUI_SHOW)
	GUICtrlSetState($iIP,$GUI_SHOW)
	GUICtrlSetState($gLANBport,$GUI_SHOW)
	GUICtrlSetState($iPort,$GUI_SHOW)

	GUICtrlSetState($rVISA,$GUI_UNCHECKED)
	GUICtrlSetState($gVISAset,$GUI_HIDE)
	GUICtrlSetState($gVISAaddr,$GUI_HIDE)
	GUICtrlSetState($iVISAaddr,$GUI_HIDE)
	If $ConOpen = $connectCOM Then toggleCOMconnection()
	If $ConOpen = $connectVISA Then toggleVISAconnection()
    Elseif $_type = $useCOM Then
	GUICtrlSetState($rCOM,$GUI_CHECKED)
	GUICtrlSetState($gCOMset,$GUI_SHOW)
	GUICtrlSetState($bScanCOM,$GUI_SHOW)
	GUICtrlSetState($gCOMPort,$GUI_SHOW)
	GUICtrlSetState($cCOM,$GUI_SHOW)
	GUICtrlSetState($gBaudrate,$GUI_SHOW)
	GUICtrlSetState($cBaud,$GUI_SHOW)
	GUICtrlSetState($gData,$GUI_SHOW)
	GUICtrlSetState($cData,$GUI_SHOW)
	GUICtrlSetState($gParity,$GUI_SHOW)
	GUICtrlSetState($cParity,$GUI_SHOW)
	GUICtrlSetState($gStopB,$GUI_SHOW)
	GUICtrlSetState($cStopbits,$GUI_SHOW)
	GUICtrlSetState($gHandshake,$GUI_SHOW)
	GUICtrlSetState($cHandShake,$GUI_SHOW)

	GUICtrlSetState($rLAN,$GUI_UNCHECKED)
	GUICtrlSetState($gLANsettings,$GUI_HIDE)
	GUICtrlSetState($gIP,$GUI_HIDE)
	GUICtrlSetState($iIP,$GUI_HIDE)
	GUICtrlSetState($gLANBport,$GUI_HIDE)
	GUICtrlSetState($iPort,$GUI_HIDE)

	GUICtrlSetState($rVISA,$GUI_UNCHECKED)
	GUICtrlSetState($gVISAset,$GUI_HIDE)
	GUICtrlSetState($gVISAaddr,$GUI_HIDE)
	GUICtrlSetState($iVISAaddr,$GUI_HIDE)
	If $ConOpen = $connectLAN Then toggleLANconnection()
	If $ConOpen = $connectVISA Then toggleVISAconnection()
	If $_init Then
	    scanCOMports()
	    regLoadLastCOMport()
	EndIf
    Else
	GUICtrlSetState($rCOM,$GUI_UNCHECKED)
	GUICtrlSetState($gCOMset,$GUI_HIDE)
	GUICtrlSetState($bScanCOM,$GUI_HIDE)
	GUICtrlSetState($gCOMPort,$GUI_HIDE)
	GUICtrlSetState($cCOM,$GUI_HIDE)
	GUICtrlSetState($gBaudrate,$GUI_HIDE)
	GUICtrlSetState($cBaud,$GUI_HIDE)
	GUICtrlSetState($gData,$GUI_HIDE)
	GUICtrlSetState($cData,$GUI_HIDE)
	GUICtrlSetState($gParity,$GUI_HIDE)
	GUICtrlSetState($cParity,$GUI_HIDE)
	GUICtrlSetState($gStopB,$GUI_HIDE)
	GUICtrlSetState($cStopbits,$GUI_HIDE)
	GUICtrlSetState($gHandshake,$GUI_HIDE)
	GUICtrlSetState($cHandShake,$GUI_HIDE)

	GUICtrlSetState($rLAN,$GUI_UNCHECKED)
	GUICtrlSetState($gLANsettings,$GUI_HIDE)
	GUICtrlSetState($gIP,$GUI_HIDE)
	GUICtrlSetState($iIP,$GUI_HIDE)
	GUICtrlSetState($gLANBport,$GUI_HIDE)
	GUICtrlSetState($iPort,$GUI_HIDE)

	GUICtrlSetState($rVISA,$GUI_CHECKED)
	GUICtrlSetState($gVISAset,$GUI_SHOW)
	GUICtrlSetState($gVISAaddr,$GUI_SHOW)
	GUICtrlSetState($iVISAaddr,$GUI_SHOW)
	If $ConOpen = $connectCOM Then toggleCOMconnection()
	If $ConOpen = $connectLAN Then toggleLANconnection()

    EndIf
    If Not $_init Then regStoreCommType($_type)
EndFunc


Func isIP ($_s)
    Local $a
;    ConsoleWrite ("IP: " & $_s)
    $a = StringSplit($_s,".")
    if ($a[0] <> 4) Then Return 0
;    ConsoleWrite("Still here")
    For $i = 1 To 4
	If $i = 1 Then
	    If Number($a[1]) = 0 Or Number($a[1]) > 255 Then Return 0
	Else
	    If Number($a[$i]) > 255 Then Return 0
	EndIf
    Next
    Return 1
EndFunc

#include "Term_hist.au3"