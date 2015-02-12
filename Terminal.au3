#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****

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
#include <GuiRichEdit.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <GuiComboBox.au3>
#include "COMMs.au3"

$DETACH_MACROS = 0


TCPStartup()

Global Enum $connectNone = 0, $connectCOM, $connectLAN, $connectVISA
Global Enum $useCOM = 0, $useLAN, $useVISA
Global $cMaxEditLen = 25000
Global $Terminal, $inputTX, $termMain, $editRX, $winMacro

Global $COMlist[200], $COMcount = 0, $curCOM = 0
Global $hCOM
Global $history[200], $histFirst = 0, $histLast = 0, $histCur = 0, $noHist = True
Global $inputTXscratch = "" ; command entered into the command line
Global $histLastString = ""
Global $ConOpen = $connectNone
Global $editTXcount = 0, $editRXcount = 0

Global $RXhead[2], $RXheadEcho, $RXtail ; parsed head/tail strings
Global $RXheadStr, $RXTailStr ; combined head/tail strings
Global $TXsent = 0, $TXtimeStamp[10] ; count the transmits
Global $useDelimiters = 0
Global $rxDelimStr = ""

Global $hLog, $logEnabled = 0, $logFileSelected = 0

Global $macroDirectory = @WorkingDir, $logDirectory = @WorkingDir

Global Const $REG_ROOT = "HKEY_CURRENT_USER\Software\SRLabs\Terminal"

Global $hVISAdll = 0
;~ $hVISAdll = "visa32.dll"
;~ $hVISAdll = DllOpen("visa32.dll")
#include "visa_functions.au3"
;~ #include <Visa.au3>

Global Const $MACRO_PER_BANK = 26, $MACRO_BANKS = 10, $MACRO_NUMBER = $MACRO_PER_BANK * $MACRO_BANKS
Global $curMbank = 0, $BankFirst
Global $iMcr[$MACRO_NUMBER], $bMcrSend[$MACRO_NUMBER], $iMcrRT[$MACRO_NUMBER], $checkMcrRsend[$MACRO_NUMBER]
Global $radioBank[$MACRO_BANKS]
Global $mcrRepeat = False, $mcrRptCur = 0

Global $ShowMacros = 0

Global $mainTime1 = TimerInit()
Global $mainState = 0
Global Const $MAINLOOP_REFRESH_PERIOD = 100 ; ms

; Accelerator keys
Global $accelKeyArray[1][2] = [["",""]]


Global $VISA32_AVAILABLE = 0
$hVISAdll = DllOpen("visa32.dll")

If $hVISAdll <> -1 Then
    DllClose($hVISAdll)
    $VISA32_AVAILABLE = 1
EndIf

$DUMMY_FILL = 0

Global $DefREtxtBgColor = 0x000000
Global $DefREbgColor = 0x000000
Global $DefREtxtColor = 0x808080
Global $REtxtBgColor = $DefREtxtBgColor
Global $REbgColor = $DefREbgColor
Global $REtxtColor = $DefREtxtColor
Global $REtxtAttr = "-bo-di-em-hi-im-it-li-ou-pr-re-sh-sm-st-sb-sp-un-al"


Global $escOn = 0 ; escape sequence in progress
Global $escStr = "" ; escape sequence string
Global $EscSel = 0

Global $editTXnewline = 0


Opt("GUIResizeMode", 802)
;~ HotKeySet("{UP}", "captureUP")
Func captureUP()
    If WinGetHandle("") <> $Terminal Then
        HotKeySet("{UP}")
        Send("{UP}")
        HotKeySet("{UP}", "captureUP")
        Return
    EndIf
    Local $a
    $a = ControlGetHandle($Terminal, "", ControlGetFocus($Terminal))
    $a = _WinAPI_GetDlgCtrlID($a)
    If $a <> $inputTX Then
        HotKeySet("{UP}")
        Send("{UP}")
        HotKeySet("{UP}", "captureUP")
        Return
    EndIf
;~     ConsoleWrite("UP!!!" & @CRLF)
    historyPrev()
EndFunc   ;==>captureUP



;~ HotKeySet("{DOWN}", "captureDOWN")
Func captureDOWN()
    If WinGetHandle("") <> $Terminal Then
        HotKeySet("{DOWN}")
        Send("{DOWN}")
        HotKeySet("{DOWN}", "captureDOWN")
        Return
    EndIf
    Local $a
    $a = ControlGetHandle($Terminal, "", ControlGetFocus($Terminal))
    $a = _WinAPI_GetDlgCtrlID($a)
    If $a <> $inputTX Then
        HotKeySet("{DOWN}")
        Send("{DOWN}")
        HotKeySet("{DOWN}", "captureDOWN")
        Return
    EndIf
;~     ConsoleWrite("DOWN!!!" & @CRLF)
    historyNext()
EndFunc   ;==>captureDOWN


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
    $a = ControlGetHandle($Terminal, "", ControlGetFocus($Terminal))
    $a = _WinAPI_GetDlgCtrlID($a)
    if $a = $editRX then
;~         if
;~         EndIf
    EndIf
    If $a = $inputTX Then
;~     ConsoleWrite("ENTER!!!" & @CRLF)
        sendInputData()
        Return
    EndIf
	For $i = 0 To $MACRO_PER_BANK - 1
		If $a = $iMcr[$i] Then
			macroSend($BankFirst + $i)
			Return
		EndIf
	Next
    HotKeySet("{ENTER}")
    Send("{ENTER}")
    HotKeySet("{ENTER}", "captureENTER")
    Return

EndFunc   ;==>captureENTER

local $__initTime = TimerInit()

#region ### START Koda GUI section ### Form=D:\scripts\Terminal\Terminal2.kxf
;$Terminal = GUICreate("SRLabs Terminal - " & FileGetVersion(@ScriptFullPath), 654, 666, -1, -1, -1, -1)
$Terminal = GUICreate("SRLabs Terminal - " & FileGetVersion(@ScriptFullPath), 586, 666, -1, -1, -1, -1)

$Connection = GUICtrlCreateGroup("", 0, 0, 123, 60) ; Connection group
$rCOM = GUICtrlCreateRadio("COM", 4, 10, 41, 15)
$rLAN = GUICtrlCreateRadio("LAN", 4, 26, 41, 15)
$rVISA = GUICtrlCreateRadio("VXI", 4, 42, 41, 15)
$bConnect = GUICtrlCreateButton("Connect", 52, 12, 65, 17)
$bScanCOM = GUICtrlCreateButton("Scan", 52, 37, 65, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gCOMset = GUICtrlCreateGroup("", 126, 0, 439, 60) ;Com group
$gCOMPort = GUICtrlCreateGroup("COM port", 132, 8, 69, 48)
$cCOM = GUICtrlCreateCombo("", 136, 28, 61, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gBaudrate = GUICtrlCreateGroup("Baudrate", 204, 8, 77, 48)
$cBaud = GUICtrlCreateCombo("", 208, 28, 69, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "600|1200|2400|4800|9600|14400|19200|28800|38400|56000|57600|115200|128000|256000", "115200")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gData = GUICtrlCreateGroup("Data bits", 284, 8, 57, 48)
$cData = GUICtrlCreateCombo("", 288, 28, 49, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "5|6|7|8", "8")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gParity = GUICtrlCreateGroup("Parity", 344, 8, 65, 48)
$cParity = GUICtrlCreateCombo("", 348, 28, 57, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "none|odd|even|mark|space", "none")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gStopB = GUICtrlCreateGroup("Stop bits", 412, 8, 57, 48)
$cStopbits = GUICtrlCreateCombo("", 416, 28, 49, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "1|1.5|2", "1")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gHandshake = GUICtrlCreateGroup("Handshake", 472, 8, 89, 48)
$cHandShake = GUICtrlCreateCombo("", 476, 28, 81, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "none|RTS/CTS|XON/XOFF|RTS+XON", "none")
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlSetState(-1, $GUI_DISABLE)

$gLANsettings = GUICtrlCreateGroup("", 128, 0, 309, 60)
$gIP = GUICtrlCreateGroup("IP address", 132, 8, 201, 48)
$iIP = GUICtrlCreateInput("", 140, 28, 185, 21)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$gLANBport = GUICtrlCreateGroup("Port", 336, 8, 97, 48)
$iPort = GUICtrlCreateInput("", 344, 28, 81, 21)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlSetState(-1, $GUI_DISABLE)

$gVISAset = GUICtrlCreateGroup("", 128, 0, 349, 60)
$gVISAaddr = GUICtrlCreateGroup("VXI address", 132, 8, 341, 48)
$iVISAaddr = GUICtrlCreateInput("", 140, 28, 329, 21)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlSetState(-1, $GUI_DISABLE)

$gRXfilter = GUICtrlCreateGroup("Receive head / tail filter", 0, 60, 138, 69)

$checkEnableRXfilter = GUICtrlCreateLabel("On", 4, 75, 16, 21, BitOR($GUI_SS_DEFAULT_CHECKBOX,$BS_MULTILINE))
$checkEnableRXfilter = GUICtrlCreateCheckbox("", 4, 92, 16, 16, BitOR($GUI_SS_DEFAULT_CHECKBOX,$BS_MULTILINE))
$InHead = GUICtrlCreateInput("", 22, 76, 112, 21)
$InTail = GUICtrlCreateInput("", 22, 102, 112, 21)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group1 = GUICtrlCreateGroup("", 141, 60, 272, 69)
$cClearRX = GUICtrlCreateButton("Clear RX buffer", 145, 80, 81, 17)
$chkShowBlanks = GUICtrlCreateCheckbox("Show blank chars", 145, 104, 109, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gLog = GUICtrlCreateGroup("Logging", 416, 60, 73, 69)
$bSelectLog = GUICtrlCreateButton("Select file", 420, 80, 65, 17)
$bStartLog = GUICtrlCreateButton("Start", 420, 104, 29, 17)
$bStopLog = GUICtrlCreateButton("Stop", 456, 104, 29, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$gMacros = GUICtrlCreateGroup("Macros", 492, 60, 93, 69)
$bMacroWindow = GUICtrlCreateButton("Macro Window ->", 496, 80, 85, 17)
$bLoadMacro = GUICtrlCreateButton("Load", 496, 104, 39, 17)
$bSaveMacro = GUICtrlCreateButton("Save", 542, 104, 39, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$editTX = GUICtrlCreateEdit("", 0, 612, 585, 53, BitOR($ES_AUTOVSCROLL,$ES_AUTOHSCROLL,$ES_READONLY,$WS_VSCROLL))
GUICtrlSetFont(-1, 9, 400, 0, "Courier New")
GUICtrlSetColor(-1, 0x000000)
;~ $editRX = GUICtrlCreateEdit("", 0, 136, 653, 389, BitOR($ES_AUTOVSCROLL,$ES_AUTOHSCROLL,$ES_READONLY,$WS_VSCROLL))
$editRX = _GUICtrlRichEdit_Create($Terminal, "", 0, 130, 585, 459, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $WS_VSCROLL, $ES_MULTILINE)) ;
_GUICtrlRichEdit_SetBkColor($editRX, $DefREbgColor) ; background black
_GUICtrlRichEdit_SetCharBkColor($editRX, $DefREtxtBgColor) ; background black
_GUICtrlRichEdit_SetCharColor($editRX, $DefREtxtColor)
_GUICtrlRichEdit_SetFont($editRX, 9, "Courier New")
GUICtrlSetFont(-1, 9, 400, 0, "Courier New")
$inputTX = GUICtrlCreateInput("", 0, 590, 504, 21)
$checkTX_CR = GUICtrlCreateCheckbox("+CR", 507, 592, 41, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$checkCRLF = GUICtrlCreateCheckbox("+LF", 549, 592, 41, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
#endregion ### END Koda GUI section ###

; Hide VISA radio buttons if VISA32 library is not found on system
If $VISA32_AVAILABLE = 0 Then GUICtrlSetState($rVISA, $GUI_HIDE)


Global Const $MCR_GRP_TOP = 0, $MCR_GRP_LEFT = 585 + 2, $MCR_GRP_WIDTH = 340, $MCR_ROW_HEIGHT = 24, $MACRO_WIN_WIDTH = $MCR_GRP_WIDTH
Global $MACRO_INPUT_W = 241, $MACRO_INPUT_DIFF = 36

Global $macroString[$MACRO_NUMBER], $macroRptTime[$MACRO_NUMBER], $macroRpt[$MACRO_NUMBER]

;~ $TerminalMacros = GUICreate("Terminal macros", 663, 683, 193, 130)
$gMacros = GUICtrlCreateGroup("Macros", $MCR_GRP_LEFT, $MCR_GRP_TOP, $MCR_GRP_WIDTH, 17 + ($MCR_ROW_HEIGHT * ($MACRO_PER_BANK + 1)))

ConsoleWrite (stringformat("...init banks\n"))

For $i = 0 to $MACRO_BANKS -1
    Local $top = $MCR_GRP_TOP + 16
    Local $left = $MCR_GRP_LEFT + 4 , $columnWidth = 32
;~     $columnWidth = Floor(($MCR_GRP_WIDTH - 8 - 31)/ ($MACRO_BANKS - 1))
    if $i < 9 Then
        $radioBank[$i] = GUICtrlCreateRadio("&"&$i+1, $left+ $columnWidth * $i, $top, 31, 21)
    elseif $i == 9 Then
        $radioBank[$i] = GUICtrlCreateRadio("1&0", $left+ $columnWidth * $i, $top, 31, 21)
    Else
        $radioBank[$i] = GUICtrlCreateRadio($i+1, $left+ $columnWidth * $i, $top, 31, 21)
    endif
    GUICtrlSetState(-1, $GUI_HIDE)
    GUICtrlSetOnEvent($radioBank[$i], "changeMacroBank")
    if $i <= 9 then ; first 10 banks get accelerator keys from 1 to 0
        local $j = $i + 1
        if $j = 10 then $j = 0
;~         addAccelerator("!" & $j, $radioBank[$i])
    EndIf
Next
    GUICtrlSetState($radioBank[0], $GUI_CHECKED)

ConsoleWrite (stringformat("...init macros\n"))
For $i = 0 To $MACRO_PER_BANK - 1
	Local $rowHeight = 24
	Local $top = $MCR_GRP_TOP + 16 + 1*$rowHeight
	local $yoff = Mod($i, $MACRO_PER_BANK)
	$iMcr[$i] = GUICtrlCreateInput("", $MCR_GRP_LEFT + 4, $top + ($MCR_ROW_HEIGHT * $yoff), $MACRO_INPUT_W, 21)
	GUICtrlSetState(-1, $GUI_HIDE)
	$bMcrSend[$i] = GUICtrlCreateButton("M" & $i + 1, $MCR_GRP_LEFT + 248, $top + ($MCR_ROW_HEIGHT * $yoff), 33, 21)
	GUICtrlSetState(-1, $GUI_HIDE)
	$iMcrRT[$i] = GUICtrlCreateInput("1000", $MCR_GRP_LEFT + 284, $top + ($MCR_ROW_HEIGHT * $yoff), 33, 21)
	GUICtrlSetState(-1, $GUI_HIDE)
	$checkMcrRsend[$i] = GUICtrlCreateCheckbox("", $MCR_GRP_LEFT + 321, $top + 2 + ($MCR_ROW_HEIGHT * $yoff), 17, 17)
	GUICtrlSetState(-1, $GUI_HIDE)
	GUICtrlSetOnEvent($bMcrSend[$i], "macroEventSend")
	GUICtrlSetOnEvent($iMcr[$i], "macroEventSend")

	GUICtrlSetOnEvent($checkMcrRsend[$i], "macroRepeatSend")
	if $i < 12 then
		local $j = $i + 1
		addAccelerator("{F"&$j&"}",$bMcrSend[$i])
                GUICtrlSetTip($bMcrSend[$i], "Shortcut: F"&$j)
	elseif $i < 24 Then
		local $j = $i - 11
		addAccelerator("+{F"&$j&"}",$bMcrSend[$i])
                GUICtrlSetTip($bMcrSend[$i], "Shortcut: Shift+F"&$j)
	EndIf

Next
For $i = 0 To $MACRO_NUMBER - 1
	$macroRptTime[$i] = 1000
	$macroRpt[$i] = 0
next
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlSetState($gMacros, $GUI_HIDE)
; end of macro GUI element declaration


ConsoleWrite (stringformat("...init events\n"))

GUICtrlSetOnEvent($checkTX_CR, "regStoreGUI")
GUICtrlSetOnEvent($checkCRLF, "regStoreGUI")

GUISetOnEvent($GUI_EVENT_CLOSE, "SpecialEvents")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "SpecialEvents")
GUISetOnEvent($GUI_EVENT_RESTORE, "SpecialEvents")

GUICtrlSetOnEvent($rCOM, "ShowCommdialog")
GUICtrlSetOnEvent($rLAN, "ShowCommdialog")
GUICtrlSetOnEvent($rVISA, "ShowCommdialog")
GUICtrlSetOnEvent($bScanCOM, "scanCOMports")

GUICtrlSetOnEvent($cCOM, "changeCOMport")

GUICtrlSetOnEvent($bConnect, "toggleConnection")
GUICtrlSetOnEvent($bMacroWindow, "toggleMacrosView")
GUICtrlSetOnEvent($cClearRX, "clearRXbuffer")

GUICtrlSetOnEvent($bLoadMacro, "readMacroFile")
GUICtrlSetOnEvent($bSaveMacro, "writeMacroFile")
GUICtrlSetOnEvent($editRX, "switchToInputTX")


GUICtrlSetOnEvent($checkEnableRXfilter, "changeDelimUsage")

GUISetState(@SW_SHOW)

Opt("GUIOnEventMode", 1)
ConsoleWrite (stringformat("...init variables\n"))

Global $GUI_width, $GUI_height
Global $GUIdimensions = WinGetPos("SRLabs Terminal")
If Not @error Then
    $GUI_width = $GUIdimensions[2]
    $GUI_height = $GUIdimensions[3]
Else
    ConsoleWrite(" Can't get window dimensions, setting random defaults!" & @CRLF)
    $GUI_width = 565
    $GUI_height = 666
EndIf
CalcFirstMacroInBank()
regLoadDirectories()
addAccelerator("","",1) ; Load accelerators
scanCOMports()
regLoadCommType()
regLoadLAN()
If $VISA32_AVAILABLE Then regLoadVISA()
regLoadRXheadTail()
regReadMacros()
regLoadGUI()
macrosVisible($ShowMacros)
;~ makeMacroWindow()

ConsoleWrite (stringformat("Running in %f ms.\n", TimerDiff ($__initTime)))

Local $temp
While 1
	If $mcrRepeat = True And $macroRpt[$mcrRptCur] Then
		macroSend($mcrRptCur)
		Local $tInit = TimerInit(), $tLimit = $macroRptTime[$mcrRptCur]
		While TimerDiff($tInit) < $tLimit
			Mainloop() ; process input during wait
			Sleep(10)
		WEnd
	Else
		Mainloop()
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
        $a = ControlGetHandle($Terminal, "", ControlGetFocus($Terminal))
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
            If $match = 0 Then
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
    If $useDelimiters Then
        If GUICtrlRead($InHead) <> $RXheadStr Then parseRXhead()
        If GUICtrlRead($InTail) <> $RXTailStr Then parseRXtail()
    EndIf
    If $ShowMacros then
        If TimerDiff ($mainTime1) > $MAINLOOP_REFRESH_PERIOD Then
            $mainTime1 = TimerInit()
            if $mainState = 0 then
                For $i = 0 To $MACRO_PER_BANK - 1
                    If $macroString[$BankFirst+$i] <> GUICtrlRead($iMcr[$i]) Then
                        parseMacroEntry($BankFirst+$i)
                    EndIf
                Next
            else
                For $i = 0 To $MACRO_PER_BANK - 1
                    If $macroRptTime[$BankFirst+$i] <> GUICtrlRead($iMcrRT[$i]) Then
                        $macroRptTime[$BankFirst+$i] = GUICtrlRead($iMcrRT[$i])
                    EndIf
                Next
            EndIf
            $mainState = 1 - $mainState
;~         Else
;~             Sleep(10)
        EndIf
    EndIf
EndFunc   ;==>Mainloop





Func makeMacroWindow()
;~     $winMacro = GUICreate("Macros", $MCR_GRP_WIDTH, 650, -1, -1,-1, -1, $Terminal) ; style 0x84ca0000
    $winMacro = GUICreate("Macros", $MCR_GRP_WIDTH, 650, -1, -1,BitOR($WS_POPUP, $WS_CLIPCHILDREN,$WS_CAPTION), -1, $Terminal)
    local $a = GUIGetStyle($winMacro)
    ConsoleWrite (StringFormat("style:%x, ext style:%x\n",$a[0],$a[1]))
    local $offleft = 0, $offtop = 4
    local $i

For $i = 0 to $MACRO_BANKS -1
    Local $top = $offtop
    Local $left = $offleft + 4 , $columnWidth = 32
;~     $columnWidth = Floor(($MCR_GRP_WIDTH - 8 - 31)/ ($MACRO_BANKS - 1))
    $radioBank[$i] = GUICtrlCreateRadio("&"&$i+1, $left+ $columnWidth * $i, $top, 31, 21)
;~     GUICtrlSetState(-1, $GUI_HIDE)
    GUICtrlSetOnEvent($radioBank[$i], "changeMacroBank")
;~     if $i <= 9 then ; first 10 banks get accelerator keys from 1 to 0
;~         local $j = $i + 1
;~         if $j = 10 then $j = 0
;~         addAccelerator("!" & $j, $radioBank[$i])
;~     EndIf
Next
    GUICtrlSetState($radioBank[0], $GUI_CHECKED)

    For $i = 0 To $MACRO_PER_BANK - 1
        Local $rowHeight = 24
        Local $top = $offtop + 1*$rowHeight
        local $yoff = Mod($i, $MACRO_PER_BANK)
        $iMcr[$i] = GUICtrlCreateInput("", $offleft + 4, $top + ($MCR_ROW_HEIGHT * $yoff), $MACRO_INPUT_W, 21)
;~         GUICtrlSetState(-1, $GUI_HIDE)
        $bMcrSend[$i] = GUICtrlCreateButton("M" & $i + 1, $offleft + 248, $top + ($MCR_ROW_HEIGHT * $yoff), 33, 21)
;~         GUICtrlSetState(-1, $GUI_HIDE)
        $iMcrRT[$i] = GUICtrlCreateInput("1000", $offleft + 284, $top + ($MCR_ROW_HEIGHT * $yoff), 33, 21)
;~         GUICtrlSetState(-1, $GUI_HIDE)
        $checkMcrRsend[$i] = GUICtrlCreateCheckbox("", $offleft + 321, $top + 2 + ($MCR_ROW_HEIGHT * $yoff), 17, 17)
;~         GUICtrlSetState(-1, $GUI_HIDE)
        GUICtrlSetOnEvent($bMcrSend[$i], "macroEventSend")
        GUICtrlSetOnEvent($iMcr[$i], "macroEventSend")

        GUICtrlSetOnEvent($checkMcrRsend[$i], "macroRepeatSend")
        if $i < 12 then
            local $j = $i + 1
;~             addAccelerator("{F"&$j&"}",$iMcr[$i])
        elseif $i < 24 Then
            local $j = $i - 11
;~             addAccelerator("+{F"&$j&"}",$iMcr[$i])
        EndIf

    Next
    For $i = 0 To $MACRO_NUMBER - 1
        $macroRptTime[$i] = 1000
        $macroRpt[$i] = 0
    next
	GUICtrlSetState($gMacros, $GUI_HIDE)
	GUISetState(@SW_SHOW)
	GUISetOnEvent($GUI_EVENT_CLOSE, "closeMacroWindow")
; end of macro GUI element declaration

EndFunc


Func closeMacroWindow()
;~     ConsoleWrite(StringFormat("Tuki.\n"))
    GUIDelete($winMacro)
EndFunc


Func CalcFirstMacroInBank()
    $BankFirst = $curMbank * $MACRO_PER_BANK
EndFunc

Func switchToInputTX()
    ConsoleWrite(StringFormat("tu\n"))
    Local $curCtrl = ControlGetFocus("")
    ControlFocus("", "", $inputTX)
EndFunc   ;==>switchToInputTX

Func changeDelimUsage()
    If BitAND(GUICtrlRead($checkEnableRXfilter), $GUI_CHECKED) = $GUI_CHECKED Then
        $useDelimiters = 1
    Else
        $useDelimiters = 0
    EndIf
EndFunc   ;==>changeDelimUsage



Func addAccelerator ($_key, $_id, $_refresh = 0)
    If $_refresh = 1 Then
;~         ConsoleWrite(stringformat("Ubound accelkeyarray = %d", UBound($accelKeyArray)))
;~             _ArrayDisplay($accelKeyArray)
        GUISetAccelerators($accelKeyArray,$Terminal)
        Return
    EndIf
    if $accelKeyArray[0][0] == "" then
        $accelKeyArray[0][0] = $_key
        $accelKeyArray[0][1] = $_id
    Else
        Redim $accelKeyArray[Ubound($accelKeyArray) + 1][2]
        $accelKeyArray[Ubound($accelKeyArray) - 1][0] = $_key
        $accelKeyArray[Ubound($accelKeyArray) - 1][1] = $_id
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


Func Terminate()
    If $logFileSelected Then FileClose($hLog)
    closeSocket()
    TCPShutdown()
    _GUICtrlRichEdit_Destroy($editRX)
    Exit
EndFunc   ;==>Terminate


Func toggleConnection()
    If BitAND(GUICtrlRead($rCOM), $GUI_CHECKED) = $GUI_CHECKED Then
        toggleCOMconnection()
    ElseIf BitAND(GUICtrlRead($rLAN), $GUI_CHECKED) = $GUI_CHECKED Then
        toggleLANconnection()
    ElseIf BitAND(GUICtrlRead($rVISA), $GUI_CHECKED) = $GUI_CHECKED Then
        toggleVISAconnection()
    EndIf
EndFunc   ;==>toggleConnection


Func clearRXbuffer()
    Local $curCtrl = ControlGetFocus("")
    _GUICtrlRichEdit_SetText($editRX, "")
    $EscSel = 0
    ControlFocus("", "", $curCtrl)
EndFunc   ;==>clearRXbuffer


Func changeCOMport()
    regLoadCOMport(0)
EndFunc   ;==>changeCOMport


Func ShowCommDialog()
    Local $type = 0
    If @GUI_CtrlId = $rVISA Then $type = 2
    If @GUI_CtrlId = $rLAN Then $type = 1
    showDialog($type)
EndFunc   ;==>ShowCommDialog


Func showDialog($_type, $_init = 0)
    If $_type = $useLAN Then
        GUICtrlSetState($rCOM, $GUI_UNCHECKED)
        GUICtrlSetState($gCOMset, $GUI_HIDE)
        GUICtrlSetState($bScanCOM, $GUI_HIDE)
        GUICtrlSetState($gCOMPort, $GUI_HIDE)
        GUICtrlSetState($cCOM, $GUI_HIDE)
        GUICtrlSetState($gBaudrate, $GUI_HIDE)
        GUICtrlSetState($cBaud, $GUI_HIDE)
        GUICtrlSetState($gData, $GUI_HIDE)
        GUICtrlSetState($cData, $GUI_HIDE)
        GUICtrlSetState($gParity, $GUI_HIDE)
        GUICtrlSetState($cParity, $GUI_HIDE)
        GUICtrlSetState($gStopB, $GUI_HIDE)
        GUICtrlSetState($cStopbits, $GUI_HIDE)
        GUICtrlSetState($gHandshake, $GUI_HIDE)
        GUICtrlSetState($cHandShake, $GUI_HIDE)

        GUICtrlSetState($rLAN, $GUI_CHECKED)
        GUICtrlSetState($gLANsettings, $GUI_SHOW)
        GUICtrlSetState($gIP, $GUI_SHOW)
        GUICtrlSetState($iIP, $GUI_SHOW)
        GUICtrlSetState($gLANBport, $GUI_SHOW)
        GUICtrlSetState($iPort, $GUI_SHOW)

        GUICtrlSetState($rVISA, $GUI_UNCHECKED)
        GUICtrlSetState($gVISAset, $GUI_HIDE)
        GUICtrlSetState($gVISAaddr, $GUI_HIDE)
        GUICtrlSetState($iVISAaddr, $GUI_HIDE)
        If $ConOpen = $connectCOM Then toggleCOMconnection()
        If $ConOpen = $connectVISA Then toggleVISAconnection()
    ElseIf $_type = $useCOM Then
        GUICtrlSetState($rCOM, $GUI_CHECKED)
        GUICtrlSetState($gCOMset, $GUI_SHOW)
        GUICtrlSetState($bScanCOM, $GUI_SHOW)
        GUICtrlSetState($gCOMPort, $GUI_SHOW)
        GUICtrlSetState($cCOM, $GUI_SHOW)
        GUICtrlSetState($gBaudrate, $GUI_SHOW)
        GUICtrlSetState($cBaud, $GUI_SHOW)
        GUICtrlSetState($gData, $GUI_SHOW)
        GUICtrlSetState($cData, $GUI_SHOW)
        GUICtrlSetState($gParity, $GUI_SHOW)
        GUICtrlSetState($cParity, $GUI_SHOW)
        GUICtrlSetState($gStopB, $GUI_SHOW)
        GUICtrlSetState($cStopbits, $GUI_SHOW)
        GUICtrlSetState($gHandshake, $GUI_SHOW)
        GUICtrlSetState($cHandShake, $GUI_SHOW)

        GUICtrlSetState($rLAN, $GUI_UNCHECKED)
        GUICtrlSetState($gLANsettings, $GUI_HIDE)
        GUICtrlSetState($gIP, $GUI_HIDE)
        GUICtrlSetState($iIP, $GUI_HIDE)
        GUICtrlSetState($gLANBport, $GUI_HIDE)
        GUICtrlSetState($iPort, $GUI_HIDE)

        GUICtrlSetState($rVISA, $GUI_UNCHECKED)
        GUICtrlSetState($gVISAset, $GUI_HIDE)
        GUICtrlSetState($gVISAaddr, $GUI_HIDE)
        GUICtrlSetState($iVISAaddr, $GUI_HIDE)
        If $ConOpen = $connectLAN Then toggleLANconnection()
        If $ConOpen = $connectVISA Then toggleVISAconnection()
        If $_init Then
            scanCOMports()
            regLoadLastCOMport()
        EndIf
    Else
        GUICtrlSetState($rCOM, $GUI_UNCHECKED)
        GUICtrlSetState($gCOMset, $GUI_HIDE)
        GUICtrlSetState($bScanCOM, $GUI_HIDE)
        GUICtrlSetState($gCOMPort, $GUI_HIDE)
        GUICtrlSetState($cCOM, $GUI_HIDE)
        GUICtrlSetState($gBaudrate, $GUI_HIDE)
        GUICtrlSetState($cBaud, $GUI_HIDE)
        GUICtrlSetState($gData, $GUI_HIDE)
        GUICtrlSetState($cData, $GUI_HIDE)
        GUICtrlSetState($gParity, $GUI_HIDE)
        GUICtrlSetState($cParity, $GUI_HIDE)
        GUICtrlSetState($gStopB, $GUI_HIDE)
        GUICtrlSetState($cStopbits, $GUI_HIDE)
        GUICtrlSetState($gHandshake, $GUI_HIDE)
        GUICtrlSetState($cHandShake, $GUI_HIDE)

        GUICtrlSetState($rLAN, $GUI_UNCHECKED)
        GUICtrlSetState($gLANsettings, $GUI_HIDE)
        GUICtrlSetState($gIP, $GUI_HIDE)
        GUICtrlSetState($iIP, $GUI_HIDE)
        GUICtrlSetState($gLANBport, $GUI_HIDE)
        GUICtrlSetState($iPort, $GUI_HIDE)

        GUICtrlSetState($rVISA, $GUI_CHECKED)
        GUICtrlSetState($gVISAset, $GUI_SHOW)
        GUICtrlSetState($gVISAaddr, $GUI_SHOW)
        GUICtrlSetState($iVISAaddr, $GUI_SHOW)
        If $ConOpen = $connectCOM Then toggleCOMconnection()
        If $ConOpen = $connectLAN Then toggleLANconnection()
    EndIf
    If Not $_init Then regStoreCommType($_type)
EndFunc   ;==>showDialog


Func isIP($_s)
    Local $a
    ;    ConsoleWrite ("IP: " & $_s)
    $a = StringSplit($_s, ".")
    If ($a[0] <> 4) Then Return 0
    ;    ConsoleWrite("Still here")
    For $i = 1 To 4
        If $i = 1 Then
            If Number($a[1]) = 0 Or Number($a[1]) > 255 Then Return 0
        Else
            If Number($a[$i]) > 255 Then Return 0
        EndIf
    Next
    Return 1
EndFunc   ;==>isIP

#include "Term_hist.au3"