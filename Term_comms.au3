
; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.
#include <GuiRichEdit.au3>
#include <GuiEdit.au3>
#include <Array.au3>


Func sendInputData()
    Local $tx = GUICtrlRead ($InputTX)
    Local $ret
    checkHistory($tx)
;    $ret = sendData($tx)
    $ret = parseString($tx)
    GUICtrlSetData ($InputTX, "")
    $histLastString = $tx
    $noHist = true
    Return $ret
EndFunc



; Parse following meta strings:
; - %DLY1234 -> Delay in milliseconds
; - #xxx -> ASCII character in decimal format
; - $xx  -> ASCII character in hexadecimal format
; - %M12  -> insert macro


Func parseString($_str, $_p1 = "", $_p2 = "", $_reentrNo = 0, $DEBUG = 1)
    Const $MAX_CMDS = 100
    Const $MAX_REENTRY = 10
    Local $cmd[$MAX_CMDS], $meta[$MAX_CMDS], $cmdNo = 0, $i
    For $i = 0 to $MAX_CMDS - 1
	$cmd[$i] = ""
	$meta[$i] = ""
    Next
    ; parse input string
    Local $quote = "", $token = "", $char = "", $str = $_str, $reentry = 0
    While true
	If Not $reentry Then
	    If StringLen($str) = 0 Then ExitLoop
	    $char = StringLeft($str, 1) ; store first remaining character of the string
	    $str = StringTrimLeft ($str, 1)
	Else
	    $reentry = 0
	EndIf
	; check for quotation marks
	if $quote <> "" Then
	    if $char = $quote Then
		$quote = ""
	    Else
		$cmd[$cmdNo] &= $char
	    EndIf
	; check for tokens
	Elseif $token <> "" then
	    If $char = $token Then ; repeated token, ignore and put single characater to output string
		$token = ""
		$cmd[$cmdNo] &= $char
		ContinueLoop
	    Elseif $char = "$" or $char = "#" or $char = "%" then
		; format error handling, interpret previous token as character and set new token
		$cmd[$cmdNo] &= $token
		$token = $char
		ContinueLoop
	    ElseIf $token = "$" then
		; check for double hexadecimal characters
		Local $a = $char & StringLeft ($str, 1)
		if StringLen($a) = 2 and StringIsXDigit($a) Then
		    $cmd[$cmdNo] &= Chr(Dec($a))
		    $str = StringTrimLeft($str, 1)
		    $token = ""
		Else
		    $cmd[$cmdNo] &= $token
		    $token = ""
		    $reentry = 1
		EndIf
	    ElseIf $token = "#" then
		; check for three decimal characters
		Local $a = $char & StringLeft ($str, 2)
		if StringLen($a) = 3 and StringIsDigit($a) Then
		    $cmd[$cmdNo] &= Chr($a)
		    $str = StringTrimLeft($str, 2)
		    $token = ""
		Else
		    $cmd[$cmdNo] &= $token
		    $token = ""
		    $reentry = 1
		EndIf
	    ElseIf $token = "%" then
		If $char = "M" Then
		    Local $a = StringLeft($str, 2)
		    if StringLen($a) = 2 and StringIsDigit($a) Then
			If $cmd[$cmdNo] <> "" Then $cmdNo += 1
			$meta[$cmdNo] = "macro"
			$cmd[$cmdNo] = Int($a)
			$cmdNo += 1
			$token = ""
			$str = StringTrimLeft($str, 2)
		    Else
			$cmd[$cmdNo] &= $token
			$token = ""
			$reentry = 1
		    EndIf
		Elseif $char = "D" Then
		    Local $a = StringLeft($str, 2)
		    Local $b = StringMid($str, 3, 4)
		    if $a = "LY" and StringLen($b) = 4 and StringIsDigit($b) Then
			If $cmd[$cmdNo] <> "" Then $cmdNo += 1
			$meta[$cmdNo] = "delay"
			$cmd[$cmdNo] = Int($b)
			$cmdNo += 1
			$token = ""
			$str = StringTrimLeft($str, 6)
		    Else
			$cmd[$cmdNo] &= $token
			$token = ""
			$reentry = 1
		    EndIf
		Elseif $char = "P" Then
		    Local $a = StringLeft($str, 1)
		    $token = ""
		    if $a = "1" Then
			$str = $_p1 & StringTrimLeft($str, 1) ; add parameter string to main string
		    Elseif $a = "2" Then
			$str = $_p2 & StringTrimLeft($str, 1) ; add parameter string to main string
		    Else
			$cmd[$cmdNo] &= $token
			$reentry = 1
		    EndIf
		Elseif $char = "C" or $char = "L" then
		    Local $a = $char & StringLeft($str, 1)
		    Local $b = $char & StringLeft($str, 3)
		    $token = ""
		    if $b = "CRLF" Then
			$str = StringTrimLeft($str, 3)
			$cmd[$cmdNo] &= Chr(13) & Chr(10)
		    elseif $a = "CR" Then
			$str = StringTrimLeft($str, 1)
			$cmd[$cmdNo] &= Chr(13)
		    elseif $a = "LF" Then
			$str = StringTrimLeft($str, 1)
			$cmd[$cmdNo] &= Chr(10)
		    Else
			$cmd[$cmdNo] &= $token
			$reentry = 1
		    EndIf
		elseif $char = "F" Then
		    ; check if this should be file entry
		    Local $a = $char & StringLeft ($str, 3)
		    if $a = "FILE" Then
			If $DEBUG Then ConsoleWrite(Stringformat("Here.\r\n"))
			Local $trimLen = 3
			local $b = StringTrimLeft ($str, 3)
			$trimLen = $trimLen + StringLen($b)
			$b = StringstripWS($b, 1)
			$trimLen = $trimLen - StringLen($b)
			if StringLeft($b, 1) = "=" then
			    $b = StringTrimLeft($b,1)
			    $trimLen = $trimLen + 1
			EndIf
			local $quote  = ""
			if StringInStr($b, "'") then
			    $quote = "'"
			elseif StringInStr($b, "'") then
			    $quote = '"'
			EndIf
			if $quote <> "" Then
			    $trimLen = $trimLen + StringInStr($b, $quote)
			    Local $c = StringSplit ($b, $quote)
			    $b = $c[2]
			    $trimLen = $trimLen + StringLen($b) + 1
			EndIf
			If $DEBUG Then ConsoleWrite(Stringformat("Using filename: %s\r\n", $b))
			ParseTXfile($b, $cmd, $cmdNo)
			$str = StringTrimLeft ($str, $trimLen)
			If $DEBUG Then ConsoleWrite(Stringformat("Remaining string: '%s'\r\n", $str))
			ExitLoop
		    Endif
		EndIf
	    EndIf
	Elseif $char = "$" or $char = "#" or $char = "%" then
	    $token = $char
	    ContinueLoop
	Elseif $char = "'" or $char = '"' then
	    $quote = $char
	    ContinueLoop
	Else
	    $cmd[$cmdNo] &= $char ; add character to command string
	EndIf
    WEnd
    if  $cmd[$cmdNo] <> "" Then $cmdNo += 1
    If $DEBUG Then
	Local $dbgoffset = "", $z
	For $z = 1 to $_reentrNo
	    $dbgoffset &= "  "
	Next
	ConsoleWrite(StringFormat ("%sSending command: '%s', ", $dbgoffset, $_str))
	ConsoleWrite(StringFormat ("param1: '%s', param2: '%s'\r\n", $_p1, $_p2))
	ConsoleWrite(StringFormat ("%s>Parse table: \r\n", $dbgoffset))
	For $i = 0 To $cmdNo - 1
	    local $c = $cmd[$i]
	    if $i = $cmdNo - 1 and $_reentrNo = 0 Then
		If GUICtrlRead ($checkTX_CR) = $GUI_CHECKED Then
		    $c &= @CR
		    If GUICtrlRead ($checkCRLF) = $GUI_CHECKED Then
			$c &= @LF
		    EndIf
		EndIf
	    EndIf
	    if $meta[$i] <> "" Then
		ConsoleWrite(StringFormat ("%s  %d -> meta: '%s' = '%s'\r\n", $dbgoffset, $i, $meta[$i], $cmd[$i]))
	    Else
		ConsoleWrite(StringFormat ("%s  %d -> command: '%s'\r\n", $dbgoffset, $i, $c))
	    EndIf
	Next
    EndIf
    ; send out
    For $i = 0 To $cmdNo - 1
	if $meta[$i] = "delay" Then
	    Local $tInit = TimerInit()
	    While TimerDiff($tInit) < $cmd[$i]
		MainLoop() ; process input during wait
	    Wend
	ElseIf $meta[$i] = "macro" Then
	    If $_reentrNo < $MAX_REENTRY Then
		; get values for macro and parse the macro
		Local $m = $macroString[$cmd[$i] - 1]
		Local $mp1 = $macroStrPar[$cmd[$i] - 1][0]
		Local $mp2 = $macroStrPar[$cmd[$i] - 1][1]
		parseString($m, $mp1, $mp2, $_reentrNo + 1)
	    EndIf
	Else
	    local $ret = sendData ($cmd[$i])
	EndIf
    Next
    If $_reentrNo = 0 And $cmdNo >= 1 then
	Local $c = ""
	If GUICtrlRead ($checkTX_CR) = $GUI_CHECKED Then
	    $c &= @CR
	    If GUICtrlRead ($checkCRLF) = $GUI_CHECKED Then $c &= @LF
	EndIf
	local $ret = sendData ($c)
    EndIf
EndFunc


Func ParseTXfile ($_fName, ByRef $cmd, ByRef $cmdNo)
    Local $i, $a, $fh, $line

    If StringInStr($_fName, ":") Then
	$a = $_fName
    Else
	if StringLeft($_fName, 1) == "." Then
	    $_fName = StringTrimLeft($_fName, 2)
	elseif StringLeft($_fName, 1) == "\" Then
	    $_fName = StringTrimLeft($_fName, 1)
	endif
	$a = @WorkingDir & "\" & $_fName
    EndIf
    $fh = FileOpen ($a, 0) ; open file in read mode
    If @error Then
	Consolewrite(Stringformat("(Parse TX file) No file '%s'.\r\n", $a))
	Return -1
    EndIf
    While True
	$line = FileReadLine($fh)
	if @error Then ExitLoop
	if stringLen($line) = 0 Then ContinueLoop
	if $line = @CRLF or $line = @LF or $line = @CR Then ContinueLoop
	Consolewrite(Stringformat("(Parse TX file) Line: '%s'.\r\n", $line))
	parseString($line)
;~ 	$cmd[$cmdNo] = $line
;~ 	$_cmdNo = $_cmdNo + 1
    WEnd
    FileClose($fh)
    Return 0
EndFunc


Func sendData($_str, $_maxRX = 2048, $_first = 1000, $_next = 100, $DEBUG = 1)
;~     Local $ending = ""
    Local $tx = $_str
;~     If GUICtrlRead ($checkTX_CR) = $GUI_CHECKED Then
;~ 	$ending &= @CR
;~ 	If GUICtrlRead ($checkCRLF) = $GUI_CHECKED Then
;~ 	    $ending &= @LF
;~ 	EndIf
;~     EndIf
;~     $tx &= $ending
    ; send the string out
    If StringLen($tx) = 0 Then Return 0 ; skip if empty string
    If $DEBUG Then ConsoleWrite("--> " & $tx)
    If $logEnabled Then FileWrite($hLog, $tx)

    If $ConOpen = $connectCOM Then
	_tx($hCOM,$tx)
    ElseIf $ConOpen = $connectLAN Then
	Local $nTimeCtr = 0, $i, $str
	$out = ""
	; send string
	TCPSend($hCOM,$tx)
	If @error Then
	    ConsoleWrite ("Error: Could not send data to TCP socket." & @CRLF)
	    toggleLANconnection()
	    SetError (1)
	    Return -1
	EndIf
    ElseIf $ConOpen = $connectVISA Then
	$out = _viExecCommand($hCOM, $tx, $_first)
	writeRXdata($out)
    ElseIf $DUMMY_FILL Then
	;
    Else
	Return 0 ; not connected
    EndIf
    ; Update TX output field
    $editTXcount += StringLen($tx)
    If $editTXcount > $cMaxEditLen Then
	$editTXcount = StringLen($tx)
	GUICtrlSetData ($editTX, "")
    EndIf
    _GUICtrlEdit_AppendText($editTX, $tx)
    Return 0
EndFunc




Func tcpRXwait($_sock, $BufferSize = 1024, $firstWaitTime = 500,$charWaitTime = 100)
    Local $rxbuf = ""
    local $tmr=TimerInit(), $charTime, $firstChar = 0
    Do
	$str = TCPRecv($_sock, $BufferSize)
	if $str <> "" then
	    $rxbuf &= $str
	    $firstChar = 1
	    $charTime = TimerInit()
	Else
	    If $firstChar = 0 AND TimerDiff($tmr) > $firstWaitTime Then
		Return($rxbuf)
	    ElseIf $firstChar = 1 AND TimerDiff($charTime) > $charWaitTime Then
		Return($rxbuf)
	    EndIf
	EndIf
    Until $BufferSize <= StringLen($rxbuf)
    Return($rxbuf)
EndFunc


Func readRXbuffer()
    Local $rxbuf = ""
    If $ConOpen = $connectCOM Then
	$rxbuf = _rxwait2($hCOM,4000,0,0)
    ElseIf $ConOpen = $connectLAN Then
	$rxbuf = TCPRecv($hCOM, 2000)
;~ 	$rxbuf = tcpRXwait($hCOM)
    EndIf
    If $rxbuf <> "" Then writeRXdata($rxbuf)
EndFunc

Func EditRXappendText($_s, $DEBUG = 0)
    Local $curCtrl = ControlGetFocus("")
;~    _GUICtrlEdit_AppendText($editRX, $s2)
    _GUICtrlRichEdit_AppendText($editRX, $_s)
    _GUICtrlRichEdit_SetSel($editRX, $EscSel, -1)
    _GUICtrlRichEdit_SetCharBkColor($editRX, $REtxtBgColor)
    _GUICtrlRichEdit_SetCharColor($editRX, $REtxtColor)
    _GUICtrlRichEdit_SetFont($editRX, 9, "Courier New")
    _GUICtrlRichEdit_SetCharAttributes($editRX, $REtxtAttr)
    _GUICtrlRichEdit_SetSel($editRX, -1, -1)
    If $DEBUG Then ConsoleWrite(StringFormat("Write '%s'", $_s))
    ControlFocus("", "", $curCtrl)
EndFunc

Func writeRXdata($_str, $blockES = 0, $DEBUG = 1)
    Local $curCtrl = ControlGetFocus("")
    If $logEnabled Then FileWrite($hLog, $_str)
    $editRXcount += StringLen($_str)
    If $editRXcount > $cMaxEditLen Then
	$editRXcount = StringLen($_str)
;~ 	GUICtrlSetData ($editRX, "")
	_GUICtrlRichEdit_SetText ($editRX, "")
	$EscSel = 0
    EndIf
    ; parse terminal escape sequences
    Local $i, $ch, $s, $s2
    $s = StringSplit($_str, "")
    $s2 = ""
    For $i = 1 To $s[0]
	$ch = $s[$i]
	if $ch = Chr(Dec("1B")) Then
	    $escOn = 1
	    $escStr = ""
	    If $s2 <> "" Then
		EditRXappendText($s2,$DEBUG)
	    Endif
	    $s2 = ""
	    ContinueLoop
	ElseIf $escOn Then
	    $escStr &= $ch
	    If StringLen ($escStr) == 1 and $ch == "[" Then
		; long ES
	    Elseif StringLen ($escStr) == 1 Then
		; finish ES and exit
		If $DEBUG Then ConsoleWrite(StringFormat("ES = '%s'\r\n", $escStr))
		If not $blockES Then ParseES($escStr, $DEBUG)
		$escOn = 0;
	    Else ; next characters
		; detect end of ES
		If StringIsDigit($ch) Or $ch == ";" Or $ch == '"' Then
		    ; ES continues
		    ContinueLoop
		Else
		    If $DEBUG Then ConsoleWrite(StringFormat("ES = '%s'\r\n", $escStr))
		    If not $blockES Then ParseES($escStr, $DEBUG)
		    $escOn = 0;
		EndIf
	    EndIf
	    ContinueLoop;
	Else
	    ;if $ch <> @LF Then
	    $s2 &= $ch
	endif
    Next
    if $s2 <> "" Then
	EditRXappendText($s2,$DEBUG)
    EndIf
    ControlFocus("", "", $curCtrl)
EndFunc


;~ Global $REtxtBgColor = $DefREtxtBgColor
;~ Global $REbgColor = $DefREbgColor
;~ Global $REtxtColor = $DefREtxtColor

Func ParseES ($_str, $DEBUG = 0)
    If StringCompare($_str, "[2J") == 0 Then
	_GUICtrlRichEdit_SetText ($editRX, "")
	$EscSel = 0
	Return
    ElseIf StringLeft($_str,1) = "[" and StringRight($_str,1) = "m" Then
	Local $prevBGColor = $REtxtBgColor
	Local $prevTxtColor = $REtxtColor
	Local $prevAttr = $REtxtAttr
	If $DEBUG then ConsoleWrite(StringFormat("Set Display Attributes escape sequence detected.\r\n"))
	$_str = StringTrimLeft (StringTrimRight($_str,1),1)
	local $arg = StringSplit($_str, ";")
	Local $params, $i
	$params = 0
	For $i = 1 to $arg[0]
	    if $arg[$i] >= 30 And $arg[$i] < 40 Then
		; foreground color
		$REtxtColor = EscCodeToColor($arg[$i] - 30)

	    ElseIf $arg[$i] >= 40 And $arg[$i] < 50 Then
		$REtxtBgColor = EscCodeToColor($arg[$i] - 40)
	    ElseIf $arg[$i] >= 0 And $arg[$i] <= 8 Then
		$params += BitShift (1, -$arg[$i])
	    EndIf
	Next
	If $DEBUG then ConsoleWrite(StringFormat("Params = %d\r\n",Binary($params)))
	If BitAND( $params, BitShift(1,0)) Then
	    ; reset all
	    $REtxtBgColor = $DefREtxtBgColor
	    $REbgColor = $DefREbgColor
	    $REtxtColor = $DefREtxtColor
	    $REtxtAttr = "-bo-di-em-hi-im-it-li-ou-pr-re-sh-sm-st-sb-sp-un-al"
	    If $DEBUG then ConsoleWrite(StringFormat("Set Default.\r\n"))

	ElseIf BitAND( $params, BitShift(1,-1)) Then
	    $REtxtColor = BrightColor($REtxtColor, $DEBUG)
	ElseIf BitAND( $params, BitShift(1,-2)) Then
	    ;$REtxtColor = DimColor($REtxtColor, $DEBUG)
	ElseIf BitAND( $params, BitShift(1,-4)) Then
	    Local $a = StringInStr ($REtxtAttr, "un")
	    $REtxtAttr = StringReplace($REtxtAttr, $a -1, "+")
	ElseIf BitAND( $params, BitShift(1,-5)) Then
	    ; blink not supported
	ElseIf BitAND( $params, BitShift(1,-7)) Then
	    Local $a = $REtxtColor
	    $REtxtColor = $REtxtBgColor
	    $REtxtBgColor = $a
	ElseIf BitAND( $params, BitShift(1,-8)) Then
	    Local $a = StringInStr ($REtxtAttr, "hi")
	    $REtxtAttr = StringReplace($REtxtAttr, $a -1, "+")

	EndIf
	If $DEBUG Then ConsoleWrite(StringFormat("New FG = %s, BG = %s\r\n", Hex($REtxtColor), Hex($REtxtBgColor)))
	If $prevBGColor <> $REtxtBgColor Or $prevTxtColor <> $REtxtColor Or $prevAttr <> $REtxtAttr Then
	    Local $sel = _GUICtrlRichEdit_GetSel($editRX)
	    $EscSel = $sel[1]
	EndIf
	Return
    EndIf

EndFunc

Func DimColor ($_c, $DEBUG =0)
    Local $r, $g, $b
    $r = Floor($_c/65536)
    $g = Floor (($_c - $r * 65536) / 256)
    $b = Mod ($_c, 256)
    If $DEBUG then ConsoleWrite(StringFormat("Dim detected - #%s (%d,%d,%d) -> ",Hex($_c, 6), $r,$g,$b))
    $r = $r / 2
    $g = $g / 2
    $b = $b / 2
    $_c = $r * 65536 + $g * 256 + $b
    If $DEBUG then ConsoleWrite(StringFormat(" #%s (%d,%d,%d)\n",Hex($_c), $r,$g,$b))
    return $_c
EndFunc

Func BrightColor ($_c, $DEBUG = 0)
    Local $r, $g, $b
    $r = Floor($_c/65536)
    $g = Floor (($_c - $r * 65536) / 256)
    $b = Mod ($_c, 256)
    If $DEBUG then ConsoleWrite(StringFormat("Bright detected - #%s (%d,%d,%d) -> ",Hex($_c, 6), $r,$g,$b))
    $r = $r * 2
    $g = $g * 2
    $b = $b * 2
    If $r >=256 Then $r = 255
    If $g >=256 Then $g = 255
    If $b >=256 Then $b = 255
    $_c = $r * 65536 + $g * 256 + $b
    If $DEBUG then ConsoleWrite(StringFormat(" #%s (%d,%d,%d)\n",Hex($_c), $r,$g,$b))
    return $_c
EndFunc


Func EscCodeToColor($_c)
    Switch($_c)
	Case 0
	    Return 0x000000
	Case 1
	    Return 0x800000
	Case 2
	    Return 0x008000
	Case 3
	    Return 0x808000
	Case 4
	    Return 0x000080
	Case 5
	    Return 0x800080
	Case 6
	    Return 0x008080
	Case 7
	    Return 0x808080
	Case Else
	    Return 0x000000
    EndSwitch
EndFunc

Func toggleCOMconnection()
    If $ConOpen = $connectCOM Then
	_CloseComm($hCOM)
	GUICtrlSetData ($bConnect, "Connect")
	$ConOpen = $connectNone
	Return 0
    Else
	; get port
	Local $port = GUICtrlRead ($cCOM)
	If stringinstr($port,"COM") = 1 Then $port = StringTrimLeft($port,3)
	; get baudrate
	Local $br = GUICtrlRead ($cBaud)
	; get data format
	Local $data = GUICtrlRead ($cData)
	; get parity
	Local $par = _GUICtrlComboBox_GetCurSel($cParity) ; get index of the selected item
	; get stop bit
	Local $stop = _GUICtrlComboBox_GetCurSel ($cStopbits) ; get index of the selected item
	; get handshake
	Local $hshake = GUICtrlRead ($cHandShake)
	;Func _OpenComm($CommPort, $CommBaud = '4800', $CommBits = '8', $CommParity = '0', $CommStop = '0', $CommCtrl = '0011', $DEBUG = '0')
	$hCOM = _OpenComm($port,$br,$data, $par, $stop)
	If $hCOM = -1 Then
	    SetError(1)
	    Return -1
	EndIf
	regStoreCOMport()
	$ConOpen = $connectCOM
	GUICtrlSetData ($bConnect, "Disconnect")
	Return 0
    EndIf
EndFunc

Func toggleLANconnection()
    If $ConOpen = $connectLAN Then
	TCPCloseSocket ($hCOM)
	GUICtrlSetData ($bConnect, "Connect")
	$ConOpen = $connectNone
	Return 0
    Else
	; get port
	Local $port = GUICtrlRead ($iPort)
	If $port <> String(Number($port)) Then
	    ConsoleWrite('Port given is not a number.' & @crlf) ;### Debug Console\
	    Return -1
	Elseif $port = 0 Or $port > 65355 Then
	    ConsoleWrite('Port exceeds possible range.' & @crlf) ;### Debug Console\
	    Return -1
	EndIf
	; get IP address
	Local $ip = GUICtrlRead ($iIP)
	If Not isIP($ip) Then
	    ConsoleWrite('Trying to resolve name ' & $ip & @crlf) ;### Debug Console\
	    $ip = TCPNameToIP ($ip)
	    If @error Then
		ConsoleWrite('Error resolving name.' & @crlf) ;### Debug Console\
		Return -1
	    EndIf
	EndIf
	ConsoleWrite('Connecting to ' & $ip & ', port ' & $port & @crlf) ;### Debug Console
	$hCOM = TCPConnect ($ip, $port)
	If $hCOM = -1 Or @error Then
	    ConsoleWrite('Error: Connection failed!' & @crlf) ;### Debug Console
	    SetError(1)
	    Return -1
	EndIf
	regStoreLAN()
	$ConOpen = $connectLAN
	GUICtrlSetData ($bConnect, "Disconnect")
	Return 0
    EndIf
EndFunc


Func toggleVISAconnection()
    If $ConOpen = $connectVISA Then
	_viClose($hCOM)
	GUICtrlSetData ($bConnect, "Connect")
	$ConOpen = $connectNone
	Return 0
    Else
	; get port
	Local $addr = GUICtrlRead ($iVISAaddr)
	$hCOM = _viOpen($addr)
	If @error Then
	    SetError(1)
	    Return -1
	EndIf
	regStoreVISA()
	$ConOpen = $connectVISA
	GUICtrlSetData ($bConnect, "Disconnect")
	Return 0
    EndIf
EndFunc


Func closeSocket()
    If $ConOpen = $connectCOM Then toggleCOMconnection()
    If $ConOpen = $connectLAN Then toggleLANconnection()
    If $ConOpen = $connectVISA Then toggleVISAconnection()
EndFunc


Func scanCOMports()
    Local $regDir = "HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM"
    Local $var, $key, $str = "", $init = "", $DEBUG = 1
    $COMcount = 0
    For $i = 1 To 200
	$key = RegEnumVal($regDir, $i)
	If @error <> 0 Then ExitLoop ; no more ports found
	$var = RegRead($regDir, $key)
	If @error <> 0 Then ContinueLoop ; error, skip
	If $DEBUG Then ConsoleWrite("SubKey #" & $i & " : " & $key & " = "& $var & @CRLF)
	$COMlist[$COMcount] = $var
	$COMcount += 1
    Next
    GUICtrlSetData($cCOM,"")
    For $i = 0 To $COMcount-1
	$str &= $COMlist[$i] & "|"
    Next
    $str = StringTrimRight($str, 1)
    $init = ""
    If $COMcount > 0 Then $init = $COMlist[0]
    GUICtrlSetData($cCOM,$str,$init)
EndFunc


