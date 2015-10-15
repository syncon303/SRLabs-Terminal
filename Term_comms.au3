
; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.
#include <GuiRichEdit.au3>
#include <GuiEdit.au3>
#include <Array.au3>


Func sendInputData()
    Local $tx = GUICtrlRead($InputTX)
    Local $ret
    checkHistory($tx)
    ;    $ret = sendData($tx)
    $ret = parseString($tx)
    GUICtrlSetData($InputTX, "")
    $histLastString = $tx
    $InputTXscratch = ""
    $noHist = True
    Return $ret
EndFunc   ;==>sendInputData



; Parse following meta strings:
; - %DLY1234 -> Delay in milliseconds
; - #xxx -> ASCII character in decimal format
; - $xx  -> ASCII character in hexadecimal format
; - %M12  -> insert macro


Func parseString($_str, $_reentrNo = 0, $DEBUG = 1)
    Const $MAX_CMDS = 100
    Const $MAX_REENTRY = 10
    Local $cmd[$MAX_CMDS], $meta[$MAX_CMDS], $cmdNo = 0, $i
    For $i = 0 To $MAX_CMDS - 1
        $cmd[$i] = ""
        $meta[$i] = ""
    Next
    Local $newline = ""
    If GUICtrlRead($checkTX_CR) = $GUI_CHECKED Then
        Local $newline = @CR
        If GUICtrlRead($checkCRLF) = $GUI_CHECKED Then $newline &= @LF
    EndIf
    ; parse input string
    Local $quote = "", $token = "", $char = "", $str = $_str, $reentry = 0
    While True
        If Not $reentry Then
            If StringLen($str) = 0 Then ExitLoop
            $char = StringLeft($str, 1) ; store first remaining character of the string
            $str = StringTrimLeft($str, 1)
        Else
            $reentry = 0
        EndIf
        ; check for quotation marks
        If $quote <> "" Then
            If $char = $quote Then
                $quote = ""
            Else
                $cmd[$cmdNo] &= $char
            EndIf
            ; check for tokens
        ElseIf $token <> "" Then
            If $char = $token Then ; repeated token, ignore and put single characater to output string
                $token = ""
                $cmd[$cmdNo] &= $char
                ContinueLoop
            ElseIf $char = "$" Or $char = "#" Or $char = "%" Then
                ; format error handling, interpret previous token as character and set new token
                $cmd[$cmdNo] &= $token
                $token = $char
                ContinueLoop
            ElseIf $token = "$" Then
                ; check for double hexadecimal characters
                Local $a = $char & StringLeft($str, 1)
                If StringLen($a) == 2 And StringIsXDigit($a) Then
                    $cmd[$cmdNo] &= Chr(Dec($a))
                    $str = StringTrimLeft($str, 1)
                    $token = ""
                    ; check for newline sequence
;~                     If $newline <> "" And StringRight($cmd[$cmdNo], StringLen($newline)) = $newline Then
;~                         ; assume new command
;~                         $cmdNo += 1
;~                     EndIf
                Else
                    ConsoleWrite("wowowow" & @CRLF)
                    $cmd[$cmdNo] &= $token
                    $token = ""
                    $reentry = 1
                EndIf
            ElseIf $token = "#" Then
                ; check for three decimal characters
                Local $a = $char & StringLeft($str, 2)
                Local $b = $char & StringLeft($str, 1)
                If StringLen($a) = 3 And StringIsDigit($a) Then
                    $cmd[$cmdNo] &= Chr($a)
                    $str = StringTrimLeft($str, 2)
                    $token = ""
                ElseIf $b == "--" Then ; check for comment
                    $token = ""
                    $str = "" ; comment till end of line, so artificially end the parsing
                else
                    $cmd[$cmdNo] &= $token
                    $token = ""
                    $reentry = 1
                EndIf
            ElseIf $token = "%" Then
                If $char = "M" Then
                    Local $a = StringLeft($str, 2)
                    Local $b = StringLeft($str, 3)
                    If StringLen($a) = 2 And StringIsDigit($a) and $a <= $MACRO_NUMBER Then
                        If $cmd[$cmdNo] <> "" Then $cmdNo += 1
                        $meta[$cmdNo] = "macro"
                        $cmd[$cmdNo] = Int($a)
                        $cmdNo += 1
                        $str = StringTrimLeft($str, 2)
                    Else
                        $cmd[$cmdNo] &= $token
                        $reentry = 1
                    EndIf
                    $token = ""
                ElseIf $char = "D" Then
                    Local $a = StringLeft($str, 2)
                    Local $b = StringMid($str, 3, 4)
                    If $a = "LY" And StringLen($b) = 4 And StringIsDigit($b) Then
                        If $cmd[$cmdNo] <> "" Then $cmdNo += 1
                        $meta[$cmdNo] = "delay"
                        $cmd[$cmdNo] = Int($b)
                        $cmdNo += 1
                        $str = StringTrimLeft($str, 6)
                    Else
                        $cmd[$cmdNo] &= $token
                        $reentry = 1
                    EndIf
                    $token = ""
                ElseIf $char = "T" Then
                    Local $a = $char & StringLeft($str, 3)
                    Local $b = $char & StringLeft($str, 5)
                    Local $c = $char & StringLeft($str, 6)
                    if $b = "TAILON" Then
                        GUICtrlSetState($checkEnableRXfilter,$GUI_CHECKED)
                        $str = StringTrimLeft($str, 5)
                    elseif $c = "TAILOFF" Then
                        GUICtrlSetState($checkEnableRXfilter,$GUI_UNCHECKED)
                        $str = StringTrimLeft($str, 6)
                    ElseIf $a = "TAIL" Then
                        Local $trimLen = 3
                        Local $b = StringTrimLeft($str, 3)
                        $trimLen = $trimLen + StringLen($b)
                        $b = StringStripWS($b, 1)
                        $trimLen = $trimLen - StringLen($b)
                        If StringLeft($b, 1) = "=" Then
                            $b = StringTrimLeft($b, 1)
                            $trimLen = $trimLen + 1
                        EndIf
                        Local $quote = ""
                        If StringInStr($b, "'") Then
                            $quote = "'"
                        ElseIf StringInStr($b, "'") Then
                            $quote = '"'
                        EndIf
                        If $quote <> "" Then
                            $trimLen = $trimLen + StringInStr($b, $quote)
                            Local $c = StringSplit($b, $quote)
                            $b = $c[2]
                            $trimLen = $trimLen + StringLen($b) + 1
                        EndIf
                        If $DEBUG Then ConsoleWrite(StringFormat("Using tail: %s\r\n", $b))
                        GUICtrlSetData($InTail, $b)
                        $str = StringTrimLeft($str, $trimLen)
                        If $DEBUG Then ConsoleWrite(StringFormat("Remaining string: '%s'\r\n", $str))
                        ExitLoop
                    EndIf
                ElseIf $char = "C" Or $char = "L" Then
                    Local $a = $char & StringLeft($str, 1)
                    Local $b = $char & StringLeft($str, 3)
                    $token = ""
                    If $b = "CRLF" Then
                        $str = StringTrimLeft($str, 3)
                        $cmd[$cmdNo] &= Chr(13) & Chr(10)
                        ; check for newline sequence
                        If $newline == @CRLF Then $cmdNo += 1 ; assume new command
                    ElseIf $a = "CR" Then
                        $str = StringTrimLeft($str, 1)
                        $cmd[$cmdNo] &= Chr(13)
                        If $newline == @CR Then $cmdNo += 1 ; assume new command
                    ElseIf $a = "LF" Then
                        $str = StringTrimLeft($str, 1)
                        $cmd[$cmdNo] &= Chr(10)
                        ; check for newline sequence
                        If $newline <> "" And StringRight($cmd[$cmdNo], StringLen($newline)) = $newline Then
                            ; assume new command
                            $cmdNo += 1
                        EndIf
                    Else
                        $cmd[$cmdNo] &= $token
                        $reentry = 1
                    EndIf
                ElseIf $char = "F" Then
                    ; check if this should be file entry
                    Local $a = $char & StringLeft($str, 3)
                    If $a = "FILE" Then
                        If $DEBUG Then ConsoleWrite(StringFormat("Here.\r\n"))
                        Local $trimLen = 3
                        Local $b = StringTrimLeft($str, 3)
                        $trimLen = $trimLen + StringLen($b)
                        $b = StringStripWS($b, 1)
                        $trimLen = $trimLen - StringLen($b)
                        If StringLeft($b, 1) = "=" Then
                            $b = StringTrimLeft($b, 1)
                            $trimLen = $trimLen + 1
                        EndIf
                        Local $quote = ""
                        If StringInStr($b, "'") Then
                            $quote = "'"
                        ElseIf StringInStr($b, '"') Then
                            $quote = '"'
                        EndIf
                        If $quote <> "" Then
                            $trimLen = $trimLen + StringInStr($b, $quote)
                            Local $c = StringSplit($b, $quote)
                            $b = $c[2]
                            $trimLen = $trimLen + StringLen($b) + 1
                        EndIf
                        If $DEBUG Then ConsoleWrite(StringFormat("Using filename: %s\r\n", $b))
                        ParseTXfile($b, $cmd, $cmdNo)
                        $str = StringTrimLeft($str, $trimLen)
                        If $DEBUG Then ConsoleWrite(StringFormat("Remaining string: '%s'\r\n", $str))
                        ExitLoop
                    EndIf
                EndIf
            EndIf
        ElseIf $char = "$" Or $char = "#" Or $char = "%" Then
            $token = $char
            ContinueLoop
        ElseIf $char = "'" Or $char = '"' Then
            $quote = $char
            ContinueLoop
        Else
            $cmd[$cmdNo] &= $char ; add character to command string
        EndIf
    WEnd
    If $cmd[$cmdNo] <> "" Then $cmdNo += 1
    ; add newlines if needed
    If $newline <> "" Then
        Local $len = StringLen($newline)
        For $i = 0 To $cmdNo - 1
            If $meta[$i] <> "" Then ContinueLoop
            If StringRight($cmd[$i], $len) == $newline Then ContinueLoop
            $cmd[$i] &= $newline
        Next
    EndIf

    If $DEBUG Then
        Local $dbgoffset = "", $z
        For $z = 1 To $_reentrNo
            $dbgoffset &= "  "
        Next
        ConsoleWrite(StringFormat("%sOriginal data: '%s', ", $dbgoffset, $_str))
        ConsoleWrite(StringFormat("%s>Parse table: \r\n", $dbgoffset))
        For $i = 0 To $cmdNo - 1
            Local $c = $cmd[$i]
            If $meta[$i] <> "" Then
                ConsoleWrite(StringFormat("%s  %d -> meta: '%s' = '%s'\r\n", $dbgoffset, $i, $meta[$i], $cmd[$i]))
            Else
                ConsoleWrite(StringFormat("%s  %d -> command: '%s'\r\n", $dbgoffset, $i, $c))
            EndIf
        Next
    EndIf
    ; send out
    For $i = 0 To $cmdNo - 1
        If $meta[$i] = "delay" Then
            Local $tInit = TimerInit()
            While TimerDiff($tInit) < $cmd[$i]
                MainLoop() ; process input during wait
            WEnd
        ElseIf $meta[$i] = "macro" Then
            If $_reentrNo < $MAX_REENTRY Then
                ; get values for macro and parse the macro
                Local $m = $macroString[$cmd[$i] - 1]
                parseString($m, $_reentrNo + 1)
            EndIf
        Else
            Local $ret = sendData($cmd[$i])
        EndIf
    Next
;~     If $_reentrNo = 0 And $cmdNo >= 1 then
;~ 	Local $c = ""
;~ 	If GUICtrlRead ($checkTX_CR) = $GUI_CHECKED Then
;~ 	    $c &= @CR
;~ 	    If GUICtrlRead ($checkCRLF) = $GUI_CHECKED Then $c &= @LF
;~ 	EndIf
;~ 	local $ret = sendData ($c)
;~     EndIf
EndFunc   ;==>parseString


Func ParseTXfile($_fName, ByRef $cmd, ByRef $cmdNo)
    Local $i, $a, $fh, $line

    If StringInStr($_fName, ":") Then
        $a = $_fName
    Else
        If StringLeft($_fName, 1) == "." Then
            $_fName = StringTrimLeft($_fName, 2)
        ElseIf StringLeft($_fName, 1) == "\" Then
            $_fName = StringTrimLeft($_fName, 1)
        EndIf
        $a = @WorkingDir & "\" & $_fName
    EndIf
    $fh = FileOpen($a, 0) ; open file in read mode
    If @error Then
        ConsoleWrite(StringFormat("(Parse TX file) No file '%s'.\r\n", $a))
        Return -1
    EndIf
    While True
        $line = FileReadLine($fh)
        If @error Then ExitLoop
        If StringLen($line) = 0 Then ContinueLoop
        If $line = @CRLF Or $line = @LF Or $line = @CR Then ContinueLoop
        If StringLeft(StringStripWS($line, 1), 1) = "#" Then ContinueLoop ; strip comments
        ConsoleWrite(StringFormat("(Parse TX file) Line: '%s'.\r\n", $line))
        parseString($line)
;~ 	$cmd[$cmdNo] = $line
;~ 	$_cmdNo = $_cmdNo + 1
    WEnd
    FileClose($fh)
    Return 0
EndFunc   ;==>ParseTXfile

Global $editTXnewline = 0

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
    If BitAND(GUICtrlRead($checkEnableRXfilter), $GUI_CHECKED) == $GUI_CHECKED Then
        While $TXsent <> 0
            If TimerDiff($TXtimeStamp[$TXsent]) > 2000 Then
                If $DEBUG Then ConsoleWrite(StringFormat("Tail detection timeout.\nBuffer content = '%s'\n", $rxDelimStr))
                $TXsent -= 1
                Local $len = StringLen($rxDelimStr) - StringLen($RXtail)
                If $len > 0 Then
                    $rxDelimStr = StringTrimLeft($rxDelimStr, $len)
                EndIf
                If $TXsent == 0 Then ExitLoop
            EndIf
            MainLoop()
        WEnd
    EndIf
    If StringLen($tx) = 0 Then Return 0 ; skip if empty string
    If $DEBUG Then ConsoleWrite("-->(" & StringLen($tx) & "chars) " & $tx )
    If $logEnabled Then FileWrite($hLog, $tx)

    If $ConOpen = $connectCOM Then
        _tx($hCOM, $tx)
    ElseIf $ConOpen = $connectLAN Then
        Local $nTimeCtr = 0, $i, $str
        $out = ""
        ; send string
        TCPSend($hCOM, $tx)
        If @error Then
            ConsoleWrite("Error: Could not send data to TCP socket." & @CRLF)
            toggleLANconnection()
            SetError(1)
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
    If BitAND(GUICtrlRead($checkEnableRXfilter), $GUI_CHECKED) == $GUI_CHECKED Then
        $TXsent += 1
        $TXtimeStamp[$TXsent] = TimerInit()
    EndIf
    $editTXcount += StringLen($tx)
    If $editTXcount > $cMaxEditLen Then
        $editTXcount = StringLen($tx)
        GUICtrlSetData($editTX, "")
    EndIf
    if Stringlen($tx) Then
        if $editTXnewline Then
            $tx = @LF & $tx
            $editTXnewline = 0
        EndIf
        if stringRight($tx, 1) = @LF Then
            $tx = StringTrimRight($tx,1)
            $editTXnewline = 1
        EndIf
        _GUICtrlEdit_AppendText($editTX, $tx)
    EndIf
    Return 0
EndFunc   ;==>sendData


Func tcpRXwait($_sock, $BufferSize = 1024, $firstWaitTime = 500, $charWaitTime = 100)
    Local $rxbuf = ""
    Local $tmr = TimerInit(), $charTime, $firstChar = 0
    Do
        $str = TCPRecv($_sock, $BufferSize)
        If $str <> "" Then
            $rxbuf &= $str
            $firstChar = 1
            $charTime = TimerInit()
        Else
            If $firstChar = 0 And TimerDiff($tmr) > $firstWaitTime Then
                Return ($rxbuf)
            ElseIf $firstChar = 1 And TimerDiff($charTime) > $charWaitTime Then
                Return ($rxbuf)
            EndIf
        EndIf
    Until $BufferSize <= StringLen($rxbuf)
    Return ($rxbuf)
EndFunc   ;==>tcpRXwait


Func readRXbuffer()
    Local $DEBUG = 0
    Local $rxbuf = ""
    If $ConOpen = $connectCOM Then
        $rxbuf = _rxwait2($hCOM, 4000, 0, 0)
    ElseIf $ConOpen = $connectLAN Then
        $rxbuf = TCPRecv($hCOM, 2000)
;~ 	$rxbuf = tcpRXwait($hCOM)
    EndIf
    If $rxbuf == "" Then Return
    writeRXdata($rxbuf)
    ; check if tail can be found in the received string
    If BitAND(GUICtrlRead($checkEnableRXfilter), $GUI_CHECKED) == $GUI_CHECKED And $TXsent <> 0 Then
        $rxDelimStr &= $rxbuf
        Local $s = StringInStr($rxDelimStr, $RXtail)
        If $s <> 0 Then
            $TXsent -= 1
            $rxDelimStr = StringTrimLeft($rxDelimStr, $s + StringLen($RXtail))
            If $DEBUG Then ConsoleWrite(StringFormat("Tail detected.\n"))

        EndIf
    EndIf
EndFunc   ;==>readRXbuffer


Func convertBlanks($_s)
    local $so = $_s
    $so = StringReplace($so,@TAB, "\t")
    $so = StringReplace($so,@CR, "\r")
    $so = StringReplace($so,@LF, "\n" & @CRLF)
    $so = StringReplace($so,Chr(7), "\a")
    $so = StringReplace($so,Chr(11), "\v")
    Return $so
EndFunc


Func EditRXappendText($_s, $DEBUG = 0)
    Local $curCtrl = ControlGetFocus("")
;~    _GUICtrlEdit_AppendText($editRX, $s2)
    ; check if show blanks is checked
    If BitAND(GUICtrlRead($chkShowBlanks), $GUI_CHECKED) = $GUI_CHECKED Then
        $_s = convertBlanks($_s)
    EndIf
    _GUICtrlRichEdit_AppendText($editRX, $_s)
    _GUICtrlRichEdit_SetSel($editRX, $EscSel, -1)
    _GUICtrlRichEdit_SetCharBkColor($editRX, $REtxtBgColor)
    _GUICtrlRichEdit_SetCharColor($editRX, $REtxtColor)
    _GUICtrlRichEdit_SetFont($editRX, 9, "Courier New")
    _GUICtrlRichEdit_SetCharAttributes($editRX, $REtxtAttr)
    _GUICtrlRichEdit_SetSel($editRX, -1, -1)
    If $DEBUG Then ConsoleWrite(StringFormat("Write '%s'", $_s))
    ControlFocus("", "", $curCtrl)
EndFunc   ;==>EditRXappendText


Func writeRXdata($_str, $blockES = 0, $DEBUG = 0, $DEBUG_ES = 0)
    Local $curCtrl = ControlGetFocus("")
    If $logEnabled Then FileWrite($hLog, $_str)
    $editRXcount += StringLen($_str)
    If $editRXcount > $cMaxEditLen Then
        $editRXcount = StringLen($_str)
;~ 	GUICtrlSetData ($editRX, "")
        _GUICtrlRichEdit_SetText($editRX, "")
        $EscSel = 0
    EndIf
    ; parse terminal escape sequences
    Local $i, $ch, $s, $s2
    $s = StringSplit($_str, "")
    $s2 = ""
    For $i = 1 To $s[0]
        $ch = $s[$i]
        If $ch = Chr(Dec("1B")) Then
            $escOn = 1
            $escStr = ""
            If $s2 <> "" Then
                EditRXappendText($s2, $DEBUG)
            EndIf
            $s2 = ""
            ContinueLoop
        ElseIf $escOn Then
            $escStr &= $ch
            If StringLen($escStr) == 1 And $ch == "[" Then
                ; long ES
            ElseIf StringLen($escStr) == 1 Then
                ; finish ES and exit
                If $DEBUG Then ConsoleWrite(StringFormat("ES = '%s'\r\n", $escStr))
                If Not $blockES Then ParseES($escStr, $DEBUG_ES)
                $escOn = 0;
            Else ; next characters
                ; detect end of ES
                If StringIsDigit($ch) Or $ch == ";" Or $ch == '"' Then
                    ; ES continues
                    ContinueLoop
                Else
                    If $DEBUG Then ConsoleWrite(StringFormat("ES = '%s'\r\n", $escStr))
                    If Not $blockES Then ParseES($escStr, $DEBUG_ES)
                    $escOn = 0;
                EndIf
            EndIf
            ContinueLoop;
        Else
            ;if $ch <> @LF Then
            $s2 &= $ch
        EndIf
    Next
    If $s2 <> "" Then
        EditRXappendText($s2, $DEBUG)
    EndIf
    ControlFocus("", "", $curCtrl)
EndFunc   ;==>writeRXdata


;~ Global $REtxtBgColor = $DefREtxtBgColor
;~ Global $REbgColor = $DefREbgColor
;~ Global $REtxtColor = $DefREtxtColor

Func ParseES($_str, $DEBUG = 0)
    If StringCompare($_str, "[2J") == 0 Then
        _GUICtrlRichEdit_SetText($editRX, "")
        $EscSel = 0
        Return
    ElseIf StringLeft($_str, 1) = "[" And StringRight($_str, 1) = "m" Then
        Local $prevBGColor = $REtxtBgColor
        Local $prevTxtColor = $REtxtColor
        Local $prevAttr = $REtxtAttr
        If $DEBUG Then ConsoleWrite(StringFormat("Set Display Attributes escape sequence detected.\r\n"))
        $_str = StringTrimLeft(StringTrimRight($_str, 1), 1)
        Local $arg = StringSplit($_str, ";")
        Local $params, $i
        $params = 0
        For $i = 1 To $arg[0]
            If $arg[$i] >= 30 And $arg[$i] < 40 Then
                ; foreground color
                $REtxtColor = EscCodeToColor($arg[$i] - 30)

            ElseIf $arg[$i] >= 40 And $arg[$i] < 50 Then
                $REtxtBgColor = EscCodeToColor($arg[$i] - 40)
            ElseIf $arg[$i] >= 0 And $arg[$i] <= 8 Then
                $params += BitShift(1, -$arg[$i])
            EndIf
        Next
        If $DEBUG Then ConsoleWrite(StringFormat("Params = %d\r\n", Binary($params)))
        If BitAND($params, BitShift(1, 0)) Then
            ; reset all
            $REtxtBgColor = $DefREtxtBgColor
            $REbgColor = $DefREbgColor
            $REtxtColor = $DefREtxtColor
            $REtxtAttr = "-bo-di-em-hi-im-it-li-ou-pr-re-sh-sm-st-sb-sp-un-al"
            If $DEBUG Then ConsoleWrite(StringFormat("Set Default.\r\n"))

        ElseIf BitAND($params, BitShift(1, -1)) Then
            $REtxtColor = BrightColor($REtxtColor, $DEBUG)
        ElseIf BitAND($params, BitShift(1, -2)) Then
            ;$REtxtColor = DimColor($REtxtColor, $DEBUG)
        ElseIf BitAND($params, BitShift(1, -4)) Then
            Local $a = StringInStr($REtxtAttr, "un")
            $REtxtAttr = StringReplace($REtxtAttr, $a - 1, "+")
        ElseIf BitAND($params, BitShift(1, -5)) Then
            ; blink not supported
        ElseIf BitAND($params, BitShift(1, -7)) Then
            Local $a = $REtxtColor
            $REtxtColor = $REtxtBgColor
            $REtxtBgColor = $a
        ElseIf BitAND($params, BitShift(1, -8)) Then
            Local $a = StringInStr($REtxtAttr, "hi")
            $REtxtAttr = StringReplace($REtxtAttr, $a - 1, "+")

        EndIf
        If $DEBUG Then ConsoleWrite(StringFormat("New FG = %s, BG = %s\r\n", Hex($REtxtColor), Hex($REtxtBgColor)))
        If $prevBGColor <> $REtxtBgColor Or $prevTxtColor <> $REtxtColor Or $prevAttr <> $REtxtAttr Then
            Local $sel = _GUICtrlRichEdit_GetSel($editRX)
            $EscSel = $sel[1]
        EndIf
        Return
    EndIf

EndFunc   ;==>ParseES


Func DimColor($_c, $DEBUG = 0)
    Local $r, $g, $b
    $r = Floor($_c / 65536)
    $g = Floor(($_c - $r * 65536) / 256)
    $b = Mod($_c, 256)
    If $DEBUG Then ConsoleWrite(StringFormat("Dim detected - #%s (%d,%d,%d) -> ", Hex($_c, 6), $r, $g, $b))
    $r = $r / 2
    $g = $g / 2
    $b = $b / 2
    $_c = $r * 65536 + $g * 256 + $b
    If $DEBUG Then ConsoleWrite(StringFormat(" #%s (%d,%d,%d)\n", Hex($_c), $r, $g, $b))
    Return $_c
EndFunc   ;==>DimColor


Func BrightColor($_c, $DEBUG = 0)
    Local $r, $g, $b
    $r = Floor($_c / 65536)
    $g = Floor(($_c - $r * 65536) / 256)
    $b = Mod($_c, 256)
    If $DEBUG Then ConsoleWrite(StringFormat("Bright detected - #%s (%d,%d,%d) -> ", Hex($_c, 6), $r, $g, $b))
    $r = $r * 2
    $g = $g * 2
    $b = $b * 2
    If $r >= 256 Then $r = 255
    If $g >= 256 Then $g = 255
    If $b >= 256 Then $b = 255
    $_c = $r * 65536 + $g * 256 + $b
    If $DEBUG Then ConsoleWrite(StringFormat(" #%s (%d,%d,%d)\n", Hex($_c), $r, $g, $b))
    Return $_c
EndFunc   ;==>BrightColor


Func EscCodeToColor($_c)
    Switch ($_c)
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
EndFunc   ;==>EscCodeToColor


Func toggleCOMconnection()
    If $ConOpen = $connectCOM Then
        _CloseComm($hCOM)
        GUICtrlSetData($bConnect, "Connect")
        WinSetTitle($Terminal,"", $TermTitle)
        $ConOpen = $connectNone
        Return 0
    Else
        ; get port
        Local $port = GUICtrlRead($cCOM)
        If StringInStr($port, "COM") = 1 Then $port = StringTrimLeft($port, 3)
        ; get baudrate
        Local $br = GUICtrlRead($cBaud)
        ; get data format
        Local $data = GUICtrlRead($cData)
        ; get parity
        Local $par = _GUICtrlComboBox_GetCurSel($cParity) ; get index of the selected item
        ; get stop bit
        Local $stop = _GUICtrlComboBox_GetCurSel($cStopbits) ; get index of the selected item
        ; get handshake
        Local $hshake = GUICtrlRead($cHandShake)
        ;Func _OpenComm($CommPort, $CommBaud = '4800', $CommBits = '8', $CommParity = '0', $CommStop = '0', $CommCtrl = '0011', $DEBUG = '0')
        $hCOM = _OpenComm($port, $br, $data, $par, $stop)
        If $hCOM = -1 Then
            SetError(1)
            Return -1
        EndIf
        regStoreCOMport()
        $ConOpen = $connectCOM
        GUICtrlSetData($bConnect, "Disconnect")
        WinSetTitle($Terminal, "", StringFormat("%s | Connected to %s @ %s baud", $TermTitle,  GUICtrlRead ($cCOM), GUICtrlRead ($cBaud)) )
        Return 0
    EndIf
EndFunc   ;==>toggleCOMconnection


Func toggleLANconnection()
    If $ConOpen = $connectLAN Then
        TCPCloseSocket($hCOM)
        GUICtrlSetData($bConnect, "Connect")
        WinSetTitle($Terminal, "", $TermTitle)
        $ConOpen = $connectNone
        Return 0
    Else
        ; get port
        Local $port = GUICtrlRead($iPort)
        If $port <> String(Number($port)) Then
            ConsoleWrite('Port given is not a number.' & @CRLF) ;### Debug Console\
            Return -1
        ElseIf $port = 0 Or $port > 65355 Then
            ConsoleWrite('Port exceeds possible range.' & @CRLF) ;### Debug Console\
            Return -1
        EndIf
        ; get IP address
        Local $ip = GUICtrlRead($iIP)
        If Not isIP($ip) Then
            ConsoleWrite('Trying to resolve name ' & $ip & @CRLF) ;### Debug Console\
            $ip = TCPNameToIP($ip)
            If @error Then
                ConsoleWrite('Error resolving name.' & @CRLF) ;### Debug Console\
                Return -1
            EndIf
        EndIf
        ConsoleWrite('Connecting to ' & $ip & ', port ' & $port & @CRLF) ;### Debug Console
        $hCOM = TCPConnect($ip, $port)
        If $hCOM = -1 Or @error Then
            ConsoleWrite('Error: Connection failed!' & @CRLF) ;### Debug Console
            SetError(1)
            Return -1
        EndIf
        regStoreLAN()
        $ConOpen = $connectLAN
        GUICtrlSetData($bConnect, "Disconnect")
        WinSetTitle($Terminal, "", StringFormat("%s | Connected to %s : %d", $TermTitle,  GUICtrlRead ($iIP), GUICtrlRead ($iPort)) )
        Return 0
    EndIf
EndFunc   ;==>toggleLANconnection


Func toggleVISAconnection()
    If $ConOpen = $connectVISA Then
        _viClose($hCOM)
        GUICtrlSetData($bConnect, "Connect")
        WinSetTitle($Terminal, "", $TermTitle)
        $ConOpen = $connectNone
        Return 0
    Else
        ; get port
        Local $addr = GUICtrlRead($iVISAaddr)
        $hCOM = _viOpen($addr)
        If @error Then
            SetError(1)
            Return -1
        EndIf
        regStoreVISA()
        $ConOpen = $connectVISA
        GUICtrlSetData($bConnect, "Disconnect")
        WinSetTitle($Terminal, "", StringFormat("%s | Connected to %s", $TermTitle,  GUICtrlRead ($iVISAaddr)) )
        Return 0
    EndIf
EndFunc   ;==>toggleVISAconnection


Func closeSocket()
    If $ConOpen = $connectCOM Then toggleCOMconnection()
    If $ConOpen = $connectLAN Then toggleLANconnection()
    If $ConOpen = $connectVISA Then toggleVISAconnection()
EndFunc   ;==>closeSocket


Func scanCOMports()
    Local $regDir = "HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM"
    Local $var, $key, $str = "", $init = "", $DEBUG = 1
    $COMcount = 0
    For $i = 1 To 200
        $key = RegEnumVal($regDir, $i)
        If @error <> 0 Then ExitLoop ; no more ports found
        $var = RegRead($regDir, $key)
        If @error <> 0 Then ContinueLoop ; error, skip
        If $DEBUG Then ConsoleWrite("SubKey #" & $i & " : " & $key & " = " & $var & @CRLF)
        $COMlist[$COMcount] = $var
        $COMcount += 1
    Next
    GUICtrlSetData($cCOM, "")
    For $i = 0 To $COMcount - 1
        $str &= $COMlist[$i] & "|"
    Next
    $str = StringTrimRight($str, 1)
    $init = ""
    If $COMcount > 0 Then $init = $COMlist[0]
    GUICtrlSetData($cCOM, $str, $init)
EndFunc   ;==>scanCOMports


