
; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.
#include <Misc.au3>

Func readMacroFile()
    Local $path
    Local $hFile, $expectMacro = 1, $par = 0, $skip = 0, $no, $a
    Local $isBray = 0
    HotKeySet("{ENTER}")
    $path = FileOpenDialog("Open macro file", $macroDirectory, "Macro file (*.tmf)|Text files (*.txt)|All files (*.*)", 3, "*.tmf", $Terminal)
    HotKeySet("{ENTER}", "captureENTER")
    If @error Or $path = "" Then Return -1
    $hFile = FileOpen($path, 0)
    If @error Then Return -2
    While 1
        $line = FileReadLine($hFile)
        If @error Then ExitLoop
        If $expectMacro Then
            If StringInStr($line, "M") = 1 Then
                $line = StringTrimLeft($line, 1)
                $a = StringSplit($line, "P")
                If $a[0] > 2 Then ContinueLoop
                If $a[0] = 2 Then
                    If $a[2] = 1 Or $a[2] = 2 Then
                        $par = $a[2]
                    Else
                        ContinueLoop
                    EndIf
                EndIf
                If String(Number($a[1])) = $a[1] And ($a[1] > 0) And ($a[1] <= $MACRO_NUMBER) Then
                    ; if Br@y terminal file, overwrite current macro bank
                    if $isBray Then
                        $no = $BankFirst + $a[1] - 1
                    Else
                        $no = $a[1] - 1
                    EndIf
                    $expectMacro = 0
                Else
                    $par = 0
                    ContinueLoop
                EndIf
            Elseif StringInStr($line, "# Terminal macro file") == 1 then
                $isBray = 1
                CalcFirstMacroInBank() ; calculate $bankFirst variable
            EndIf
        Else
            $macroString[$no] = $line
                if $no >= $BankFirst and $no < $BankFirst + $MACRO_PER_BANK then
                        GUICtrlSetData($iMcr[$no-$BankFirst], $line)
                EndIf
            regStoreMacro($no)
            $par = 0
            $expectMacro = 1
        EndIf
    WEnd
    FileClose($hFile)
    Return 0
EndFunc   ;==>readMacroFile

Func writeMacroFile()
    Local $writePars = 1
    Local $path, $hFile, $i
    HotKeySet("{ENTER}")
    $path = FileSaveDialog("Save macro file", $macroDirectory, "Macro file (*.tmf)", 16, "*.tmf", $Terminal)
    HotKeySet("{ENTER}", "captureENTER")
    If @error Or $path = "" Then Return -1
    If StringRight($path, 4) <> ".tmf" Then $path = $path & ".tmf"
    $hFile = FileOpen($path, 10)
    If @error Then Return -2
    $path = extractPath($path)
    If Not @error Then
        RegWrite($REG_ROOT, "MacroDirectory", "REG_SZ", $path)
        $macroDirectory = $path
    EndIf
    FileWriteLine($hFile, "# SRLabs terminal macro file")
    For $i = 1 To $MACRO_NUMBER
        local $str
		$str = $macroString[$i - 1]
        if $str = "" then ContinueLoop
        FileWriteLine($hFile, "M" & $i)
        FileWriteLine($hFile, $str)
    Next
    FileClose($hFile)
    Return 0
EndFunc   ;==>writeMacroFile


Func toggleMacrosView()
    $ShowMacros = 1 - $ShowMacros
    if $ShowMacros == 0 and $MacrosFloat == 1 then
        closeMacroWindow()
    elseif $ShowMacros == 1 then
        createMacroWindow()
    EndIf
    macrosVisible($ShowMacros)
    regStoreMacroVisibility()
EndFunc   ;==>toggleMacrosView


Func toggleMacrosWindow()
    If BitAND(GUICtrlRead($checkMcrFloat), $GUI_CHECKED) == $GUI_CHECKED Then
        $MacrosFloat = 1
    else
        $MacrosFloat = 0
    endif
    if ($ShowMacros == 1) Then
        if $MacrosFloat == 0 then closeMacroWindow()
        createMacroWindow()
        macrosVisible($ShowMacros)
    EndIf
    regStoreMacroVisibility()
EndFunc   ;==>toggleMacrosView


Func parseMacroEntry($_mcrNo)
    $macroString[$_mcrNo] = GUICtrlRead($iMcr[$_mcrNo-$BankFirst])
    regStoreMacro($_mcrNo)
EndFunc   ;==>parseMacroEntry

Func parseRXhead()
    Local $str = GUICtrlRead($inHead)
    generateHeadTail($str, 0)
    regStoreRXheadTail($str, 0)
    $RXheadStr = $str
EndFunc   ;==>parseRXhead

Func parseRXtail()
    Local $str = GUICtrlRead($InTail)
    generateHeadTail($str, 1)
    regStoreRXheadTail($str, 1)
    $RXtailStr = $str
EndFunc   ;==>parseRXtail


Func generateHeadTail($_input, $_isTail, $DEBUG = 0)
    Local $str, $a, $s1, $splitDone = 0
    $str = StringReplace($_input, "\n", @LF)
    $str = StringReplace($str, "\r", @CR)
    If $DEBUG Then ConsoleWrite("---" & @CRLF & "Data:" & $str & @CRLF & "---" & @CRLF)
    If Not $_isTail Then
        $RXheadEcho = 0
        $RXhead[0] = ""
        $RXhead[1] = ""
        $a = StringSplit($str, "%ECHO", 1)
        $RXhead[0] = $a[1]
        For $i = 2 To $a[0]
            If $splitDone Then ;ok
                $RXhead[1] = $RXhead[1] & "%ECHO" & $a[$i]
            ElseIf StringLen($a[$i - 1]) And StringRight($a[$i - 1], 1) == "%" Then
                $RXhead[0] = $RXhead[0] & "%ECHO" & $a[$i]
            Else
                ; catch the special cases where '%ECHO' is intended content of the string
                $RXhead[1] = $a[$i]
                $RXheadEcho = 1
                $splitDone = 1
                If $DEBUG Then ConsoleWrite("Echo found. ")
            EndIf
        Next
        If $DEBUG Then ConsoleWrite(StringFormat("Head -> '%s'|'%s', echo: %d\n", $RXhead[0], $RXhead[1], $RXheadEcho))
    Else
        $RXtail = $str
        If $DEBUG Then ConsoleWrite(StringFormat("Tail -> '%s'\n", $RXtail))
    EndIf
EndFunc   ;==>generateHeadTail




#cs
    Func parseMacroEntry($_no)
    Local $i, $isQuote = "", $isOP = "", $str =
    Local $s = StringSplit(GUICtrlRead($iMcr[$_no]),"")
    $isQuote = ""
    For $i = 1 To $s[0]
    If $s[$i] = "'" Or $s[$i] = '"' Then
    if $isQuote = "" Then
    $isQuote = ($s[$i])
    ElseIf $s[$i] = $isQuote Then
    ; end of expression
    $isQuote = ""
    EndIf
    ElseIf $isQuote <> "" Then
    $str &= $s[$i] ; add char to string
    ContinueLoop
    Else
    If $s[$i] = "%" Then
    If $opFound Then		; ,( | (( | *(
    ; new level
    ElseIf StringLen($str) And findFunction($str,1) <> -1 Then ; func(
    ; this is function
    $isFunc = 1
    Else
    $op = "*"
    $opFound = 1
    EndIf
    $defLevels += 1
    $levList[$defLevels][$levPARENT] = $cLevel
    $levList[$defLevels][$levCNT] = 0
    $levList[$defLevels][$levENTER] = $defExp
    $levList[$defLevels][$levFUNC] = $isFunc
    $isFunc = 0
    ;~ 		conWrite("New lev " &$defLevels& " : [" & $levList[$defLevels][0] & ","& $levList[$defLevels][1]& "," &$levList[$defLevels][2]& "]" &@CRLF)
    $cLevel = $defLevels
    $expEnd = 1
    $opFound = 1
    ContinueLoop
    ElseIf $s[$i] = "," Or $s[$i] = ";" Then
    If $opFound = 1 Then
    If($op = "," Or $op = ";" Or $op = "++" Or $op = "--" Or ($cLevel > 0 And $levList[$cLevel][$levCNT] > 0)) Then ; assume empty function argument
    $levList[$cLevel][$levCNT] += 1 ; increment exp count on this level
    $expList[$defExp][$expSTR] = $str
    $expList[$defExp][$expOP] = $op ; store operand
    newExpEntry($expList, $defExp, $cLevel) ; add new empty expression to list
    $str = ""
    $op = $s[$i] ; add operand
    Else
    conWrite ("@" &$inLineNo &" (parseVarFuncs) Operand preceeding separator in a row @" & $i & ", check syntax." & @CRLF)
    SetError(1)
    Return
    EndIf
    Else
    $op = $s[$i] ; add operand
    $opFound = 1
    EndIf
    ContinueLoop
    ElseIf $s[$i] = " " Then
    $op = " "
    $expEnd = 1
    ContinueLoop
    EndIf
    EndIf
    If $opFound Or $expEnd Then
    If $expEnd And $levList[$defLevels][$levEXIT] = $defExp And $op = "" Then $op = "*" ; if closing parenthesis and no op, add * as op
    $expEnd = 0
    $opFound = 0
    ; add new expression
    $levList[$cLevel][$levCNT] += 1 ; increment exp count on this level
    $expList[$defExp][$expSTR] = $str
    $expList[$defExp][$expOP] = $op ; store operand
    newExpEntry($expList, $defExp, $cLevel)
    $str = ""
    $op = ""
    ;conWrite("New exp " &$defExp& " @ " & $i & @CRLF)
    EndIf
    $str &= $s[$i] ; add char to string
    Next

    EndFunc
#ce

Func macroEventSend()
    Local $DEBUG = 1
;~     ConsoleWrite(stringformat("tuki, ctrlID = %d\n", @GUI_CtrlId))
    For $i = 0 To $MACRO_NUMBER - 1
        If @GUI_CtrlId = $bMcrSend[$i] Then ;Or @GUI_CTRLID = $iMcr[$i] Then
            local $keyM = 0
            if _IsPressed("10") then
                local $j = Dec("70")
                $keyM = 1
                for $j = Dec("70") to Dec("7B")
                    if _IsPressed(Hex($j)) then
                        $keyM = 0
                        ExitLoop
                    EndIf
                Next
            EndIf
            if $keyM Then ; if Shift is pressed copy macro to input box
                GUICtrlSetData ($InputTX, $macroString[$i + $BankFirst])
                Return 0
            else
                If $DEBUG Then ConsoleWrite("Send macro " & $i + $BankFirst & ",data: " & $macroString[$i + $BankFirst] & @CRLF)
                Return macroSend($i + $BankFirst)
            EndIf
        EndIf
    Next
EndFunc   ;==>macroEventSend

Func macroSend($_no)
    Local $mp1 = "", $mp2 = ""
    ; parse string
    ;    Return sendData($str)
    Return parseString($macroString[$_no])
EndFunc   ;==>macroSend

Func macroRepeatSend()
    Local $i
	For $i = 0 To $MACRO_PER_BANK - 1
		If @GUI_CtrlId = $checkMcrRsend[$i] Then
			ExitLoop
		EndIf
	Next
	; if new macro is enable while previous was selected, radio button the check boxes and exit
	If $mcrRepeat = True And $i + $BankFirst <> $mcrRptCur Then
		$macroRpt[$mcrRptCur] = 0
		if $mcrRptCur >= $BankFirst and $mcrRptCur < $BankFirst + $MACRO_PER_BANK Then
			GUICtrlSetState($checkMcrRsend[$mcrRptCur - $BankFirst], $GUI_UNCHECKED)
		EndIf
		ConsoleWrite(StringFormat("(macroRepeatSend) Change of macro repeat, from #%d to #%d\r\n", $mcrRptCur, $i+$BankFirst))
		$mcrRptCur = $i + $BankFirst
		$macroRpt[$mcrRptCur] = 1
		Return
	EndIf
	If $mcrRepeat = True Then
		$mcrRepeat = False
		$macroRpt[$mcrRptCur] = 0
		ConsoleWrite(StringFormat("(macroRepeatSend) End macro repeat\r\n"))
	Else
		$mcrRepeat = True
		ConsoleWrite(StringFormat("(macroRepeatSend) New macro repeat, #%d\r\n", $i + $BankFirst))
		$mcrRptCur = $i + $BankFirst
		$macroRpt[$mcrRptCur] = 1
	EndIf
    Return
EndFunc   ;==>macroRepeatSend


Func changeMacroBank ()
    local $bank = 0, $i
    for $i = 0 to $MACRO_BANKS - 1
        if BitAND(GUICtrlRead($radioBank[$i]), $GUI_CHECKED) = $GUI_CHECKED Then ExitLoop
        $bank += 1
    Next
    if $bank < $MACRO_BANKS then
        $curMbank = $bank
        $BankFirst = $curMbank * $MACRO_PER_BANK
        For $i = 0 To $MACRO_PER_BANK - 1
            GUICtrlSetData($iMcr[$i], $macroString[$i + $BankFirst]) ; macro input
            GUICtrlSetData($bMcrSend[$i], "M" & ($i + $BankFirst + 1)) ; button name
            GUICtrlSetData($iMcrRT[$i], $macroRptTime[$i + $BankFirst]) ; repeat time
            if $macroRptTime[$i + $BankFirst] Then ; repeat checkbox
                GUICtrlSetData($checkMcrRsend[$i], $GUI_CHECKED)
            Else
                GUICtrlSetData($checkMcrRsend[$i], $GUI_UNCHECKED)
            EndIf
        Next
    EndIf
    regStoreMacroBank()
EndFunc




Func extractPath($_str)
    Local $a = StringSplit($_str, "\")
    Local $path
    If $a[0] = 1 Then
        SetError(1)
        Return ""
    EndIf
    $path = StringTrimRight($_str, StringLen($a[$a[0]]) + 1)
    Return $path
EndFunc   ;==>extractPath

