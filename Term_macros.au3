
; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.


Func readMacroFile()
    Local $path
    Local $hFile, $expectMacro = 1, $par = 0, $skip = 0, $no, $a
    HotKeySet("{ENTER}")
    $path = FileOpenDialog ( "Open macro file", $macroDirectory, "Macro file (*.tmf)|Text files (*.txt)|All files (*.*)",3,"*.tmf", $Terminal)
    HotKeySet("{ENTER}", "captureENTER")
    If @error Or $path = "" Then Return -1
    $hFile = FileOpen ($path, 0)
    If @error Then Return -2
    while 1
	$line = FileReadLine ( $hFile)
	If @error Then exitloop
	If $expectMacro Then
	    If StringinStr($line, "M") = 1 Then
		$line = stringTrimLeft($line, 1)
		$a = StringSplit($line,"P")
		If $a[0] > 2 Then ContinueLoop
		If $a[0] = 2 Then
		    If $a[2] = 1 Or $a[2] = 2 Then
			$par = $a[2]
		    Else
			ContinueLoop
		    EndIf
		EndIf
		If String(Number($a[1])) = $a[1] And ($a[1] > 0) And ($a[1] <= $MACRO_NUMBER) Then
		    $no = $a[1]
		    $expectMacro = 0
		Else
		    $par = 0
		    ContinueLoop
		EndIf
	    EndIf
	Else
	    If Not $skip Then
		If $par Then
		    GUICtrlSetData($iMcrP[$par-1][$no-1], $line)
		Else
		    $macroString[$no-1] = $line
		    GUICtrlSetData($iMcr[$no-1],$line)
		    parseMacroEntry($no-1)
		EndIf
	    EndIf
	    $skip = 0
	    $par = 0
	    $expectMacro = 1
	EndIf
    Wend
    FileClose($hFile)
    Return 0
EndFunc

Func writeMacroFile()
    Local $writePars = 1
    Local $path, $hFile, $i
    HotKeySet("{ENTER}")
    $path = FileSaveDialog ( "Save macro file", $macroDirectory, "Macro file (*.tmf)",16,"*.tmf", $Terminal)
    HotKeySet("{ENTER}", "captureENTER")
    If @error Or $path = "" Then Return -1
    If StringRight ($path,4) <> ".tmf" then $path = $path & ".tmf"
    $hFile = FileOpen ($path, 10)
    If @error Then Return -2
    $path = extractPath($path)
    If Not @error Then
	RegWrite ( $REG_ROOT, "MacroDirectory", "REG_SZ", $path)
	$macroDirectory = $path
    EndIf
    FileWriteLine($hFile,"# SRLabs terminal macro file")
    For $i = 1 To $MACRO_NUMBER
	FileWriteLine($hFile,"M" & $i)
	FileWriteLine($hFile, GUICtrlRead($iMcr[$i-1]))
	If $writePars Then
	    FileWriteLine($hFile,"M" & $i & "P1")
	    FileWriteLine($hFile, GUICtrlRead($iMcrP[0][$i-1]))
	    FileWriteLine($hFile,"M" & $i & "P2")
	    FileWriteLine($hFile, GUICtrlRead($iMcrP[1][$i-1]))
	EndIf
    Next
    FileClose($hFile)
    Return 0
EndFunc


Func toggleMacrosView ()
    $ShowMacros = 1 - $ShowMacros
    macrosVisible($ShowMacros)
    regStoreMacroVisibility()
EndFunc


Func parseMacroEntry($_no)
    $macroString[$_no] = GUICtrlRead($iMcr[$_no])
    $macroStrCat[$_no][0] = $macroString[$_no]
    $macroStrCat[$_no][1] = ""
    $macroStrCat[$_no][2] = ""
    $macroStrPar[$_no][0] = ""
    $macroStrPar[$_no][1] = ""
    regStoreMacro($_no)
EndFunc

Func parseRXhead()
    Local $str = GUICtrlRead($inHead)
    generateHeadTail($str, 1)
    regStoreRXheadTail($str, 0)
    $RXheadStr = $str
EndFunc

Func parseRXtail()
    Local $str = GUICtrlRead($InTail)
    generateHeadTail($str, 0)
    regStoreRXheadTail($str, 1)
    $RXtailStr = $str
EndFunc


Func generateHeadTail ($_input, $_isHead, $DEBUG = 0)
    Local $str, $a, $s1, $splitDone = 0
    $str = StringReplace($_input, "\n", @LF)
    $str = StringReplace($str, "\r", @CR)
    $RXheadEcho = 0
    $RXhead[0] = ""
    $RXhead[1] = ""
    If $DEBUG Then ConsoleWrite("---" &@CRLF&"Data:" & $str & @CRLF &"---" & @CRLF)
    if $_isHead Then
	$a = StringSplit($str, "%ECHO")
	For $i = 1 To $a[0] - 1
	    If StringLen($a[$i]) And StringRight ($a[$i],1) <> "%" Then
		$RXhead[0] = $a[$i]
		$RXheadEcho = 1
		$splitDone = 1
		If $DEBUG Then ConsoleWrite("Echo found @ " & StringInStr($str,"%ECHO",0,$i) & @CRLF)
	    ElseIf StringRight ($a[$i],1) == "%" Then
		$a[$i+1] = $a[$i] & "%ECHO" & $a[$i+1]
	    ElseIf $splitDone Then
		$RXhead[1] &= "%ECHO" & $a[$i]
	    EndIf
	Next
	If Not $splitDone Then
	    $RXhead[0] = $a[$a[0]]
	EndIf
    Else
	$RXtail = $str
    EndIf
EndFunc




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
    For $i = 0 To $MACRO_NUMBER - 1
	If @GUI_CTRLID = $bMcrSend[$i] Then ;Or @GUI_CTRLID = $iMcr[$i] Then
	    If $DEBUG Then ConsoleWrite("Send macro " & $i & ",data: " &$macroStrCat[$i][0]&@CRLF)
	    Return macroSend($i)
	EndIf
    Next
EndFunc

Func macroSend($_no)
    Local $mp1 = "", $mp2 = ""
;~     If $ConOpen = $connectNone Then Return
    $str = $macroStrCat[$_no][0]
    If $macroParVisible[$_no][0] Then
	$mp1 &= $macroStrPar[$_no][0]
    EndIf
    If $macroParVisible[$_no][1] Then
	$mp2 &= $macroStrPar[$_no][1]
    EndIf
    ; parse string
;    Return sendData($str)
    Return parseString($macroString[$_no], $mp1, $mp2)
EndFunc

Func macroRepeatSend ()
    Local $i
    For $i = 0 To $MACRO_NUMBER - 1
	If @GUI_CTRLID = $checkMcrRsend[$i] Then
	    ExitLoop
	EndIf
    Next
    ; if new macro is enable while previous was selected, radio button the check boxes and exit
    If $mcrRepeat = true and $i <> $mcrRptCur Then
	GUICtrlSetState($checkMcrRsend[$mcrRptCur],$GUI_UNCHECKED)
;	GUICtrlSetState($checkMcrRsend[$i],$GUI_CHECKED)
	ConsoleWrite (StringFormat("(macroRepeatSend) Change of macro repeat, from #%d to #%d\r\n", $mcrRptCur, $i))
	$mcrRptCur = $i
	Return
    Endif
    If $mcrRepeat = true Then
	$mcrRepeat = False
	ConsoleWrite (StringFormat("(macroRepeatSend) End macro repeat\r\n"))
	Return
    Else
	$mcrRepeat = true
	ConsoleWrite (StringFormat("(macroRepeatSend) New macro repeat, #%d\r\n", $i))
	$mcrRptCur = $i
    EndIf
    Return
EndFunc


Func macrosVisible($_on)
    Local $i, $task = $GUI_HIDE
    Local $wPos = WinGetPos("SRLabs Terminal")
    If $_on Then $task = $GUI_SHOW
    If $_on = 1 Then
	WinMove("SRLabs Terminal","",$wPos[0],$wPos[1],$GUI_width+$MACRO_WIN_WIDTH+4,$GUI_height)
	GUICtrlSetData($bMacroWindow,"Hide Macros <-")
    Else
	WinMove("SRLabs Terminal","",$wPos[0],$wPos[1],$GUI_width,$GUI_height)
	GUICtrlSetData($bMacroWindow,"Show Macros ->")
    EndIf
    For $i = 0 To $MACRO_NUMBER - 1
	adjustMacroInput($_on,$i)
	GUICtrlSetState($bMcrSend[$i],$task)
	GUICtrlSetState($iMcrRT[$i],$task)
	GUICtrlSetState($checkMcrRsend[$i],$task)
    Next
    GUICtrlSetState($gMacros,$task)
    Return
EndFunc


Func adjustMacroInput ($_show,$_no)
    Local $cnt = 0, $width
    If not $_show Then
	GUICtrlSetState($iMcrP[0][$_no],$GUI_HIDE)
	GUICtrlSetState($iMcrP[1][$_no],$GUI_HIDE)
	GUICtrlSetState($iMcr[$_no],$GUI_HIDE)
	Return
    EndIf
    If $macroParVisible[$_no][1] Then
	$cnt = 1
	GUICtrlSetState($iMcrP[1][$_no],$GUI_SHOW)
    EndIf
    If $macroParVisible[$_no][0] Then
	$cnt += 1
	If $cnt = 2 Then
	    GUICtrlSetPos($iMcrP[0][$_no],$MCR_GRP_LEFT+176)
	Else
	    GUICtrlSetPos($iMcrP[0][$_no],$MCR_GRP_LEFT+212)
	EndIf
	GUICtrlSetState($iMcrP[0][$_no],$GUI_SHOW)
    EndIf
    GUICtrlSetPos($iMcr[$_no], Default, Default, $MACRO_INPUT_W-($cnt*$MACRO_INPUT_DIFF))
    GUICtrlSetState($iMcr[$_no],$GUI_SHOW)
EndFunc

Func extractPath($_str)
    Local $a = StringSplit($_str, "\")
    Local $path
    if $a[0] = 1 Then
	SetError(1)
	Return ""
    EndIf
    $path = StringTrimRight($_str,StringLen($a[$a[0]])+1)
    Return $path
EndFunc

