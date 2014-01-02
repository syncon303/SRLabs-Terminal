
; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.



Func logSelectFile()
    Local $path = FileSaveDialog ( "Select log file", $logDirectory, "Log file (*.log)",18,"*.log")
    If @error Or $path = "" Then Return -1
    $hLog = FileOpen ($path, 10)
    If @error Then Return -2
    $logFileSelected = 1
    Return 0
EndFunc


Func logStartLog()
    If $logFileSelected = 0 Then
	If logSelectFile() Then
	    Return
	Endif
    Endif
    $logEnabled = 1
EndFunc


Func logStop()
    $logEnabled = 0
EndFunc

Func regLoadDirectories()
    Local $temp
    $temp = RegRead($REG_ROOT, "MacroDirectory")
    If Not @error Then $macroDirectory = $temp
    $temp = RegRead($REG_ROOT, "LogDirectory")
    If Not @error Then $logDirectory = $temp
EndFunc


; Root:"HKEY_CURRENT_USER\Software\SRLabs\Terminal"
; Connection: key .\Connection
;   - value LastComType (REG_SZ):  COM|LAN|VISA
;
; COM port settings: key .\Connection\COM
;   - value LastPort : name of last port used
;
;

Func loadRegistry ()

EndFunc

Func regLoadLastCOMport() ; TODO
    Local $aList, $prev
    $prev = RegRead($REG_ROOT & "\Connection", "LastCOMPort")
    If $prev = "" Then Return
    $aList = StringSplit(_GUICtrlComboBox_GetList($cCOM), "|")
    For $i = 1 To $aList[0]
	If $aList[$i] = $prev Then
	    GUICtrlSetData($cCOM,$aList[$i])
	    regLoadCOMport(1)
	    Return
	Endif
    Next
EndFunc

Func regLoadCommType()
    Local $temp
    $temp = RegRead($REG_ROOT, "LastCommType")
    If @error Then
;~ 	registryError(@error)
	$temp = "COM"
    Endif
;~     ConsoleWrite("Last comm. type was " & $temp & @CRLF)
    If $temp = "VISA" And $VISA32_AVAILABLE Then
	$temp = $useVISA
    elseif $temp = "LAN" Then
	$temp = $useLAN
    else
	$temp = $useCOM
    Endif
    showDialog($temp, 1)
EndFunc

Func regStoreCommType($_type)
    Local $type = "COM", $DEBUG = 1
    If $_type = $useLAN Then $type = "LAN"
    If $_type = $useVISA Then $type = "VISA"
    If $DEBUG Then ConsoleWrite("Storing last comm. type as " & $type & @CRLF)
    RegWrite ( $REG_ROOT, "LastCommType", "REG_SZ", $type )
EndFunc



Func regLoadRXheadTail()
    Local $temp
    $temp = RegRead($REG_ROOT & "\Connection", "LastRXhead")
    GUICtrlSetData($inHead, $temp)
    $temp = RegRead($REG_ROOT & "\Connection", "LastRXtail")
    GUICtrlSetData($InTail, $temp)
EndFunc


Func regStoreRXheadTail($_str, $_isTail)
    If $_isTail Then
	RegWrite ( $REG_ROOT & "\Connection", "LastRXtail", "REG_SZ", $_str)
    Else
	RegWrite ( $REG_ROOT & "\Connection", "LastRXhead", "REG_SZ", $_str)
    Endif
EndFunc

Func regLoadLAN()
    Local $temp
    $temp = RegRead($REG_ROOT & "\Connection", "LastLANaddr")
    GUICtrlSetData($iIP, $temp)
    $temp = RegRead($REG_ROOT & "\Connection", "LastLANport")
    GUICtrlSetData($iPort, $temp)
EndFunc


Func regStoreLAN()
    RegWrite ( $REG_ROOT & "\Connection", "LastLANaddr", "REG_SZ", GUICtrlRead ($iIP) )
    RegWrite ( $REG_ROOT & "\Connection", "LastLANport", "REG_SZ", GUICtrlRead ($iPort))
EndFunc


Func regLoadVISA()
    Local $temp
    $temp = RegRead($REG_ROOT & "\Connection", "LastVISAaddr")
    GUICtrlSetData($iVISAaddr, $temp)
EndFunc


Func regStoreVISA()
    RegWrite ( $REG_ROOT & "\Connection", "LastVISAaddr", "REG_SZ", GUICtrlRead ($iVISAaddr) )
EndFunc


Func regStoreCOMport()
    Local $port = GUICtrlRead ($cCOM)
    Local $baud = GUICtrlRead ($cBaud)
    Local $dataLen = GUICtrlRead ($cData)
    Local $par = GUICtrlRead ($cParity)
    Local $sb = GUICtrlRead ($cStopbits)
    Local $handshake = GUICtrlRead ($cHandShake)
    RegWrite ( $REG_ROOT & "\Connection", "LastCOMport", "REG_SZ", $port )
    RegWrite ( $REG_ROOT & "\Connection\"& $port, "Baudrate", "REG_SZ", $baud)
    RegWrite ( $REG_ROOT & "\Connection\"& $port, "DataFormat", "REG_SZ", $dataLen)
    RegWrite ( $REG_ROOT & "\Connection\"& $port, "Parity", "REG_SZ", $par)
    RegWrite ( $REG_ROOT & "\Connection\"& $port, "StopBits", "REG_SZ", $sb)
    RegWrite ( $REG_ROOT & "\Connection\"& $port, "Handshake", "REG_SZ", $handshake)
EndFunc


Func regLoadCOMport($_init = 0)
    Local $port = GUICtrlRead ($cCOM)
    Local $temp
    $temp = RegRead($REG_ROOT & "\Connection\" & $port, "Baudrate")
    If @error Then $temp = 112500
    If not @error Or $_init Then GUICtrlSetData($cBaud, $temp)
    $temp = RegRead($REG_ROOT & "\Connection\" & $port, "DataFormat")
    If @error Then $temp = 8
    If not @error Or $_init Then GUICtrlSetData($cData, $temp)
    $temp = RegRead($REG_ROOT & "\Connection\" & $port, "Parity")
    If @error Then $temp = "none"
    If not @error Or $_init Then GUICtrlSetData($cParity, $temp)
    $temp = RegRead($REG_ROOT & "\Connection\" & $port, "StopBits")
    If @error Then $temp = 1
    If not @error Or $_init Then GUICtrlSetData($cStopbits, $temp)
    $temp = RegRead($REG_ROOT & "\Connection\" & $port, "Handshake")
    If @error Then $temp = "none"
    If not @error Or $_init Then GUICtrlSetData($cHandShake, $temp)
    Return
EndFunc

Func regStoreMacro($_no)
    Local $err = 0
    $_no += 1
    If StringLen($_no) = 1 Then $_no = "0" & $_no
    RegWrite ( $REG_ROOT & "\Macros", "M" & $_no&"_string", "REG_SZ", $macroString[$_no-1] )
    If @error Then
	registryError(@error)
	$err = 1
    Endif
    RegWrite ( $REG_ROOT & "\Macros", "M" & $_no&"_parameter1", "REG_SZ", $macroStrPar[$_no-1][0])
    If @error Then
	registryError(@error)
	$err = 1
    Endif
    RegWrite ( $REG_ROOT & "\Macros", "M" & $_no&"_parameter2", "REG_SZ", $macroStrPar[$_no-1][1] )
    If @error Then
	registryError(@error)
	$err = 1
    Endif
EndFunc

Func regReadMacros()
    Local $err = 0, $no
    $ShowMacros = RegRead($REG_ROOT , "ShowMacros")
    If @error Then
	$ShowMacros = 0
    Endif
    For $i = 0 To $MACRO_NUMBER - 1
	$no = $i + 1
	If StringLen($no) = 1 Then $no = "0" & $no
	$macroString[$i] = RegRead($REG_ROOT & "\Macros", "M" & $no & "_string")
;~ 	If @error Then
;~ 	    registryError(@error)
;~ 	    $err = 1
;~ 	Endif
	GUICtrlSetData($iMcr[$i],$macroString[$i])
	If StringLen($macroString[$i]) Then parseMacroEntry($i)
	$macroStrPar[$i][0] = RegRead($REG_ROOT & "\Macros", "M" & $no & "_parameter1")
;~ 	If @error Then
;~ 	    registryError(@error)
;~ 	    $err = 1
;~ 	Endif
	GUICtrlSetData($iMcrP[0][$i],$macroStrPar[$i][0])
	$macroStrPar[$i][1] = RegRead($REG_ROOT & "\Macros", "M" & $no & "_parameter2")
;~ 	If @error Then
;~ 	    registryError(@error)
;~ 	    $err = 1
;~ 	Endif
	GUICtrlSetData($iMcrP[1][$i],$macroStrPar[$i][1])
    Next
EndFunc

Func regStoreMacroVisibility()
        RegWrite ( $REG_ROOT , "ShowMacros", "REG_SZ", $ShowMacros )
EndFunc


Func registryError($_err)
    Switch $_err
	Case 1
	    ConsoleWrite("Error: Could not open requested registry key." & @CRLF)
	Case 2
	    ConsoleWrite("Error: Could not open requested main key." & @CRLF)
	Case 3
	    ConsoleWrite("Error: Unable to remote connect tot registry." & @CRLF)
	Case -1
	    ConsoleWrite("Error: Unable to open requested registry value." & @CRLF)
	Case 1
	    ConsoleWrite("Error: Registry value type not supported." & @CRLF)
    endswitch
EndFunc

Func wipeRegistry()
    RegDelete($REG_ROOT)
EndFunc