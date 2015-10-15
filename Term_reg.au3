
; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.


; Registry manipulation functions


Func logSelectFile()
    Local $path = FileSaveDialog("Select log file", $logDirectory, "Log file (*.log)", 18, "*.log")
    If @error Or $path = "" Then Return -1
    $hLog = FileOpen($path, 10)
    If @error Then Return -2
    $logFileSelected = 1
    Return 0
EndFunc   ;==>logSelectFile


Func logStartLog()
    If $logFileSelected = 0 Then
        If logSelectFile() Then
            Return
        EndIf
    EndIf
    $logEnabled = 1
EndFunc   ;==>logStartLog


Func logStop()
    $logEnabled = 0
EndFunc   ;==>logStop

Func regLoadDirectories()
    Local $temp
    $temp = RegRead($REG_ROOT, "MacroDirectory")
    If Not @error Then $macroDirectory = $temp
    $temp = RegRead($REG_ROOT, "LogDirectory")
    If Not @error Then $logDirectory = $temp
EndFunc   ;==>regLoadDirectories


; Root:"HKEY_CURRENT_USER\Software\SRLabs\Terminal"
; Connection: key .\Connection
;   - value LastComType (REG_SZ):  COM|LAN|VISA
;
; COM port settings: key .\Connection\COM
;   - value LastPort : name of last port used
;
;

Func regLoadGUI()
    $WinPosX = RegRead($REG_ROOT, "WindowPositionX")
    If @error or $WinPosX < -$WindowWidth + 40 or $WinPosX >= @DesktopWidth - 40 Then
        $WinPosX = (@DesktopWidth - $WindowWidth) / 2
    EndIf
    $WinPosY = RegRead($REG_ROOT, "WindowPositionY")
    If @error or $WinPosY < -20 or $WinPosY >= @DesktopHeight - 20 Then
        $WinPosY = @DesktopHeight / 10
    EndIf
    ; load newline character settings from registry
    Local $reg = RegRead($REG_ROOT, "NewLine")
    If $reg = "" Or $reg = "CRLF" Then
        GUICtrlSetState($checkTX_CR, $GUI_CHECKED)
        GUICtrlSetState($checkCRLF, $GUI_CHECKED)
        If $reg = "" Then regStoreGUI()
    ElseIf $reg = "CR" Then
        GUICtrlSetState($checkTX_CR, $GUI_CHECKED)
        GUICtrlSetState($checkCRLF, $GUI_UNCHECKED)
    ElseIf $reg = "LF" Then
        GUICtrlSetState($checkTX_CR, $GUI_UNCHECKED)
        GUICtrlSetState($checkCRLF, $GUI_CHECKED)
    Else
        GUICtrlSetState($checkTX_CR, $GUI_UNCHECKED)
        GUICtrlSetState($checkCRLF, $GUI_UNCHECKED)
    EndIf
EndFunc   ;==>regLoadGUI

Func regStoreGUI()
    ; store newline characters to registry
    Local $reg = ""
    If BitAND(GUICtrlRead($checkTX_CR), $GUI_CHECKED) = $GUI_CHECKED Then
        $reg = "CR"
    EndIf
    If BitAND(GUICtrlRead($checkCRLF), $GUI_CHECKED) = $GUI_CHECKED Then
        $reg &= "LF"
    EndIf
    RegWrite($REG_ROOT, "NewLine", "REG_SZ", $reg)
EndFunc   ;==>regStoreGUI

Func regLoadLastCOMport() ; TODO
    Local $aList, $prev
    $prev = RegRead($REG_ROOT & "\Connection", "LastCOMPort")
    If $prev = "" Then Return
    $aList = StringSplit(_GUICtrlComboBox_GetList($cCOM), "|")
    For $i = 1 To $aList[0]
        If $aList[$i] = $prev Then
            GUICtrlSetData($cCOM, $aList[$i])
            regLoadCOMport(1)
            Return
        EndIf
    Next
EndFunc   ;==>regLoadLastCOMport

Func regLoadCommType()
    Local $temp
    $temp = RegRead($REG_ROOT, "LastCommType")
    If @error Then
;~ 	registryError(@error)
        $temp = "COM"
    EndIf
;~     ConsoleWrite("Last comm. type was " & $temp & @CRLF)
    If $temp = "VISA" And $VISA32_AVAILABLE Then
        $temp = $useVISA
    ElseIf $temp = "LAN" Then
        $temp = $useLAN
    Else
        $temp = $useCOM
    EndIf
    showDialog($temp, 1)
EndFunc   ;==>regLoadCommType

Func regStoreCommType($_type)
    Local $type = "COM", $DEBUG = 1
    If $_type = $useLAN Then $type = "LAN"
    If $_type = $useVISA Then $type = "VISA"
    If $DEBUG Then ConsoleWrite("Storing last comm. type as " & $type & @CRLF)
    RegWrite($REG_ROOT, "LastCommType", "REG_SZ", $type)
EndFunc   ;==>regStoreCommType



Func regLoadRXheadTail()
    Local $temp
    $temp = RegRead($REG_ROOT & "\Connection", "LastRXhead")
    GUICtrlSetData($inHead, $temp)
    generateHeadTail($temp, 0)
    $RXheadStr = $temp
    $temp = RegRead($REG_ROOT & "\Connection", "LastRXtail")
    GUICtrlSetData($InTail, $temp)
    generateHeadTail($temp, 1)
    $RXtailStr = $temp
EndFunc   ;==>regLoadRXheadTail


Func regStoreRXheadTail($_str, $_isTail)
    If $_isTail Then
        RegWrite($REG_ROOT & "\Connection", "LastRXtail", "REG_SZ", $_str)
    Else
        RegWrite($REG_ROOT & "\Connection", "LastRXhead", "REG_SZ", $_str)
    EndIf
EndFunc   ;==>regStoreRXheadTail

Func regLoadLAN()
    Local $temp
    $temp = RegRead($REG_ROOT & "\Connection", "LastLANaddr")
    GUICtrlSetData($iIP, $temp)
    $temp = RegRead($REG_ROOT & "\Connection", "LastLANport")
    GUICtrlSetData($iPort, $temp)
EndFunc   ;==>regLoadLAN


Func regStoreLAN()
    RegWrite($REG_ROOT & "\Connection", "LastLANaddr", "REG_SZ", GUICtrlRead($iIP))
    RegWrite($REG_ROOT & "\Connection", "LastLANport", "REG_SZ", GUICtrlRead($iPort))
EndFunc   ;==>regStoreLAN


Func regLoadVISA()
    Local $temp
    $temp = RegRead($REG_ROOT & "\Connection", "LastVISAaddr")
    GUICtrlSetData($iVISAaddr, $temp)
EndFunc   ;==>regLoadVISA


Func regStoreVISA()
    RegWrite($REG_ROOT & "\Connection", "LastVISAaddr", "REG_SZ", GUICtrlRead($iVISAaddr))
EndFunc   ;==>regStoreVISA


Func regStoreCOMport()
    Local $port = GUICtrlRead($cCOM)
    Local $baud = GUICtrlRead($cBaud)
    Local $dataLen = GUICtrlRead($cData)
    Local $par = GUICtrlRead($cParity)
    Local $sb = GUICtrlRead($cStopbits)
    Local $handshake = GUICtrlRead($cHandShake)
    RegWrite($REG_ROOT & "\Connection", "LastCOMport", "REG_SZ", $port)
    RegWrite($REG_ROOT & "\Connection\" & $port, "Baudrate", "REG_SZ", $baud)
    RegWrite($REG_ROOT & "\Connection\" & $port, "DataFormat", "REG_SZ", $dataLen)
    RegWrite($REG_ROOT & "\Connection\" & $port, "Parity", "REG_SZ", $par)
    RegWrite($REG_ROOT & "\Connection\" & $port, "StopBits", "REG_SZ", $sb)
    RegWrite($REG_ROOT & "\Connection\" & $port, "Handshake", "REG_SZ", $handshake)
EndFunc   ;==>regStoreCOMport


Func regLoadCOMport($_init = 0)
    Local $port = GUICtrlRead($cCOM)
    Local $temp
    $temp = RegRead($REG_ROOT & "\Connection\" & $port, "Baudrate")
    If @error Then $temp = 112500
    If Not @error Or $_init Then GUICtrlSetData($cBaud, $temp)
    $temp = RegRead($REG_ROOT & "\Connection\" & $port, "DataFormat")
    If @error Then $temp = 8
    If Not @error Or $_init Then GUICtrlSetData($cData, $temp)
    $temp = RegRead($REG_ROOT & "\Connection\" & $port, "Parity")
    If @error Then $temp = "none"
    If Not @error Or $_init Then GUICtrlSetData($cParity, $temp)
    $temp = RegRead($REG_ROOT & "\Connection\" & $port, "StopBits")
    If @error Then $temp = 1
    If Not @error Or $_init Then GUICtrlSetData($cStopbits, $temp)
    $temp = RegRead($REG_ROOT & "\Connection\" & $port, "Handshake")
    If @error Then $temp = "none"
    If Not @error Or $_init Then GUICtrlSetData($cHandShake, $temp)
    Return
EndFunc   ;==>regLoadCOMport

Func regStoreMacro($_no)
    Local $err = 0
    $_no += 1
    If StringLen($_no) = 1 Then $_no = "0" & $_no
    RegWrite($REG_ROOT & "\Macros", "M" & $_no & "_string", "REG_SZ", $macroString[$_no - 1])
    If @error Then
        registryError(@error)
        $err = 1
    EndIf
EndFunc   ;==>regStoreMacro

Func regReadMacros()
    Local $err = 0, $no, $i, $j
    $ShowMacros = RegRead($REG_ROOT, "ShowMacros")
    If @error Then
        $ShowMacros = 0
    EndIf
    $MacrosFloat = RegRead($REG_ROOT, "FloatMacros")
    If @error Then
        $MacrosFloat = 0
    EndIf
    if $MacrosFloat <> 0 Then
        GUICtrlSetState($checkMcrFloat, $GUI_CHECKED)
    else
        GUICtrlSetState($checkMcrFloat, $GUI_UNCHECKED)
    EndIf

    $McrWinPosX = RegRead($REG_ROOT, "MacroWindowPositionX")
    If @error or $McrWinPosX < -$MCR_GRP_WIDTH + 40 or $McrWinPosX >= @DesktopWidth - 40 Then
        $McrWinPosX = (@DesktopWidth - $MCR_GRP_WIDTH) / 2
    EndIf
    $McrWinPosY = RegRead($REG_ROOT, "MacroWindowPositionY")
    If @error or $McrWinPosY < -20 or $McrWinPosY >= @DesktopHeight - 20 Then
        $McrWinPosY = @DesktopHeight / 10
    EndIf
    $curMbank = RegRead($REG_ROOT, "CurrentMacroBank")
    If @error Or ($curMbank >= $MACRO_BANKS or $curMbank < 0) Then
        $curMbank = 0
    EndIf
    For $i = 0 To $MACRO_NUMBER - 1
        $no = $i + 1
        If StringLen($no) = 1 Then $no = "0" & $no
        $macroString[$i] = RegRead($REG_ROOT & "\Macros", "M" & $no & "_string")
        if $i >= $BankFirst and $i < $BankFirst + $MACRO_PER_BANK then
            GUICtrlSetData($iMcr[$i-$BankFirst], $macroString[$i])
        EndIf
	$macroRptTime[$i] = 1000
	$macroRpt[$i] = 0
    Next
EndFunc   ;==>regReadMacros

Func regStoreMacroVisibility()
    RegWrite($REG_ROOT, "ShowMacros", "REG_SZ", $ShowMacros)
    RegWrite($REG_ROOT, "FloatMacros", "REG_SZ", $MacrosFloat)
EndFunc   ;==>regStoreMacroVisibility

Func regStoreMacroBank()
    RegWrite($REG_ROOT, "CurrentMacroBank", "REG_SZ", $curMbank)
EndFunc

Func regStoreMacroWinPos()
    RegWrite($REG_ROOT, "MacroWindowPositionX", "REG_SZ", $McrWinPosX)
    RegWrite($REG_ROOT, "MacroWindowPositionY", "REG_SZ", $McrWinPosY)
EndFunc

Func regStoreWinPos()
    RegWrite($REG_ROOT, "WindowPositionX", "REG_SZ", $WinPosX)
    RegWrite($REG_ROOT, "WindowPositionY", "REG_SZ", $WinPosY)
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
    EndSwitch
EndFunc   ;==>registryError

Func wipeRegistry()
    RegDelete($REG_ROOT)
EndFunc   ;==>wipeRegistry
