
; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.

Func sendInputData()
    Local $tx = GUICtrlRead ($InputTX)
    Local $ret
    checkHistory($tx)
    $ret = sendData($tx)
    GUICtrlSetData ($InputTX, "")
    Return $ret
EndFunc

Func sendData($_str, $_maxRX = 2048, $_first = 1000, $_next = 100, $DEBUG = 1)
    Local $ending = ""
    Local $tx = $_str
    If GUICtrlRead ($checkTX_CR) = $GUI_CHECKED Then
	$ending &= @CR
	If GUICtrlRead ($checkCRLF) = $GUI_CHECKED Then
	    $ending &= @LF
	EndIf
    EndIf
    $tx &= $ending
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
    writeRXdata($rxbuf)
EndFunc

Func writeRXdata($_str)
    If $logEnabled Then FileWrite($hLog, $_str)
    $editRXcount += StringLen($_str)
    If $editRXcount > $cMaxEditLen Then
	$editRXcount = StringLen($_str)
	GUICtrlSetData ($editRX, "")
    EndIf
    _GUICtrlEdit_AppendText($editRX, $_str)
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


