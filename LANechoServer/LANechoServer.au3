
#cs
Resources:
    Internet Assigned Number Authority - all Content-Types: http://www.iana.org/assignments/media-types/
    World Wide Web Consortium - An overview of the HTTP protocol: http://www.w3.org/Protocols/

Credits:
    Manadar for starting on the webserver.
    Alek for adding POST and some fixes
    Creator for providing the "application/octet-stream" MIME type.
#ce

; // OPTIONS HERE //
Local $sRootDir = @ScriptDir & "\www" ; The absolute path to the root directory of the server.
Local $sIP = @IPAddress1 ; ip address as defined by AutoIt
Local $iPort = 1234 ; the listening port
Local $iMaxUsers = 3 ; Maximum number of users who can simultaneously get/post
Global $INIT_FILE = @WorkingDir & "\COMsrv.ini"
Global $COMhead[3], $COMtail, $COMbaud, $COMport, $COMopen = 0, $hCOM
Local $cmd,$resp
; // END OF OPTIONS //

Local $aSocket[$iMaxUsers] ; Creates an array to store all the possible users
Local $sBuffer[$iMaxUsers] ; All these users have buffers when sending/receiving, so we need a place to store those
Local $sIsWin[$iMaxUsers] ; All these users have buffers when sending/receiving, so we need a place to store those
Global $socketCount = 0
For $x = 0 to UBound($aSocket)-1 ; Fills the entire socket array with -1 integers, so that the server knows they are empty.
    $aSocket[$x] = -1
Next


#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=C:\Users\Darko\Desktop\AS3923_Test_Suite\_source\COMsrv.kxf
$Form1_1 = GUICreate("COM Server", 263, 151, 346, 262)
$GroupBox1 = GUICtrlCreateGroup("", 8, 1, 249, 97)
GUICtrlCreateLabel("LAN echo server", 16, 16, 226, 28, $SS_CENTER)
GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
GUICtrlCreateLabel("Clients connected", 56, 66, 89, 17)
$InputClients = GUICtrlCreateInput("", 160, 62, 49, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_READONLY))
GUICtrlCreateGroup("", -99, -99, 1, 1)
$ButtonClose = GUICtrlCreateButton("Close server", 161, 115, 75, 25)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

TCPStartup() ; AutoIt needs to initialize the TCP functions

$iMainSocket = TCPListen($sIP,$iPort) ;create main listening socket
If @error Then ; if you fail creating a socket, exit the application
    MsgBox(0x20, "COM port server", "Unable to create a socket on port " & $iPort & ".") ; notifies the user that the HTTP server will not run
    Exit ; if your server is part of a GUI that has nothing to do with the server, you'll need to remove the Exit keyword and notify the user that the HTTP server will not work.
EndIf


ConsoleWrite( "Server created." & @CRLF) ; If you're in SciTE,
GUIupdateClients()
While 1
    $nMsg = GUIGetMsg()
    Switch $nMsg
	    Case $GUI_EVENT_CLOSE
		Terminate()
	    Case $ButtonClose
		Terminate()

    EndSwitch
    $iNewSocket = TCPAccept($iMainSocket) ; Tries to accept incoming connections

    If $iNewSocket >= 0 Then ; Verifies that there actually is an incoming connection
        For $x = 0 to $iMaxUsers-1 ; Attempts to store the incoming connection
            If $aSocket[$x] = -1 Then
                $aSocket[$x] = $iNewSocket ;store the new socket
		$sBuffer[$x] = ""
		$sIsWin[$x] = "dunno"
		$socketCount += 1
		GUIupdateClients()
                ExitLoop
            EndIf
        Next
    EndIf

    For $x = 0 to $iMaxUsers-1 ; A big loop to receive data from everyone connected
        If $aSocket[$x] = -1 Then ContinueLoop ; if the socket is empty, it will continue to the next iteration, doing nothing
        $sNewData = TCPRecv($aSocket[$x],1024) ; Receives a whole lot of data if possible
        If @error Then ; Client has disconnected
            $aSocket[$x] = -1 ; Socket is freed so that a new user may join
	    $socketCount -= 1
	    GUIupdateClients()
            ContinueLoop ; Go to the next iteration of the loop, not really needed but looks oh so good
        ElseIf $sNewData Then ; data received
	    TCPSend($aSocket[$x], $sNewData)
        EndIf
    Next

;~     Sleep(10)
WEnd

Func Terminate()
    TCPShutdown()
    Exit
EndFunc



Func GUIupdateClients()
    GUICtrlSetData($InputClients, $socketCount)
EndFunc
