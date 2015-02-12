
; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.



Func checkHistory($_str)
    If StringLen($_str) = 0 Or $_str == $histLastString Or $_str == $history[decHistPointer($histLast)] Then
        Return ; Unchanged string
    EndIf
    newHistory($_str)
    Return
EndFunc   ;==>checkHistory

Func newHistory($_str)
    Local $incLast = incHistPointer($histLast)
    If $histFirst = incHistPointer($incLast) Then ; detect buffer overflow
        ; history full, roll first history forward
        Local $incFirst = incHistPointer($histFirst)
        If $histFirst = $histCur Then $histCur = $incFirst
        $histFirst = $incFirst
    EndIf
    $history[$histLast] = $_str
    If $histLast = incHistPointer($histCur) Then $histCur = $histLast
    $histCur = $histLast
    $histLast = $incLast ; increment last history position
    ;    $history[$histLast] = ""
    Return
EndFunc   ;==>newHistory

Func historyNext()
    Local $incCur = incHistPointer($histCur)
    If $noHist = True Then Return ; last command was sent and no history is selected
    If $incCur = $histLast Then
        If $noHist = False Then
            $noHist = True
            GUICtrlSetData($InputTX, $inputTXscratch)
            Return
        EndIf
    EndIf
    $histCur = $incCur
    GUICtrlSetData($InputTX, $history[$histCur])
    Return
EndFunc   ;==>historyNext

Func historyPrev()
    Local $decCur = decHistPointer($histCur)
    If $noHist = True Then
        $noHist = False
        $inputTXscratch = GUICtrlRead($InputTX)
        GUICtrlSetData($InputTX, $history[$histCur])
        Return
    EndIf
    If $histCur = $histFirst Then Return
    ;    If $histCur = $histLast Then
    ;	$history[$histLast] = GUICtrlRead($InputTX)
    ;    EndIf
    $histCur = $decCur
    GUICtrlSetData($InputTX, $history[$histCur])
    Return
EndFunc   ;==>historyPrev

Func incHistPointer($_p)
    If $_p = UBound($history) - 1 Then Return 0
    Return ($_p + 1)
EndFunc   ;==>incHistPointer

Func decHistPointer($_p)
    If $_p = 0 Then Return (UBound($history) - 1)
    Return ($_p - 1)
EndFunc   ;==>decHistPointer

