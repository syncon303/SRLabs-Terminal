
; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.



Func checkHistory($_str)
    If StringLen($_str) = 0 Or $_str == $histLastString Or $_str == $history[decHistPointer($histLast)] Then
	Return ; Unchanged string
    EndIf
;~     If $noHist then
;~ 	newHistory($_str)
;~ 	Return
;~     EndIf
    newHistory($_str)
    Return
EndFunc

Func newHistory($_str)
    Local $incFirst = incHistPointer($histFirst)
    Local $incLast = incHistPointer($histLast)
    If $histFirst = incHistPointer($incLast) Then ; detect buffer overflow
	; history full, roll first history forward
	if $histFirst = $histCur Then $histCur = $incFirst
	$histFirst = $incFirst
    EndIf
    $history[$histLast] = $_str
    If $histLast = incHistPointer($histCur) then $histCur = $histLast
    $histLast = $incLast ; increment last history position
;    $histCur = $histLast
;    $history[$histLast] = ""
    Return
EndFunc

Func historyNext()
    Local $incCur = incHistPointer($histCur)
;    If $noHist = true then Return ; last command was sent and no history is selected
    If $incCur = $histLast then
	if $noHist = true then $noHist = False
	Return
    EndIf
    $histCur = $incCur
    GUICtrlSetData ($InputTX, $history[$histCur])
    Return
EndFunc

Func historyPrev()
    Local $decCur = decHistPointer($histCur)
    if $noHist = true Then
	$noHist = False
	GUICtrlSetData ($InputTX, $history[$histCur])
	Return
    EndIf
    If $histCur = $histFirst then Return
;    If $histCur = $histLast Then
;	$history[$histLast] = GUICtrlRead($InputTX)
;    EndIf
    $histCur = $decCur
    GUICtrlSetData ($InputTX, $history[$histCur])
    Return
EndFunc

Func incHistPointer($_p)
    If $_p = UBound($history)-1 Then Return 0
    Return ($_p + 1)
EndFunc

Func decHistPointer($_p)
    If $_p = 0 Then Return (UBound($history) - 1)
    Return ($_p - 1)
EndFunc

