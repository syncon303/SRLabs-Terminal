
; This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
; To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/deed.en_US.


Func checkHistory($_str)
    If $histFirst = $histLast Then
	newHistory($_str)
	Return
    ElseIf StringLen($_str) = 0 Or $_str = $history[$histCur] Then
	Return ; Unchanged string
    Else
	newHistory($_str)
	Return
    EndIf
EndFunc

Func newHistory($_str)
    Local $incFirst = incHistPointer($histFirst)
    Local $incLast = incHistPointer($histLast)
    Local $dblLast = incHistPointer($incLast)
    If $dblLast = $histFirst Then
	; history full
	$histFirst = $incFirst
    EndIf
    $history[$histLast] = $_str
    $histLast = $incLast
    $histCur = $histLast
    $history[$histLast] = ""
    Return
EndFunc

Func historyNext()
    Local $incCur = incHistPointer($histCur)
    If $histCur = $histLast then Return
    $histCur = $incCur
    GUICtrlSetData ($InputTX, $history[$histCur])
    Return
EndFunc

Func historyPrev()
    Local $decCur = decHistPointer($histCur)
    If $histCur = $histFirst then Return
    If $histCur = $histLast Then
	$history[$histLast] = GUICtrlRead($InputTX)
    EndIf
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

