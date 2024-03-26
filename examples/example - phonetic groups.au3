#include "..\FuzzyString.au3"

; example: put words with same phonetic code into groups

$aAutoItFuncs = StringRegExp(FileRead(StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 1, -1)) & 'SciTe\api\au3.api'), '(?m)^_?\K\w+(?=\h*\()', 3)
$mGroups = _FS_groupByPhonetic($aAutoItFuncs)

For $sCode In MapKeys($mGroups)
	$aWords = $mGroups[$sCode]

	; only groups with > 1 members:
	if UBound($aWords) < 2 Then ContinueLoop

	ConsoleWrite($sCode & ":" & @CRLF)
	For $i = 0 To UBound($aWords) - 1
		ConsoleWrite(@TAB & $aWords[$i] & @CRLF)
	Next
Next
