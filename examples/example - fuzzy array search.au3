#include "..\FuzzyString.au3"


; create a list of AutoIt functions
$aAutoItFuncs = StringRegExp(FileRead(StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 1, -1)) & 'SciTe\api\au3.api'), '(?m)^\w+(?=\h*\()', 3)


#Region example 1 - levenshtein 1D-array search by maximum distance

$aSimilar = _FS_ArraySearchFuzzy($aAutoItFuncs, "StringIsAlpha", 4)
_ArrayDisplay($aSimilar, "Similar functions", "", 64, "|", "function|distance|similarity [%]")

#EndRegion



#region keyboard distance based search

$aSimilar = _FS_ArraySearchFuzzy($aAutoItFuncs, "StringIsAlpha", 7, __searchWrapper)
_ArrayDisplay($aSimilar, "Similar functions", "", 64, "|", "function|distance|similarity [%]")

; wrapper function to bring _FS_Keyboard_Distance_Strings into the required form
Func __searchWrapper($sA, $sB, $iMax)
	Local Static $mKeyb = _fs_keyboard_getLayout("QWERTY", False) ; case insensitive qwerty-layout

	Local $aRet = _FS_Keyboard_Distance_Strings($mKeyb, $sA, $sB, $iMax)
	Return SetError(@error, @extended, $aRet)

EndFunc

#EndRegion




#Region Hamming 2D-array search by minimal similarity

; 1D-Array --> 2D-Arraay
_ArrayColInsert($aAutoItFuncs, 0)

$aSimilar = _FS_ArraySearchFuzzy($aAutoItFuncs, "StringIsAlpha", 0.6, _FS_Hamming, 1)
_ArrayDisplay($aSimilar, "Similar functions", "", 64, "|", "dummy|function|distance|similarity [%]")

#EndRegion



