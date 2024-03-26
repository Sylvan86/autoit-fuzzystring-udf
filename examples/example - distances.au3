#include "..\FuzzyString.au3"
#include <String.au3>


ConsoleWrite(StringFormat("% 15s % 15s % 13s % 13s % 11s % 11s % 13s % 11s % 11s\n", "String A", "String B", "Hamming", "Levenshtein", "OSA", "Soundex", "SoundexGerman", "Cologne", "QWERTY") & _StringRepeat("-", 121) & @CRLF)

Local $mKeyb = _fs_keyboard_getLayout("QWERTY", False) ; case insensitive qwertz-layout

Global $aPairs[][2] = [["plauge", "plague"], ["123", ""], ["Bar", "Bier"], ["uninformiert", "uniformiert"], ["kitten", "sitting"], ["Test", "Tes"], ["Spass", "pass"], ["Meier", "Mayr"], ["Meier", "Meoer"]]

For $i = 0 To UBound($aPairs) - 1
	Local $sA = $aPairs[$i][0], $sB = $aPairs[$i][1]

	Local $aLeven = _fs_Levenshtein($sA, $sB), _              ; levenshtein distance
	$aOSA = _fs_OSA($sA, $sB), _                              ; optimal string alignment distance
	$aHamming = _fs_hamming($sA, $sB), _                      ; hamming distance
	$aSoundex = _fs_Soundex_distance($sA, $sB), _             ; soundex distance
	$aSoundexGerman = _fs_SoundexGerman_distance($sA, $sB), _ ; german soundex distance
	$aCologne = _fs_Cologne_distance($sA, $sB), _             ; cologne phonetics distance
	$aKbd = _fs_keyboard_distance_Strings($mKeyb, $sA, $sB)   ; keyboard key (QWERTY) distance

	ConsoleWrite(StringFormat("%15s %15s % 6d (%3d%%) % 6d (%3d%%) % 4d (%3d%%) % 4d (%3d%%) % 6d (%3d%%) % 4d (%3d%%) % 4d (%3d%%)\r\n", _
		$sA, $sB, _
		$aHamming[0], $aHamming[1] * 100, _
		$aLeven[0], $aLeven[1] * 100, _ ; distance (similarity)
		$aOSA[0], $aOSA[1] * 100, _ ; distance (similarity)
		$aSoundex[0], $aSoundex[1] * 100, _ ; distance (similarity)
		$aSoundexGerman[0], $aSoundexGerman[1] * 100, _ ; distance (similarity)
		$aCologne[0], $aCologne[1] * 100, _ ; distance (similarity)
		$aKbd[0], $aKbd[1] * 100 _ ; distance (similarity)
	))
Next

