#include-once
#include <Array.au3>

; #INDEX# =======================================================================================================================
; Title .........: Fuzzy string UDF
; AutoIt Version : 3.3.16.1
; Description ...: Metrics and functions for the fuzzy comparison and search of strings
; Author(s) .....: Andreas Tharang (AspirinJunkie)
; ===============================================================================================================================


; #CURRENT# =====================================================================================================================
; --------- fuzzy array handling:
; _FS_ArraySearchFuzzy           - finds similar entries for a search value in an array
; _FS_ArrayToPhoneticGroups      - groups the values of an array according to their phonetics
;
; --------- character-based metrics:
; _FS_Levenshtein                - calculate the levenshtein distance between two strings
; _FS_OSA                        - calculate the OSA ("optimal string alignment") between two strings
; _FS_Hamming                    - calculate the hamming distance between two strings
;
; --------- phonetic metrics:
; _FS_Soundex_getCode            - calculate the soundex code for a given word to represent the pronounciation in english
; _FS_Soundex_distance           - calculate the soundex-pattern for both input values
; _FS_SoundexGerman_getCode      - calculate the modified soundex code for german language for a given word to represent the pronounciation in german
; _FS_SoundexGerman_distance     - calculate the soundexGerman-pattern for both input values
; _FS_Cologne_getCode            - calculate the cologne phonetics code for german language for a given word to represent the pronounciation in german
; _FS_Cologne_distance           - calculate the cologne phonetics distance between both input values
;
; --------- key-position based metrics:
; _FS_Keyboard_GetLayout         - return a map with coordinates for the characters for using in _FS_Keyboard_Distance_Chars()
; _FS_Keyboard_Distance_Chars    - calculates the geometric key spacing between two characters on a keyboard
; _FS_Keyboard_Distance_Strings  - calculate the keyboard-distance between two strings
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __FS_Min2
; __FS_Min3
; ===============================================================================================================================


#Region fuzzy array handling

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_ArraySearchFuzzy
; Description ...: Groups the values of an array according to their phonetics
; Syntax.........: _FS_groupByPhonetic($aArray [, $cbPhoneticCode = _FS_Soundex_getCode [, $iCol = 0]])
; Parameters ....: $aArray:         {Array 1D} Array with values to be grouped
;                  $cbPhoneticCode: {Function($sWord) -> String} Function that calculates a phonetic code for a value
; Return values .: Success: {Map{PhoneticCode: Array[Values]} : Map with the phonetic code as key and an array with corresponding values as value.
;                  Failure: Null and set error to:
;                           | 1: $cbPhoneticCode is not a function
;                           | 2: $aArray is not an array
;                           | 3: $aArray has more than 2 dimensions
; Author ........: AspirinJunkie
; Modified.......: 2024-03-26
; Related .......: _FS_Soundex_getCode, _FS_SoundexGerman_getCode, _FS_Cologne_getCode
; ===============================================================================================================================
Func _FS_groupByPhonetic($aArray, $cbPhoneticCode = _FS_Soundex_getCode)
	Local $mRet[], $sCode, $aValues

	If Not IsFunc($cbPhoneticCode) Then Return SetError(1, 0, Null)

	Switch UBound($aArray, 0)
		Case 0
			Return SetError(2, 0, Null)

		Case 1
			For $i = 0 To UBound($aArray, 1) - 1
				$sCode = $cbPhoneticCode($aArray[$i])
				If @error Then ContinueLoop

				If MapExists($mRet, $sCode) Then
					$aValues = $mRet[$sCode]
					Redim $aValues[UBound($aValues) + 1]
					$aValues[UBound($aValues) - 1] = $aArray[$i]
				Else
					Local $aValues[1] = [$aArray[$i]]
				EndIf
				$mRet[$sCode] = $aValues
			Next

		Case Else ; n-dim arrays
			Return SetError(3, 0, Null)

	EndSwitch

	Return $mRet
EndFunc


; #FUNCTION# ====================================================================================================================
; Name...........: _FS_ArraySearchFuzzy
; Description ...: Finds similar entries for a search value in an array
; Syntax.........: _FS_ArraySearchFuzzy($aArray, $sSearchTerm [, $iMax = 1 [, $cbDistance = _FS_Levenshtein [, $iCol = 0 [, $bSort = True]]]])
; Parameters ....: $aArray:         {Array 1D/2D} Array to be searched for similar entries
;                  $sSearchTerm:    {String} Search value for which similar entries are to be found
;                  $sB:             {String} string to be compared with $sA
;                  $iMax:           depends on the value range:
;                                   | {Int: 0.. } Maximum distance to be considered similar
;                                   | {Float: 0.0 .. 1.0} Minimal similarity to still be used as a find
;                  $cbDistance:     {Function($sA, $sB, $iMax) -> Array[0.., 0.0 .. 1.0]} callback function for distance calculation
;                                   examples: _FS_Soundex_distance, _FS_Hamming, _FS_Levenshtein, _FS_OSA
;                  $iCol:          	{Int} If $aArray = 2D - column index for the comparison value
;                  $bSort: 			{Boolean} if True: result is sorted descending by similarity
; Return values .: Success: {Array[n][3..]} with:
;                           | $aArray[-2]: {Int:   0.. } distance
;                           | $aArray[-1]: {Float: 0.0..1.0 } similarity in percent
;                  Failure: Null and set error to:
;                           | 1: $cbDistance is not a function
;                           | 2: $aArray is not an array
;                           | 3: $aArray has more than 2 dimensions
;                           | 4: $iCol is out of range
; Author ........: AspirinJunkie
; Modified.......: 2024-03-26
; Related .......: _FS_Keyboard_GetLayout, _FS_Keyboard_Distance_Chars, __FS_Min3, __FS_Min2
; Example .......: $aAutoItFuncs = StringRegExp(FileRead(StringLeft(@AutoItExe, StringInStr(@AutoItExe, "\", 1, -1)) & 'SciTe\api\au3.api'), '(?m)^\w+(?=\h*\()', 3)
;                  $aSimilar = _FS_ArraySearchFuzzy($aAutoItFuncs, "StringIsAlpha", 4)
;                  _ArrayDisplay($aSimilar, "Similar functions", "", 64, "|", "function|distance|similarity [%]")
; ===============================================================================================================================
Func _FS_ArraySearchFuzzy($aArray, $sSearchTerm, $iMax = 1, $cbDistance = _FS_Levenshtein , $iCol = 0, $bSort = True)
	Local $bSimilarity = IsFloat($iMax) And $iMax <= 1.0 And $iMax >= 0.0 ; type of $iMax
	If Not IsFunc($cbDistance) Then Return SetError(1, 0, Null)

	Switch UBound($aArray, 0)
		Case 0
			Return SetError(2, 0, Null)

		Case 1 ; 1D array
			Local $iX = 0, $aComp, $aRet[UBound($aArray)][3]

			; filter the values in-place
			For $i = 0 To UBound($aArray) - 1
				$aComp = $cbDistance($sSearchTerm, $aArray[$i], $iMax)
				If @error Then ContinueLoop

				; check if words are similar
				If $bSimilarity Then
					If $aComp[1] < $iMax Then ContinueLoop
				ElseIf $aComp[0] > $iMax Then
					ContinueLoop
				EndIf

				$aRet[$iX][0] = $aArray[$i]
				$aRet[$iX][1] = $aComp[0]
				$aRet[$iX][2] = $aComp[1]
				$iX += 1
			Next
			Redim $aRet[$iX][3]
			If $bSort Then _ArraySort($aRet, 1, 0, 0, 2)
			Return $aRet

		Case 2 ; 2D array
			Local $nCols = UBound($aArray, 2)
			If $iCol < 0 Or $iCol >= $nCols Then Return SetError(4, $nCols, Null)

			Local $iX = 0, $aComp
			Redim $aArray[UBound($aArray)][$nCols + 2]

			; filter the values in-place
			For $i = 0 To UBound($aArray) - 1
				$aComp = $cbDistance($sSearchTerm, $aArray[$i][$iCol], $iMax)
				If @error Then ContinueLoop

				; check if words are similar
				If $bSimilarity Then
					If $aComp[1] < $iMax Then ContinueLoop
				ElseIf $aComp[0] > $iMax Then
					ContinueLoop
				EndIf

				For $j = 0 To $nCols - 1
					$aArray[$iX][$j] = $aArray[$i][$j]
				Next
				$aArray[$iX][$nCols] = $aComp[0]
				$aArray[$iX][$nCols + 1] = $aComp[1]
				$iX += 1
			Next

			Redim $aArray[$iX][$nCols + 2]
			If $bSort Then _ArraySort($aArray, 1, 0, 0, $nCols + 1)
			Return $aArray

		Case Else ; n-dim arrays
			Return SetError(3, 0, Null)

	EndSwitch
EndFunc

#EndRegion fuzzy array handling

#Region character-based metrics

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_Levenshtein
; Description ...: calculate the levenshtein distance between two strings
; Syntax.........: _FS_Levenshtein($sA, $sB, [$iMax = Default])
; Parameters ....: $sA:     {String} first string to be compared
;                  $sB:     {String} string to be compared with $sA
;                  $iMax:   depends on the value range (set to improve performance):
;                           | {Int: 0.. } maximum distance
;                           | {Float: 0.0 .. 1.0} minimum similarity in percent
;                           | {Default} no limit
; Return values .: Success: {Array[2]} with:
;                           | $aArray[0]: {Int:   0.. } distance in chars
;                           | $aArray[1]: {Float: 0.0..1.0 } similarity in percent
;                           set @extended to 1 if $iMax is reached
;                  Failure: ___ and set error to:
;                           | @error = 1 : $aA is not a 1D/2D array
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Related .......: __FS_Min3
; Example .......: Local $aLeven = _FS_Levenshtein("plauge", "plague")
;                  MsgBox(0, "Levenshtein distance", "distance: " & $aLeven[0] & @CRLF & "similarity: " & Round($aLeven[1]*100, 1) & " %")
; ===============================================================================================================================
Func _FS_Levenshtein($sA, $sB, $iMax = Default)
	Local $i, $j

    Local $iLenA = StringLen($sA)
    Local $iLenB = StringLen($sB)
	Local $iLenMax = $iLenA > $iLenB ? $iLenA : $iLenB
	If IsKeyword($iMax) = 1 Then $iMax = $iLenMax

	; $iMax = Minimum similarity
	If IsFloat($iMax) And $iMax <= 1.0 And $iMax >= 0.0 Then $iMax = Round((1.0 - $iMax) * $iLenMax)
	If $iMax <= 0 Then $iMax = 1

	; early leaving:
	If Abs($iLenA - $iLenB) > $iMax Then
		Local $aReturn[2] = [$iMax + 1, 1.0 - ($iMax + 1) / $iLenMax]
		Return SetExtended(1, $aReturn)
	EndIf

	; string to char array
	Local $aA = StringSplit($sA, "", 2)
	Local $aB = StringSplit($sB, "", 2)

	Local $iBound = $iLenB > $iMax ? $iMax : $iLenB

	; previous values
	Local $aPrev[$iBound + 2 + $iLenB]
	For $i = 0 To UBound($aPrev) - 1
		$aPrev[$i] = $i > $iBound ? 9223372036854700000 : $i
	Next

	; current values
	Local $aCurrent[$iLenB + 1]
	For $i = 1 To $iLenB
		$aCurrent[$i] = 9223372036854700000
	Next

	Local $iCostSubstitution, $iStart, $iEnd
	For $i = 0 To $iLenA - 1
		$aCurrent[0] = $i + 1

		; bounds for current stripe
		$iStart = ($i-$iMax) > 0 ? $i-$iMax : 0
		$iEnd = (($i+$iMax) < $iLenB ? $i+$iMax : $iLenB) - 1

		; ignore left of leftmost
		If $iStart > 0 Then $aCurrent[$iStart] = 9223372036854700000

		; loop over stripe
		For $j = $iStart To $iEnd
			$iCostSubstitution = $aA[$i] = $aB[$j] ? 0 : 1

			$aCurrent[$j + 1] = __FS_Min3($aPrev[$j + 1] + 1, $aCurrent[$j] + 1, $aPrev[$j] + $iCostSubstitution)  ; deletion, insertion, substitution
		Next

		$aPrev = $aCurrent
	Next

	Local $iDist = $aPrev[$iLenB] = 9223372036854700000 ? $iMax + 1 : $aPrev[$iLenB]

	Local $aReturn[2] = [$iDist, 1.0 - $iDist / $iLenMax]
	Return $iDist > $iMax ? SetExtended(1, $aReturn) : $aReturn
EndFunc


; #FUNCTION# ====================================================================================================================
; Name...........: _FS_OSA
; Description ...: calculate the OSA ("optimal string alignment") between two strings
; Syntax.........: _FS_OSA($sA, $sB, [$iMax = Default])
; Parameters ....: $sA:     {String} first string to be compared
;                  $sB:     {String} string to be compared with $sA
;                  $iMax:   depends on the value range (set to improve performance):
;                           | {Int: 0.. } maximum distance
;                           | {Float: 0.0 .. 1.0} minimum similarity in percent
;                           | {Default} no limit
; Return values .: Success: {Array[2]} with:
;                           | $aArray[0]: {Int:   0.. } distance in chars
;                           | $aArray[1]: {Float: 0.0..1.0 } similarity in percent
;                           set @extended to 1 if $iMax is reached
;                  Failure: Null and set error to:
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Remarks .......: includes the operations deletion, insertion and substitution and transposition with restriction: substrings are edit only once
;                  The OSA therefore does not fully represent the Damerau-Levenshtein distance
; Related .......: __FS_Min2
; Example .......: Local $aOSA = _FS_OSA("plauge", "plague")
;                  MsgBox(0, "OSA distance", "distance: " & $aOSA[0] & @CRLF & "similarity: " & Round($aOSA[1]*100, 1) & " %")
; ===============================================================================================================================
Func _FS_OSA($sA, $sB, $iMax = Default)
	Local $i, $j

    Local $iLenA = StringLen($sA)
    Local $iLenB = StringLen($sB)
	Local $iLenMax = $iLenA > $iLenB ? $iLenA : $iLenB
	If IsKeyword($iMax) = 1 Then $iMax = $iLenMax

	; $iMax = Minimum similarity
	If IsFloat($iMax) And $iMax <= 1.0 And $iMax >= 0.0 Then $iMax = Round((1.0 - $iMax) * $iLenMax)

	; early leaving:
	If Abs($iLenA - $iLenB) > $iMax Then
		Local $aReturn[2] = [$iMax + 1, 1.0 - ($iMax + 1) / $iLenMax]
		Return SetExtended(1, $aReturn)
	EndIf

	If $iLenA = 0 Or $iLenB = 0 Then
		Local $aReturn[2] = [$iLenMax, 0]
		Return $aReturn
	EndIf

	; initialize the distance matrix
    Local $aDistances[$iLenA + 1][$iLenB + 1]
    For $i = 0 To $iLenA
        $aDistances[$i][0] = $i
    Next
    For $j = 0 To $iLenB
        $aDistances[0][$j] = $j
    Next

    ; calculate the osa
	Local $aA = StringSplit($sA, "", 2)
	Local $aB = StringSplit($sB, "", 2)
	Local $iCost = 0, $i, $j
    For $i = 1 To $iLenA
        For $j = 1 To $iLenB
			$iCost = $aA[$i-1] = $aB[$j-1] ? 0 : 1
            $iDistMin = __FS_Min3($aDistances[$i - 1][$j] + 1, $aDistances[$i][$j - 1] + 1, $aDistances[$i - 1][$j - 1] + $iCost)
            $aDistances[$i][$j] = $iDistMin

			if $i > 1 _
				And $j > 1 _
				And $aA[$i-1] = $aB[$j - 2] _
				And $aA[$i-2] = $aB[$j - 1] Then _
				$aDistances[$i][$j] = _
				__FS_Min2($aDistances[$i][$j], $aDistances[$i-2][$j - 2] + $iCost) ; transposition

		Next
    Next

    ; return the OSA distance
	Local $iDist = $aDistances[$iLenA][$iLenB] = 9223372036854700000 ? $iMax + 1 : $aDistances[$iLenA][$iLenB]

	Local $aReturn[2] = [$iDist, 1.0 - $iDist / $iLenMax]
	Return $iDist > $iMax ? SetExtended(1, $aReturn) : $aReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_Hamming
; Description ...: calculate the hamming distance between two strings
; Syntax.........: _FS_Hamming($sA, $sB, [$bSimilarity = False])
; Parameters ....: $sA:     {String} first string to be compared
;                  $sB:     {String} string to be compared with $sA
; Return values .: Success: {Array[2]} with:
;                           | $aArray[0]: {Int:   0.. } distance in chars
;                           | $aArray[1]: {Float: 0.0..1.0 } similarity in percent
;                           set @extended to Max(Len($sA), Len($sB))
;                  Failure: Null and set error to:
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Remarks .......: includes the operation substitution only
; Example .......: Local $aHamm = _FS_Hamming("plauge", "plague")
;                  MsgBox(0, "hamming distance", "distance: " & $aHamm[0] & @CRLF & "similarity: " & Round($aHamm[1]*100, 1) & " %")
; ===============================================================================================================================
Func _FS_Hamming($sA, $sB, $iDummy = NULL)
	Local Const $nA = StringLen($sA), $nB = StringLen($sB)
	Local $nMax = $nA > $nB ? $nA : $nB

	; fill with spaces to create the same length
	If $nA < $nMax Then $sA = StringFormat("%-" & $nMax & "s", $sA)
	If $nB < $nMax Then $sB = StringFormat("%-" & $nMax & "s", $sB)

	Local $aA = StringSplit($sA, "", 2)
	Local $aB = StringSplit($sB, "", 2)
	Local $iHamming = 0, $i
	For $i = 0 To $nMax - 1
		If $aA[$i] <> $aB[$i] Then $iHamming += 1
	Next

	Local $aReturn[2] = [$iHamming, 1.0 - $iHamming / $nMax]
	Return SetExtended($nMax, $aReturn)
EndFunc

#EndRegion


#Region phonetic algorithms

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_Soundex_distance
; Description ...: calculate the soundex-pattern for both input values
;                  and calculate the levenshtein distance between them as a value for phonetic similarity
; Syntax.........: _FS_Soundex_distance($sA, $sB)
; Parameters ....: $sA:     {String} first string to be compared
;                  $sB:     {String} string to be compared with $sA
; Return values .: Success: {Array[2]} with:
;                           | $aArray[0]: {Int:   0.. } distance in chars
;                           | $aArray[1]: {Float: 0.0..1.0 } similarity in percent
;                  Failure: [-1,0] and set error to:
;                           | @error = 1 : error during create the soundex-code for $sA
;                           | @error = 2 : error during create the soundex-code for $sB
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Related .......: _FS_Soundex_getCode, _FS_Levenshtein
; Example .......: Local $aSE = _FS_Soundex_distance("plauge", "plague")
;                  MsgBox(0, "soundex distance", "distance: " & $aSE[0] & @CRLF & "similarity: " & Round($aSE[1]*100, 1) & " %")
; ===============================================================================================================================
Func _FS_Soundex_distance($sA, $sB, $iDummy = NULL)
	Local $sCodeA = _FS_Soundex_getCode($sA)
	If @error Then
		Local $aReturn[2] = [-1, 0]
		Return SetError(1, @error, $aReturn)
	EndIf
	Local $sCodeB = _FS_Soundex_getCode($sB)
	If @error Then
		Local $aReturn[2] = [-1, 0]
		Return SetError(2, @error, $aReturn)
	EndIf

	Return _FS_Levenshtein($sCodeA, $sCodeB)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_Soundex_getCode
; Description ...: calculate the soundex code for a given word to represent the pronounciation in english
; Syntax.........: _FS_Soundex_getCode($sWord)
; Parameters ....: $sWord: {String}
; Return values .: Success: The word for which the soundex code is to be determined
;                  Failure: "" and set error to:
;                           | @error = 1 : Insufficient structure of the word to create code
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; ===============================================================================================================================
Func _FS_Soundex_getCode($sWord)
	; static code-map - build only once
	Local Static $mCodes[]
	If UBound($mCodes) = 0 Then
		Local $aCodes[18][2] = [["B", "1"],["F", "1"],["P", "1"],["V", "1"],["C", "2"],["G", "2"],["J", "2"],["K", "2"],["Q", "2"],["S", "2"],["X", "2"],["Z", "2"],["D", "3"],["T", "3"],["L", "4"],["M", "5"],["N", "5"],["R", "6"]]
		For $i = 0 To 17
			$mCodes[$aCodes[$i][0]] = $aCodes[$i][1]
		Next
	EndIf

	$sWord = StringUpper($sWord)
	Local $sCode = StringLeft($sWord, 1), $sDigit, $sPrev = 0

	; loop over soundex-characters (ignore double)
	Local $aChars = StringRegExp(StringTrimLeft($sWord, 1), '([BCDFGJ-NP-TVXZ])\1?', 3)
	If @error Then Return SetError(1, @error, "")
	For $cChar In $aChars
		$sDigit = $mCodes[$cChar]
		If $sDigit = $sPrev Then ContinueLoop

		$sCode &= $sDigit
		If StringLen($sCode) = 4 Then ExitLoop
		$sPrev = $sDigit
	Next

	Return StringLeft($sCode & "000", 4)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_SoundexGerman_distance
; Description ...: calculate the soundexGerman-pattern for both input values
;                  and calculate the levenshtein distance between them as a value for phonetic similarity
; Syntax.........: _FS_SoundexGerman_distance($sA, $sB)
; Parameters ....: $sA:     {String} first string to be compared
;                  $sB:     {String} string to be compared with $sA
; Return values .: Success: {Array[2]} with:
;                           | $aArray[0]: {Int:   0.. } distance in chars
;                           | $aArray[1]: {Float: 0.0..1.0 } similarity in percent
;                  Failure: [-1,0] and set error to:
;                           | @error = 1 : error during create the soundex-code for $sA
;                           | @error = 2 : error during create the soundex-code for $sB
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Remarks .......: Modified form of soundex, which fits better with German words
; Related .......: _FS_SoundexGerman_getCode, _FS_Levenshtein
; Example .......: Local $aSEG = _FS_SoundexGerman_distance("Meier", "Mayr")
;                  MsgBox(0, "soundex distance", "distance: " & $aSEG[0] & @CRLF & "similarity: " & Round($aSEG[1]*100, 1) & " %")
; ===============================================================================================================================
Func _FS_SoundexGerman_distance($sA, $sB, $iDummy = NULL)
	Local $sCodeA = _FS_SoundexGerman_getCode($sA)
	If @error Then
		Local $aReturn[2] = [-1, 0]
		Return SetError(1, @error, $aReturn)
	EndIf
	Local $sCodeB = _FS_SoundexGerman_getCode($sB)
	If @error Then
		Local $aReturn[2] = [-1, 0]
		Return SetError(2, @error, $aReturn)
	EndIf

	Return _FS_Levenshtein($sCodeA, $sCodeB)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_SoundexGerman_getCode
; Description ...: calculate the modified soundex code for german language for a given word to represent the pronounciation in german
; Syntax.........: _FS_Soundex_getCode($sWord)
; Parameters ....: $sWord: {String}
; Return values .: Success: The word for which the soundex code is to be determined
;                  Failure: "" and set error to:
;                           | @error = 1 : Insufficient structure of the word to create code
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Remarks .......: modified table for better fitting to the german language
; ===============================================================================================================================
Func _FS_SoundexGerman_getCode($sWord)
	; static code-map - build only once
	Local Static $mCodes[]
	If UBound($mCodes) = 0 Then
		Local $aCodes[20][2] = [["B", "1"],["P", "1"],["F", "1"],["V", "1"],["W", "1"],["C", "2"],["G", "2"],["K", "2"],["Q", "2"],["X", "2"],["S", "2"],["Z", "2"],["ß", "2"],["D", "3"],["T", "3"],["L", "4"],["M", "5"],["N", "5"],["R", "6"],["CH", "7"]]
		For $i = 0 To 19
			$mCodes[$aCodes[$i][0]] = $aCodes[$i][1]
		Next
	EndIf

	$sWord = StringUpper($sWord)

	Local $sCode = StringLeft($sWord, 1), $sDigit, $sPrev = 0

	; loop over soundex-characters (ignore double)
	Local $aChars = StringRegExp(StringTrimLeft($sWord, 1), '(CH|[BCDFGK-NP-TVWXZß])\1?', 3)
	If @error Then Return SetError(1, @error, "")
	For $cChar In $aChars
		$sDigit = $mCodes[$cChar]
		If $sDigit = $sPrev Then ContinueLoop

		$sCode &= $sDigit
		If StringLen($sCode) = 4 Then ExitLoop
		$sPrev = $sDigit
	Next

	Return StringLeft($sCode & "000", 4)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_Cologne_distance
; Description ...: calculate the cologne phonetics pattern for both input values
;                  and calculate the levenshtein distance between them as a value for phonetic similarity
; Syntax.........: _FS_Cologne_distance($sA, $sB)
; Parameters ....: $sA:     {String} first string to be compared
;                  $sB:     {String} string to be compared with $sA
; Return values .: Success: {Array[2]} with:
;                           | $aArray[0]: {Int:   0.. } distance in chars
;                           | $aArray[1]: {Float: 0.0..1.0 } similarity in percent
;                  Failure: [-1,0] and set error to:
;                           | @error = 1 : error during create the soundex-code for $sA
;                           | @error = 2 : error during create the soundex-code for $sB
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Remarks .......: cologne phonetics was explicitly designed for the German language and therefore works best with it
; Related .......: _FS_cologne_getCode, _FS_Levenshtein
; Example .......: Local $aCologne = _FS_Cologne_distance("Meier", "Mayr")
;                  MsgBox(0, "cologne phonetics distance", "distance: " & $aCologne[0] & @CRLF & "similarity: " & Round($aCologne[1]*100, 1) & " %")
; ===============================================================================================================================
Func _FS_Cologne_distance($sA, $sB, $iDummy = NULL)
	Local $sCodeA = _FS_cologne_getCode($sA)
	If @error Then
		Local $aReturn[2] = [-1, 0]
		Return SetError(1, @error, $aReturn)
	EndIf
	Local $sCodeB = _FS_cologne_getCode($sB)
	If @error Then
		Local $aReturn[2] = [-1, 0]
		Return SetError(2, @error, $aReturn)
	EndIf

	Return _FS_Levenshtein($sCodeA, $sCodeB)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_Cologne_getCode
; Description ...: calculate the cologne phonetics code for german language for a given word to represent the pronounciation in german
; Syntax.........: _FS_Cologne_getCode($sInput)
; Parameters ....: $sWord: {String}
; Return values .: Success: The word for which the cologne phonetics code is to be determined
;                  Failure: "" and set error to:
;                           | @error = 1 : Insufficient structure of the word to create code
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Remarks .......: cologne phonetics was explicitly designed for the German language and therefore works best with it
; ===============================================================================================================================
Func _FS_Cologne_getCode($sInput)
	$sInput = StringUpper($sInput)
	$sInput = StringRegExpReplace($sInput, '[^A-ZÄÖÜß ]+', "")

	Local $sReturn = "", $iDigit, $iLast, $iGroup, $sWord
	For $sWord in StringSplit($sInput, " ", 2)
		$iLast = 9

		Local $aRegEx = StringRegExp($sWord, "(\A[AEIOUJYÄÖÜH])|((?|B|P(?!H)))|([DT])(?![CSZ])|([FVWP])|((?|\AC(?=[AHKLOQRUX])|[GKQ]|(?<![SZ])C(?=[AHKOQUX])))|(L)|([MN])|(R)|((?|[SZß]|C|[DT]|(?<=[CKQ])X))|(X)", 4)
		If @error Then Return SetError(1, @error, "")
		For $aChars in $aRegEx
			$iGroup = UBound($aChars)
			$iDigit = $iGroup > 10 ? 48 : $iGroup - 2
			If $iDigit <> $iLast Then $sReturn &= $iDigit
			$iLast = $iDigit
		Next
		$sReturn &= " "
	Next

	Return StringTrimRight($sReturn, 1)
EndFunc

#EndRegion



#Region keyboard-focused metrics

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_Keyboard_Distance_Strings
; Description ...: calculate the keyboard-distance between two strings. Modified OSA with keyboard-char distance as substitution cost
; Syntax.........: _FS_Keyboard_Distance_Strings(Byref $mKeyboard, $sA, $sB, [$iMax = Default, [$bEuclidian = True, [$fCostDeletion = 1, [$fCostInsertion = 1]]]])
; Parameters ....: $mKeyboard:      {Map[Arrays]} list of keyboard key coordinates as returned by _FS_Keyboard_GetLayout
;                  $sA:             {String} first string to be compared
;                  $sB:             {String} string to be compared with $sA
;                  $iMax:           depends on the value range (set to improve performance):
;                                   | {Int: 0.. } maximum distance
;                                   | {Float: 0.0 .. 1.0} minimum similarity in percent
;                                   | {Default} no limit
;                  $bEuclidian:     {Boolean} If True: euclidian distance (L2 norm) between coordinates is used
;                                             If False: manhattan distance (L1 norm/taxicab) between coordinates is used
;                  $fCostDeletion:  {Number} distance cost for a deletion to fine tune the distance result
;                  $fCostInsertion: {Number} distance cost for a insertion to fine tune the distance result
; Return values .: Success: {Array[2]} with:
;                           | $aArray[0]: {Int:   0.. } distance in chars
;                           | $aArray[1]: {Float: 0.0..1.0 } similarity in percent
;                           set @extended to 1 if $iMax is reached
;                  Failure: Null and set error to:
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Related .......: _FS_Keyboard_GetLayout, _FS_Keyboard_Distance_Chars, __FS_Min3, __FS_Min2
; Example .......: $mKeyb = _FS_Keyboard_GetLayout("QWERTY", False) ; case insensitive qwerty-layout
;                  $aKeyDist = _FS_Keyboard_Distance_Strings($mKeyb, "Meier", "Meyer")
;                  MsgBox(0, "keyboard distance", "distance: " & $aKeyDist[0] & @CRLF & "similarity: " & Round($aKeyDist[1]*100, 1) & " %")
; ===============================================================================================================================
Func _FS_Keyboard_Distance_Strings(Byref $mKeyboard, $sA, $sB, $iMax = Default, $bEuclidian = True, $fCostDeletion = 1, $fCostInsertion = 1)
	Local $i, $j

    Local $iLenA = StringLen($sA)
    Local $iLenB = StringLen($sB)
	Local $iLenMax = $iLenA > $iLenB ? $iLenA : $iLenB
	If IsKeyword($iMax) = 1 Then $iMax = $iLenMax

	; $iMax = Minimum similarity
	If IsFloat($iMax) And $iMax <= 1.0 And $iMax >= 0.0 Then $iMax = Round((1.0 - $iMax) * $iLenMax)

	; early leaving:
	If Abs($iLenA - $iLenB) > $iMax Then
		Local $aReturn[2] = [$iMax + 1, 1.0 - ($iMax + 1) / $iLenMax]
		Return SetExtended(1, $aReturn)
	EndIf

	If $iLenA = 0 Or $iLenB = 0 Then
		Local $aReturn[2] = [$iLenMax, 0]
		Return $aReturn
	EndIf

	; initialize the distance matrix
    Local $aDistances[$iLenA + 1][$iLenB + 1]
    For $i = 0 To $iLenA
        $aDistances[$i][0] = $i
    Next
    For $j = 0 To $iLenB
        $aDistances[0][$j] = $j
    Next

    ; calculate the osa keyboard distance
	Local $aA = StringSplit($sA, "", 2)
	Local $aB = StringSplit($sB, "", 2)
	Local $fCostSubstitution = 0, $i, $j
    For $i = 1 To $iLenA
        For $j = 1 To $iLenB
			$fCostSubstitution = _FS_Keyboard_Distance_Chars($mKeyboard, $aA[$i-1], $aB[$j-1], $bEuclidian)
            $iDistMin = __FS_Min3($aDistances[$i - 1][$j] + $fCostDeletion, $aDistances[$i][$j - 1] + $fCostInsertion, $aDistances[$i - 1][$j - 1] + $fCostSubstitution)
            $aDistances[$i][$j] = $iDistMin

			if $i > 1 _
				And $j > 1 _
				And $aA[$i-1] = $aB[$j - 2] _
				And $aA[$i-2] = $aB[$j - 1] Then _
				$aDistances[$i][$j] = _
				__FS_Min2($aDistances[$i][$j], $aDistances[$i-2][$j - 2] + $fCostSubstitution) ; transposition
		Next
    Next

	; return the OSA distance
	Local $iDist = $aDistances[$iLenA][$iLenB] = 9223372036854700000 ? $iMax + 1 : $aDistances[$iLenA][$iLenB]

	Local $fSimilarity = 1.0 - $iDist / $iLenMax
	Local $aReturn[2] = [$iDist, ($fSimilarity < 0 ? 0 : $fSimilarity) ]
	Return $iDist > $iMax ? SetExtended(1, $aReturn) : $aReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_Keyboard_Distance_Chars
; Description ...: Calculates the geometric key spacing between two characters on a keyboard
; Syntax.........: _FS_Keyboard_Distance_Chars(Byref $mKeyboard, $char1, $char2 [, $bEuclidian = False])
; Parameters ....: $mKeyboard:      {Map[Arrays]} list of keyboard key coordinates as returned by _FS_Keyboard_GetLayout or self created
;                  $char1:          {String} first char to be compared
;                  $char2:          {String} char to be compared with $char1
;                  $bEuclidian:     {Boolean} If True: euclidian distance (L2 norm) between coordinates is used
;                                             If False: manhattan distance (L1 norm/taxicab) between coordinates is used
; Return values .: Success: the geometric distance between the two keys
;                  Failure: Null and set error to:
;                           | @error = 1 : $char1 is not included in the keyboard layout $mKeyboard
;                           | @error = 2 : $char2 is not included in the keyboard layout $mKeyboard
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Remarks .......: If you add up the letter-by-letter distances in a word, you can deduce how long it takes to write the word.
; ===============================================================================================================================
Func _FS_Keyboard_Distance_Chars(Byref $mKeyboard, $char1, $char2, $bEuclidian = False)
	If $char1 == $char2 Then Return 0

	Local $aC1 = $mKeyboard[$char1]
	If Not IsArray($aC1) Then Return SetError(1, 0, Null)

	Local $aC2 = $mKeyboard[$char2]
	If Not IsArray($aC2) Then Return SetError(2, 0, Null)

	If $bEuclidian Then ; euclidian distance (L2-norm)
		Local $fSqSum = 0

		For $i = 0 To UBound($aC1) - 1
			$fSqSum += ($aC1[$i] - $aC2[$i])^2
		Next
		Return Sqrt($fSqSum)

	Else ; mahattan distance (L1-norm) or "taxicab" distance
		Local $fSum = 0

		For $i = 0 To UBound($aC1) - 1
			$fSum += Abs($aC1[$i] - $aC2[$i])
		Next
		Return $fSum
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _FS_Keyboard_GetLayout
; Description ...: return a map with coordinates for the characters that can be created with it for using in _FS_Keyboard_Distance_Chars()
;                  Currently, only QWERTY and QWERTZ are already available. In principle, however, any layout can be used with any number of coordinate dimensions. 
; Syntax.........: _FS_Keyboard_GetLayout([$vLayout = "QWERTY", [$bCaseSensitive = True, [$fZCoord = 0.5]]])
; Parameters ....: $vLayout:        {String} "QWERTY" or "QWERTZ" return the corresponding keyboard layout
;                                   {Array[n][1..]} keyboard key coordinate as 2D-Array
;                  $bCaseSensitive: {Boolean} True - Characters that are generated by a combination key such as Shift or Alt should receive an additional Z coordinate.
;                                             False - no additional z-coordinate for these characters
;                  $fZCoord:        {float} z-coordinate if bCaseSensitive is true
; Return values .: Success: {Map[Char:Array[n]]} map of key coordinates for using in _FS_Keyboard_Distance_Chars()
;                  Failure: -
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; Remarks .......: The coordinates of the two predefined layouts are very precise and also contain partial shifts between them
; ===============================================================================================================================
Func _FS_Keyboard_GetLayout($vLayout = "QWERTY", $bCaseSensitive = True, $fZCoord = 0.5)

	Local $mLayout[], $cChar, $aLayout
	Switch $vLayout
		Case "QWERTY"
			Local $aLayout = [["~","`","","",0,0],["1","!","¡","",1,0],["2","@","²","",2,0],["3","#","³","",3,0],["4","$","¤","£",4,0],["5","%","€","",5,0],["6","^","","",6,0],["7","&","","",7,0],["8","*","","",8,0],["9","(","","",9,0],["0",")","’","",10,0],["-","_","¥","",11,0],["=","+","×","÷",12,0],["q","","","ä",1.5,1],["w","","","å",2.5,1],["e","","","é",3.5,1],["r","","®","",4.5,1],["t","","","þ",5.5,1],["y","","","ü",6.5,1],["u","","","ú",7.5,1],["i","","","í",8.5,1],["o","","","ó",9.5,1],["p","","","ö",10.5,1],["[","{","«","",11.5,1],["]","}","»","",12.5,1],["\","|","¬","¦",13.5,1],["a","","","á",1.75,2],["s","","ß","§",2.75,2],["d","","","ð",3.75,2],["f","","","",4.75,2],["g","","","",5.75,2],["h","","","",6.75,2],["j","","","",7.75,2],["k","","","",8.75,2],["l","","","ø",9.75,2],[";",":","¶","°",10.75,2],["'",'"',"´","",11.75,2],["z","","","æ",2.25,3],["x","","","",3.25,3],["c","","©","",4.25,3],["v","","","",5.25,3],["b","","","",6.25,3],["n","","","ñ",7.25,3],["m","","μ","",8.25,3],[",","<","","ç",9.25,3],[".",">","","",10.25,3],["/","?","¿","",11.25,3],[' ',"","","",7,4]]
		Case "QWERTZ"
			Local $aLayout = [["^","°","",0,0],["1","!","",1,0],["2",'"',"²",2,0],["3","§","³",3,0],["4","$","",4,0],["5","%","",5,0],["6","&","",6,0],["7","/","{",7,0],["8","(","[",8,0],["9",")","]",9,0],["0","=","}",10,0],["ß","?","\",11,0],["´","`","",12,0],["q","","@",1.5,1],["w","","",2.5,1],["e","","€",3.5,1],["r","","",4.5,1],["t","","",5.5,1],["z","","",6.5,1],["u","","",7.5,1],["i","","",8.5,1],["o","","",9.5,1],["p","","",10.5,1],["ü","","",11.5,1],["+","*","~",12.5,1],["a","","",1.75,2],["s","","",2.75,2],["d","","",3.75,2],["f","","",4.75,2],["g","","",5.75,2],["h","","",6.75,2],["j","","",7.75,2],["k","","",8.75,2],["l","","",9.75,2],["ö","","",10.75,2],["ä","","",11.75,2],["#","'","",12.75,2],["<",">","|",1.25,3],["y","","",2.25,3],["x","","",3.25,3],["c","","",4.25,3],["v","","",5.25,3],["b","","",6.25,3],["n","","",7.25,3],["m","","μ",8.25,3],[",",";","",9.25,3],[".",":","",10.25,3],["-","_","",11.25,3],[' ',"","",7,4]]
	EndSwitch

	Local $iCols = UBound($aLayout, 2)

	For $i = 0 To UBound($aLayout, 1) - 1
		For $j = 0 To $iCols - 3
			$cChar = $aLayout[$i][$j]
			If $cChar = "" Then ContinueLoop

			Local $aCoords[3] = [$aLayout[$i][$iCols - 2], $aLayout[$i][$iCols - 1], ($j = 0 ? 0 : $fZCoord)]
			$mLayout[$cChar] = $aCoords

			; uppercase variant (get z-coord because you need shift-key)
			If Not (StringUpper($cChar) == $cChar) And Not MapExists($mLayout, StringUpper($cChar)) Then
				If $bCaseSensitive Then $aCoords[2] += $fZCoord
					$mLayout[StringUpper($cChar)] = $aCoords
				EndIf

		Next
	Next

	Return $mLayout
EndFunc

#EndRegion


#Region internal functions

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __FS_Min2
; Description ...: Determines the minimum of the two passed parameters
; Syntax.........: __FS_Min2($a, $b)
; Parameters ....: $a and $b
; Return values .: Success: min($a, $b)
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; ===============================================================================================================================
Func __FS_Min2($a, $b)
	Return $a < $b ? $a : $b
EndFunc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __FS_Min3
; Description ...: Determines the minimum of the three passed parameters
; Syntax.........: __FS_Min3($a, $b, $c)
; Parameters ....: $a, $b and $c
; Return values .: Success: min($a, $b, $c)
; Author ........: AspirinJunkie
; Modified.......: 2024-03-21
; ===============================================================================================================================
Func __FS_Min3($a, $b, $c)
    If $a < $b And $a < $c Then Return $a
    If $b < $a And $b < $c Then Return $b
    Return $c
EndFunc

#EndRegion

