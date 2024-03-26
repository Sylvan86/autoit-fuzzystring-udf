This UDF provides algorithms for fuzzy string comparison and the associated similarity search in string arrays.

It offers functions for character-based comparisons, comparing the phonetics of words and the geometric distance of characters on a keyboard.

In this way, typing errors can be recognized, similar-sounding words can be detected and other spellings of words can be included in further processing.

The function list of the UDF:
| Function | description |
|---- | --- |
| ***fuzzy array handling*** |
|`_FS_ArraySearchFuzzy`          | finds similar entries for a search value in an array |
|`_FS_ArrayToPhoneticGroups`   | groups the values of an array according to their phonetics |
| ***character-based metrics*** | |
|`_FS_Levenshtein`           | calculate the levenshtein distance between two strings |
|`_FS_OSA` | calculate the OSA ("optimal string alignment") between two strings |
|`_FS_Hamming`          | calculate the hamming distance between two strings |
| ***phonetic metrics*** |
|`_FS_Soundex_getCode`    | calculate the soundex code for a given word to represent the pronounciation in english |
|`_FS_Soundex_distance`   | calculate the soundex-pattern for both input values |
|`_FS_SoundexGerman_getCode`   | calculate the modified soundex code for german language for a given word to represent the pronounciation in german |
|`_FS_SoundexGerman_distance`        | calculate the soundexGerman-pattern for both input values |
|`_FS_Cologne_getCode`  | calculate the cologne phonetics code for german language for a given word to represent the pronounciation in german |
|`_FS_Cologne_distance` | calculate the cologne phonetics distance between both input values |
| ***key-position based metrics*** |
|`_FS_Keyboard_GetLayout`| return a map with coordinates for the characters for using in _FS_Keyboard_Distance_Chars() |
|` _FS_Keyboard_Distance_Chars`  | calculates the geometric key spacing between two characters on a keyboard |
|`_FS_Keyboard_Distance_Strings`          | calculate the keyboard-distance between two strings |
