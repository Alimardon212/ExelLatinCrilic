Attribute VB_Name = "Module1"
Function LatinToCyrillic(text As String) As String
    ' Katta harflar
    text = Replace(text, "A", ChrW(&H410)) ' ?
    text = Replace(text, "B", ChrW(&H411)) ' ?
    text = Replace(text, "V", ChrW(&H412)) ' ?
    text = Replace(text, "G", ChrW(&H413)) ' ?
    text = Replace(text, "D", ChrW(&H414)) ' ?
    text = Replace(text, "E", ChrW(&H415)) ' ?
    text = Replace(text, "Zh", ChrW(&H416)) ' ?
    text = Replace(text, "Z", ChrW(&H417)) ' ?
    text = Replace(text, "I", ChrW(&H418)) ' ?
    text = Replace(text, "Y", ChrW(&H419)) ' ?
    text = Replace(text, "K", ChrW(&H41A)) ' ?
    text = Replace(text, "L", ChrW(&H41B)) ' ?
    text = Replace(text, "M", ChrW(&H41C)) ' ?
    text = Replace(text, "N", ChrW(&H41D)) ' ?
    text = Replace(text, "O", ChrW(&H41E)) ' ?
    text = Replace(text, "P", ChrW(&H41F)) ' ?
    text = Replace(text, "R", ChrW(&H420)) ' ?
    text = Replace(text, "S", ChrW(&H421)) ' ?
    text = Replace(text, "T", ChrW(&H422)) ' ?
    text = Replace(text, "U", ChrW(&H423)) ' ?
    text = Replace(text, "F", ChrW(&H424)) ' ?
    text = Replace(text, "H", ChrW(&H425)) ' ?
    text = Replace(text, "Ts", ChrW(&H426)) ' ?
    text = Replace(text, "Ch", ChrW(&H427)) ' ?
    text = Replace(text, "Sh", ChrW(&H428)) ' ?
    text = Replace(text, "Sht", ChrW(&H429)) ' ?
    text = Replace(text, "A", ChrW(&H42A)) ' ?
    text = Replace(text, "Yu", ChrW(&H42E)) ' ?
    text = Replace(text, "Ya", ChrW(&H42F)) ' ?
    
    ' Kichik harflar
    text = Replace(text, "a", ChrW(&H430)) ' ?
    text = Replace(text, "b", ChrW(&H431)) ' ?
    text = Replace(text, "v", ChrW(&H432)) ' ?
    text = Replace(text, "g", ChrW(&H433)) ' ?
    text = Replace(text, "d", ChrW(&H434)) ' ?
    text = Replace(text, "e", ChrW(&H435)) ' ?
    text = Replace(text, "zh", ChrW(&H436)) ' ?
    text = Replace(text, "z", ChrW(&H437)) ' ?
    text = Replace(text, "i", ChrW(&H438)) ' ?
    text = Replace(text, "y", ChrW(&H439)) ' ?
    text = Replace(text, "k", ChrW(&H43A)) ' ?
    text = Replace(text, "l", ChrW(&H43B)) ' ?
    text = Replace(text, "m", ChrW(&H43C)) ' ?
    text = Replace(text, "n", ChrW(&H43D)) ' ?
    text = Replace(text, "o", ChrW(&H43E)) ' ?
    text = Replace(text, "p", ChrW(&H43F)) ' ?
    text = Replace(text, "r", ChrW(&H440)) ' ?
    text = Replace(text, "s", ChrW(&H441)) ' ?
    text = Replace(text, "t", ChrW(&H442)) ' ?
    text = Replace(text, "u", ChrW(&H443)) ' ?
    text = Replace(text, "f", ChrW(&H444)) ' ?
    text = Replace(text, "h", ChrW(&H445)) ' ?
    text = Replace(text, "ts", ChrW(&H446)) ' ?
    text = Replace(text, "ch", ChrW(&H447)) ' ?
    text = Replace(text, "sh", ChrW(&H448)) ' ?
    text = Replace(text, "sht", ChrW(&H449)) ' ?
    text = Replace(text, "a", ChrW(&H44A)) ' ?
    text = Replace(text, "yu", ChrW(&H44E)) ' ?
    text = Replace(text, "ya", ChrW(&H44F)) ' ?
    
    LatinToCyrillic = text
End Function

Function CyrillicToLatin(text As String) As String
    ' Katta harflar
    text = Replace(text, ChrW(&H410), "A") ' ?
    text = Replace(text, ChrW(&H411), "B") ' ?
    text = Replace(text, ChrW(&H412), "V") ' ?
    text = Replace(text, ChrW(&H413), "G") ' ?
    text = Replace(text, ChrW(&H414), "D") ' ?
    text = Replace(text, ChrW(&H415), "E") ' ?
    text = Replace(text, ChrW(&H416), "Zh") ' ?
    text = Replace(text, ChrW(&H417), "Z") ' ?
    text = Replace(text, ChrW(&H418), "I") ' ?
    text = Replace(text, ChrW(&H419), "Y") ' ?
    text = Replace(text, ChrW(&H41A), "K") ' ?
    text = Replace(text, ChrW(&H41B), "L") ' ?
    text = Replace(text, ChrW(&H41C), "M") ' ?
    text = Replace(text, ChrW(&H41D), "N") ' ?
    text = Replace(text, ChrW(&H41E), "O") ' ?
    text = Replace(text, ChrW(&H41F), "P") ' ?
    text = Replace(text, ChrW(&H420), "R") ' ?
    text = Replace(text, ChrW(&H421), "S") ' ?
    text = Replace(text, ChrW(&H422), "T") ' ?
    text = Replace(text, ChrW(&H423), "U") ' ?
    text = Replace(text, ChrW(&H424), "F") ' ?
    text = Replace(text, ChrW(&H425), "H") ' ?
    text = Replace(text, ChrW(&H426), "Ts") ' ?
    text = Replace(text, ChrW(&H427), "Ch") ' ?
    text = Replace(text, ChrW(&H428), "Sh") ' ?
    text = Replace(text, ChrW(&H429), "Sht") ' ?
    text = Replace(text, ChrW(&H42A), "A") ' ?
    text = Replace(text, ChrW(&H42E), "Yu") ' ?
    text = Replace(text, ChrW(&H42F), "Ya") ' ?
    
    ' Kichik harflar
    text = Replace(text, ChrW(&H430), "a") ' ?
    text = Replace(text, ChrW(&H431), "b") ' ?
    text = Replace(text, ChrW(&H432), "v") ' ?
    text = Replace(text, ChrW(&H433), "g") ' ?
    text = Replace(text, ChrW(&H434), "d") ' ?
    text = Replace(text, ChrW(&H435), "e") ' ?
    text = Replace(text, ChrW(&H436), "zh") ' ?
    text = Replace(text, ChrW(&H437), "z") ' ?
    text = Replace(text, ChrW(&H438), "i") ' ?
    text = Replace(text, ChrW(&H439), "y") ' ?
    text = Replace(text, ChrW(&H43A), "k") ' ?
    text = Replace(text, ChrW(&H43B), "l") ' ?
    text = Replace(text, ChrW(&H43C), "m") ' ?
    text = Replace(text, ChrW(&H43D), "n") ' ?
    text = Replace(text, ChrW(&H43E), "o") ' ?
    text = Replace(text, ChrW(&H43F), "p") ' ?
    text = Replace(text, ChrW(&H440), "r") ' ?
    text = Replace(text, ChrW(&H441), "s") ' ?
    text = Replace(text, ChrW(&H442), "t") ' ?
    text = Replace(text, ChrW(&H443), "u") ' ?
    text = Replace(text, ChrW(&H444), "f") ' ?
    text = Replace(text, ChrW(&H445), "h") ' ?
    text = Replace(text, ChrW(&H446), "ts") ' ?
    text = Replace(text, ChrW(&H447), "ch") ' ?
    text = Replace(text, ChrW(&H448), "sh") ' ?
    text = Replace(text, ChrW(&H449), "sht") ' ?
    text = Replace(text, ChrW(&H44A), "a") ' ?
    text = Replace(text, ChrW(&H44E), "yu") ' ?
    text = Replace(text, ChrW(&H44F), "ya") ' ?
    
    CyrillicToLatin = text
End Function

