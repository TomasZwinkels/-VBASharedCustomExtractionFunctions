Function MyOwnRegexReplace(myrange As String, strPattern As String, replacewithvalue As String) As String
    Dim regEx As New RegExp
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String
    
    If strPattern <> "" Then
        strInput = myrange

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.Test(strInput) Then
             
            MyOwnRegexReplace = regEx.Replace(myrange, replacewithvalue)
        
        Else
        
            MyOwnRegexReplace = myrange
            
        End If
    End If
    
End Function

Function GetTitles(myrange As Range) As String
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String

    strPattern = "Dre in legge |Dre\. in legge |Dr\. en droit et Dr\. ès sc\. écon\. |Dr\. h\. c\. |Dr\. en m.decine, |Dr\. en droit, |Dr\. en droit |Dr en droit |Dr\. |iur\. |phil\. |med\. |vet\. |jur\. |rer\. |pol\. |oec\. |publ\. |dent\. |nat\. |pharm\. |cam\. |HSG\. |rel\. |agr\. |inf\. |math\. |techn\. |poi\. |math\.  |sc\. |theol. |iur "

    If strPattern <> "" Then
        strInput = myrange.Value

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPattern
        End With

        If regEx.Test(strInput) Then
        
            Dim i As Integer
            Dim titlevec As String
             
            For i = 0 To regEx.Execute(myrange).Count - 1
               titlevec = titlevec & regEx.Execute(myrange)(i)
            Next i
            
            GetTitles = titlevec
        
        Else
            GetTitles = ""
        End If
    End If
End Function

Function SelectPartFromFirstCapitalToHypenOrEnd(WholeName As String)

  'first, only select everything from the capital onwards
     Dim regEx As New RegExp
     Dim strPattern As String
     Dim strInput As String
     
     strPattern = "[A-Z].+?(?=(-|\b))" '"[A-Z].+?(?=-)"
    
    If strPattern <> "" Then

    strInput = WholeName '.Value

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = False
            .Pattern = strPattern
        End With

            If regEx.Test(strInput) Then
            
                SelectPartFromFirstCapitalToHypenOrEnd = regEx.Execute(strInput)(0)
            Else
                SelectPartFromFirstCapitalToHypenOrEnd = "ERROR! Name input has an unexpected format!"
            End If
    
    End If

End Function

Function GetYear(WholeDateOfBirthString As String)

  'first, only select everything from the capital onwards
     Dim regEx As New RegExp
     Dim strPattern As String
     Dim strInput As String
     
     strPattern = "[1][0-9]{3}"
    
    If strPattern <> "" Then

    strInput = WholeDateOfBirthString '.Value

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = False
            .Pattern = strPattern
        End With

            If regEx.Test(strInput) Then
            
                GetYear = regEx.Execute(strInput)(0)
            Else
                GetYear = "Error! Pattern that looks like a year not found"
            End If
    
    End If

End Function

Function id_pers(LastName As String, FirstName As String, DateOfBirth As String) As String

mySep = "_"
text = ""
text = "CH" & mySep & SelectPartFromFirstCapitalToHypenOrEnd(LastName) & mySep & SelectPartFromFirstCapitalToHypenOrEnd(FirstName) & mySep & GetYear(DateOfBirth)

id_pers = text

End Function


Function GetFirstName(Cellvalue As Range)

    Dim LArray() As String
    Dim CellvalueWithoutTitle As String

    'remove title
        CellvalueWithoutTitle = Replace(Cellvalue, GetTitles(Cellvalue), "")

    'put vons together
        CellvalueWithVonClosed = VonReplace(CellvalueWithoutTitle)
    
    'split by space
        LArray = Split(CellvalueWithVonClosed, " ")
    
    'get rid of comma
        result = Replace(LArray(0), ",", "")
        GetFirstName = result

'GetFirstName = Left(Right(Cellvalue, (Len(Cellvalue)) - Len(GetTitles(Cellvalue))), WorksheetFunction.Find(" ", Right(Cellvalue, (Len(Cellvalue)) - Len(GetTitles(Cellvalue)))) - 1)

End Function

Function GetPlaceName(Cellvalue As Range)

CellvalueWithoutTitle = Replace(Cellvalue, GetTitles(Cellvalue), "")

GetPlaceName = Trim(Right(CellvalueWithoutTitle, Len(CellvalueWithoutTitle) - WorksheetFunction.Find(",", CellvalueWithoutTitle)))

End Function

Function VonReplace(Cellvalue As String) As String

    'get vons closed in (more can be added)
    Cellvalue = Replace(CStr(Cellvalue), "von ", "von") 'range was converted into string
    Cellvalue = Replace(CStr(Cellvalue), "de la ", "dela")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdu ", "du")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdi ", "di")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bda ", "da")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdel ", "del")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bde ", "de")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bde ", "de")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bde ", "de")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bde ", "de")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bde ", "de")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bde ", "de")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdello ", "dello")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdell' ", "dell'")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdei ", "dei")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdegli ", "degli")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdelle ", "delle")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdella ", "della")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdal ", "dal")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdall' ", "dall'")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdallo ", "dallo")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdai ", "dai")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdagli ", "dagli")
    Cellvalue = MyOwnRegexReplace(CStr(Cellvalue), "\bdalla ", "dalla")
    
    VonReplace = MyOwnRegexReplace(CStr(Cellvalue), "\bdalle ", "dalle")

End Function


Function GetLastName(Cellvalue As Range) As String

    Dim LArray() As String
    Dim CellvalueWithoutTitle As String

    'remove title
        CellvalueWithoutTitle = Replace(Cellvalue, GetTitles(Cellvalue), "")

    'put vons together
        CellvalueWithVonClosed = VonReplace(CellvalueWithoutTitle)
    
    'split by space
        LArray = Split(CellvalueWithVonClosed, " ")
    
    'get rid of comma
        result = Replace(LArray(1), ",", "")
        GetLastName = result

    'GetLastName = Trim(WorksheetFunction.Substitute(WorksheetFunction.Substitute(WorksheetFunction.Substitute(WorksheetFunction.Substitute(Cellvalue, GetTitles(Cellvalue), ""), GetFirstName(Cellvalue), ""), GetPlaceName(Cellvalue), ""), ",", ""))

End Function


Function CleanName(Cellvalue As String)

Dim almostclean As String

almostclean = StripAccent(Cellvalue)
almostclean = UMLAUT(almostclean)
almostclean = Replace(almostclean, "von ", "von")
almostclean = Replace(almostclean, "Von ", "Von")
almostclean = Replace(almostclean, "van ", "van")
almostclean = Replace(almostclean, "Van ", "Van")
almostclean = Replace(almostclean, "Van der", "Vander")
almostclean = Replace(almostclean, "Van Der", "VanDer")
almostclean = Replace(almostclean, "van der", "vander")
almostclean = Replace(almostclean, "van Der", "vanDer")

'von e.t.c.
CleanName = almostclean

End Function

Function StripAccent(thestring As String)

'function taken from https://www.extendoffice.com/documents/excel/707-excel-replace-accented-characters.html

Dim A As String * 1
Dim B As String * 1
Dim i As Integer
Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
For i = 1 To Len(AccChars)
A = Mid(AccChars, i, 1)
B = Mid(RegChars, i, 1)
thestring = Replace(thestring, A, B)
Next
StripAccent = thestring

End Function


Function UMLAUT(text As String) As String

'function taken from http://stackoverflow.com/questions/18850981/excel-user-defined-formula-to-eliminate-special-characters

    UMLAUT = Replace(text, "ü", "ue")
    UMLAUT = Replace(UMLAUT, "Ü", "Ue")
    UMLAUT = Replace(UMLAUT, "ä", "ae")
    UMLAUT = Replace(UMLAUT, "Ä", "Ae")
    UMLAUT = Replace(UMLAUT, "ö", "oe")
    UMLAUT = Replace(UMLAUT, "Ö", "Oe")
    UMLAUT = Replace(UMLAUT, "ß", "ss")
    
End Function

Function GetDistrict(Cellvalue As String)

GetDistrict = Trim(Right(Cellvalue, Len(Cellvalue) - Workbook.Function.Find(",", Cellvalue)))

End Function


Function GetDistrictWhenInSeperateCell(Cellvalue As String)

text = Cellvalue
text = Replace(text, "Wahlkreis ", "")
text = Replace(text, "Arrondissement de ", "")
text = Replace(text, "Arrondissement du ", "")
text = Replace(text, "Arrondissement de la ", "")
text = Replace(text, "Circondario ", "")

GetDistrict = text

End Function

Function GetDateOfBirth(myrange As String) As String
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String
    Dim BirthDateUnclean As String
    Dim regEx3 As New RegExp
    Dim strInput3 As String

    strPatternOrder = "[1][0-9][0-9][0-9].+?(Né en|Né le|Ne le|N6 le|N6 ä|N6 lc|ne ä|Né à|Né 1c|Geboren|Geb\.|Geh\.)"
    
    'check if the first year commes before the Geboren e.t.c. if so, trow an error!
        If strPatternOrder <> "" Then
        strInput3 = myrange

        With regEx3
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPatternOrder
        End With

            If regEx3.Test(strInput3) Then
            
                ' if this is the case, trow an error!
                GetDateOfBirth = "!!ERROR!! Date before 'Geboren' "
                Debug.Print "Yes I am running!"

            Else
                ' if not you can continue with the rest
    

                        strPattern = "(Né en|Né le|Ne le|N6 le|N6 ä|N6 lc|ne ä|Né à|Né 1c|Geboren|Geb\.|Geh\.).+?[1][0-9][0-9][0-9]"
                    
                        If strPattern <> "" Then
                            strInput = myrange
                    
                            With regEx
                                .Global = True
                                .MultiLine = False
                                .IgnoreCase = True
                                .Pattern = strPattern
                            End With
                    
                                If regEx.Test(strInput) Then
                                
                                       BirthDateUnclean = Trim(regEx.Execute(myrange)(0))
                                       GetDateOfBirth = Replace(BirthDateUnclean, "Geboren ", "")
                                       
                                       
                                Else
                                    GetDateOfBirth = ""
                                End If
                           
                        End If
                
        End If 'end the inner regEx3 if
    
    End If 'end the outer regEx3 if
                        
End Function

Function CleanDateOfBirth(text As String) As String

'first only selects what really looks like a date

Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String

    strPattern = "(([0-9][0-9]|[0-9]).+?(Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember|janvier|février|mars|avril|mai|juin|julliet|août|septembre|octobre|novembre|decembre|gennaio|febbraio|marzo|aprilemaggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre).+?[1][0-9][0-9][0-9])|([1][0-9][0-9][0-9])"

    If strPattern <> "" Then
        strInput = text

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With

            If regEx.Test(strInput) Then
            
                   text = regEx.Execute(strInput)(0)
                   
            Else
                text = "!!ERROR!! Nothing that looks like a date?!"
            End If

    End If


text = Replace(text, " ", "")

text = Replace(text, ".", "")

text = Replace(text, "Januar", "jan")
text = Replace(text, "Februar", "feb")
text = Replace(text, "März", "mar")
text = Replace(text, "April", "apr")
text = Replace(text, "Mai", "may")
text = Replace(text, "Juni", "jun")
text = Replace(text, "Juli", "jul")
text = Replace(text, "August", "aug")
text = Replace(text, "September", "sep")
text = Replace(text, "Oktober", "oct")
text = Replace(text, "November", "nov")
text = Replace(text, "Dezember", "dec")

text = Replace(text, "janvier", "jan")
text = Replace(text, "février", "feb")
text = Replace(text, "mars", "mar")
text = Replace(text, "avril", "apr")
text = Replace(text, "mai", "may")
text = Replace(text, "juin", "jun")
text = Replace(text, "julliet", "jul")
text = Replace(text, "août", "aug")
text = Replace(text, "septembre", "sep")
text = Replace(text, "octobre", "oct")
text = Replace(text, "novembre", "nov")
text = Replace(text, "decembre", "dec")

text = Replace(text, "il", "")
text = Replace(text, "gennaio", "jan")
text = Replace(text, "febbraio", "feb")
text = Replace(text, "marzo", "mar")
text = Replace(text, "aprile", "apr")
text = Replace(text, "maggio", "may")
text = Replace(text, "giugno", "jun")
text = Replace(text, "luglio", "jul")
text = Replace(text, "agosto", "aug")
text = Replace(text, "settembre", "sep")
text = Replace(text, "ottobre", "oct")
text = Replace(text, "novembre", "nov")
text = Replace(text, "dicembre", "dec")

If Not (IsNumeric(Mid(text, 2, 1))) Then
    text = 0 & text
End If

CleanDateOfBirth = text

End Function

Function GetPlaceOfBirth(myrange As String) As String
    Dim regEx As New RegExp
    Dim regEx2 As New RegExp
    Dim strPattern As String
    Dim strPatternPlace As String
    Dim strInput As String
    Dim strInput2 As String
    Dim strReplace As String
    Dim strOutput As String
    Dim DateAndPlace As String
    

    ' extract the whole bit, but date and place
    strPatternOverall = "(Né en|Né le|Ne le|N6 le|N6 ä|N6 lc|ne ä|Né à|Né 1c|Geboren)(.+?)([1][0-9][0-9][0-9])( à | in | als )?.+?\."

    If strPatternOverall <> "" Then
        strInput = myrange

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPatternOverall
        End With

            If regEx.Test(strInput) Then
            
                   DateAndPlace = regEx.Execute(myrange)(0)
                   
            Else
                GetPlaceOfBirth = ""
            End If
    
    End If
    
    'subselect the only the place bit
        strPatternPlace = "(à | in | als ).+?\."

    If strPatternPlace <> "" Then
        strInput2 = DateAndPlace

        With regEx2
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPatternPlace
        End With

            If regEx2.Test(strInput2) Then
            
                GetPlaceOfBirth = regEx2.Execute(strInput2)(0)
                
            Else
                GetPlaceOfBirth = ""
            End If
    
    End If
    
End Function

Function CleanPlaceOfBirth(PlaceOfBirthDirty As String)

    PlaceOfBirthDirty = Replace(PlaceOfBirthDirty, " in ", "")
    PlaceOfBirthDirty = Replace(PlaceOfBirthDirty, " à ", "")
    PlaceOfBirthDirty = Replace(PlaceOfBirthDirty, " als ", "")
    PlaceOfBirthDirty = Replace(PlaceOfBirthDirty, "Bürger von ", "")
    
    CleanPlaceOfBirth = Trim(Replace(PlaceOfBirthDirty, ".", ""))

End Function

Function GetCode1(myrange As String) As String
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String

    strPattern = "K |R |S |B |U |L |DE |T |V |P |C |R |LE |— "

    If strPattern <> "" Then
        strInput = Left(myrange, 5)

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPattern
        End With

            If regEx.Test(strInput) Then
            
                   GetCode1 = Trim(regEx.Execute(myrange)(0))
                   
            Else
                GetCode1 = ""
            End If
       
    End If
End Function

Function GetCode2(myrange As String) As String
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String

    strPattern = " VW | EP | GW | EGC "

    If strPattern <> "" Then
        strInput = myrange

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPattern
        End With

            If regEx.Test(strInput) Then
            
                   GetCode2 = Trim(regEx.Execute(myrange)(0))
                   
            Else
                GetCode2 = ""
            End If
       
    End If
End Function

Function GetKantonCitizenShip(myrange As String) As String
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String
    Dim KantonCitizenShipDirty As String
    Dim SearchStr As String

    strPattern = "(Von|Originaire de )(.+?\.)"
    
    'only try to find this in the first 100 characters
        SearchStr = myrange
        SearchStr = Left(SearchStr, 150)

    If strPattern <> "" Then
        strInput = SearchStr

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPattern
        End With

            If regEx.Test(strInput) Then
            
                   KantonCitizenShipDirty = Trim(regEx.Execute(myrange)(0)) 'always just return the first hit
                   
                   GetKantonCitizenShip = Replace(KantonCitizenShipDirty, "von ", "", 1, 1)
                   
                   
            Else
                GetKantonCitizenShip = ""
            End If
       
    End If
End Function

Function GetMaritalStatus(myrange As String)

    If InStr(myrange, "verheiratet") > 0 Or InStr(myrange, "Verheiratet") > 0 Then
       
        GetMaritalStatus = "married"
        
    Else
    
        GetMaritalStatus = ""
    
    End If

End Function

Function GetNumberOfChildren(myrange As String) As String
 
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String

    strPattern = "[\w]+(?=\ kind) (kindern|kinder|kind)"

    If strPattern <> "" Then
        strInput = myrange

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPattern
        End With

            If regEx.Test(strInput) Then
            
                GetNumberOfChildren = Trim(regEx.Execute(myrange)(0))
                   
            Else
                GetNumberOfChildren = ""
            End If
       
    End If

End Function

Function GuessGender(GermanString As String, FrenchString As String)

Dim StringStart As String
Dim GermanVote As String
Dim FrenchVote As String

' the german votes
    GermanStringStart = Left(GermanString, 100)
    
    If InStr(GermanStringStart, "Bürgerin") > 0 Or InStr(GermanStringStart, "Burgerin") > 0 Then
           
            GermanVote = "female"
            
        Else
        
            If InStr(GermanStringStart, "Bürger ") > 0 Or InStr(GermanStringStart, "Burger ") > 0 Or InStr(GermanStringStart, " er ") > 0 Then
            
            GermanVote = "male"
            
            Else
               
            GermanVote = "No know gender markers!"
            
            End If
        
        End If
    

' the french votes
    FrenchStringStart = Left(FrenchString, 100)
    
    If InStr(FrenchStringStart, "Née") > 0 Or InStr(FrenchStringStart, "Nata") > 0 Then
           
            FrenchVote = "female"
            
        Else
        
            If InStr(FrenchStringStart, "Né") > 0 Or InStr(FrenchStringStart, "Nato") > 0 Then
            
            FrenchVote = "male"
            
            Else
               
            FrenchVote = "No know gender markers!"
            
            End If
        
        End If


    ' collect the votes
    If FrenchVote = GermanVote Then
        
        GuessGender = FrenchVote
        
    Else
        
       If GermanVote = "No know gender markers!" And (FrenchVote = "female" Or FrenchVote = "male") Then
           
       GuessGender = FrenchVote
       
       Else
       
            If FrenchVote = "No know gender markers!" And (GermanVote = "female" Or GermanVote = "male") Then
    
            GuessGender = GermanVote
            
            Else
            
                GuessGender = "Can't figure it out!"
            
            End If
       
       End If
        
    End If
    

End Function

Function FindBothCodes(myrange As String)

  Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String

    strPattern = "(K |R |S |B |U |L |DE |T |V |P |C |R |LE |- )|(VW |EP |GW |EGC )"

    If strPattern <> "" Then
        strInput = myrange

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPattern
        End With

            If regEx.Test(strInput) Then
            
                   FindBothCodes = regEx.Execute(myrange)(0)
                   
            Else
                FindBothCodes = ""
            End If
       
    End If

End Function


Function GetListToExtractEntriesFrom(CompleteTextField As String)

Dim ReducedTextField As String

    'filter out of the matches you have already, what is left are the entries
    ReducedTextField = Replace(CompleteTextField, GetDateOfBirth(CompleteTextField), "")
    ReducedTextField = Replace(ReducedTextField, GetPlaceOfBirth(ReducedTextField), "", , 1) ' this one was problematic! often also matches part of educational enties e.t.c, we thus only want to replace the first occurence! this is what the , , 1) at the end specifies
    ReducedTextField = Replace(ReducedTextField, FindBothCodes(ReducedTextField), "") 'spaces at the start and end are added here to avoid that all 'u' for example are removed
    ReducedTextField = Replace(ReducedTextField, GetKantonCitizenShip(ReducedTextField), "")
    
    'some other phrases to get rid off
    ReducedTextField = Replace(ReducedTextField, "Geboren", "")
    ReducedTextField = Replace(ReducedTextField, "Bürger von", "")
    ReducedTextField = Replace(ReducedTextField, "Burger von", "")
    
    'trim and get rid dots and other characters at the start
    
    Do While Left(ReducedTextField, 1) = " " Or Left(ReducedTextField, 1) = "." Or Left(ReducedTextField, 1) = "(" Or Left(ReducedTextField, 1) = ")" Or Left(ReducedTextField, 1) = "," 'remove the first character as long as one of them is in this list
        ReducedTextField = Right(ReducedTextField, Len(ReducedTextField) - 1)
    Loop
    
    GetListToExtractEntriesFrom = Trim(ReducedTextField)

End Function

Function ReplaceTitleDots(text As String) As String

'function taken from http://stackoverflow.com/questions/18850981/excel-user-defined-formula-to-eliminate-special-characters

text = Replace(text, "Dr. Ir. ", "Dr** Ir**")
text = Replace(text, "Dr. med. ", "Dr** med**")
text = Replace(text, "Dr. h.c. ", "Dr** h**c** ")
text = Replace(text, "Dr. biol.", "Dr** biol**")
text = Replace(text, "Dr. cult.", "Dr** cult**")
text = Replace(text, "St.", "St**")
text = Replace(text, "Dr.", "Dr**")
text = Replace(text, "iur.", "iur**")
text = Replace(text, "Lie.", "Lie**")
text = Replace(text, "rer.", "rer**")
text = Replace(text, "oec.", "oec**")
text = Replace(text, "Agr.", "Agr**")
text = Replace(text, "Ing.", "Ing**")
    
ReplaceTitleDots = text

End Function



Function GetEntry(myrange As Range, NR As Integer) As String
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String
    
    'clean up the range
    Dim CleanedUpRange As String
    CleanedUpRange = myrange.Value

    'regexy bit
    strPattern = ".+?\." 'only split when longer then 4, otherwise it is probably an academic title.

    If strPattern <> "" Then
        strInput = CleanedUpRange

        With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPattern
        End With

        If NR < regEx.Execute(myrange).Count Then
        
            If regEx.Test(strInput) Then
            
                   GetEntry = Trim(regEx.Execute(myrange)(NR))
                   
            Else
                GetEntry = ""
            End If
        Else
            GetEntry = ""
        End If
        
    End If
End Function

Function CheckNameFormat(myrange As Range)
    
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim nrofmatches As Integer
    Dim index As Integer
    Dim ErrorList As String
    
    strPattern = "[a-z]"
    'I am a new line just for trying
    'I am a second new line just for trying
    'I am a third new line just for trying
    'I am a fourth new line just for trying
    'I am a fifth new line and I would love to end up in github online!
    'Oliver is watching, scary shit!

    With regEx
            .Global = True
            .MultiLine = False
            .IgnoreCase = True
            .Pattern = strPattern
        End With

    'do this for all entries! get the line of the ones that are being a      bitch

    index = 0
    Do Until index = myrange.Count

        'get count of numbers of characters that meet this criteria
            nrofmatches = regEx.Execute(myrange.Item(index + 1)).Count
        'get total number of characters
            entrylength = Len(myrange.Item(index + 1))
        
        If Not nrofmatches = entrylength Then
            ErrorList = ErrorList + CStr(myrange.Row + index) + ","
        End If
    
    index = index + 1
    Loop
    
    If Len(ErrorList) = 0 Then
        CheckNameFormat = "it all checks out"
    Else
        CheckNameFormat = "Integrity issues!! on line(s): " + ErrorList
    End If
    
End Function

