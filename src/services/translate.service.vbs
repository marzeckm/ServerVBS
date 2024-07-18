Class TranslateService
    Public Function initLang(language)
        'Initializes the given language
        Select Case LCase(language)
            Case "de"
                Set initLang = initStrings("de")
            Case Else
                Set initLang = initStrings("en")
        End Select
    End Function

    Public Function languageExists(language)
        if (language = "de" Or language = "Deutsch") Then
            languageExists = 1 'TRUE
        ElseIf (language = "en" Or language = "English") Then
            languageExists = 1 'TRUE
        Else
            languageExists = 0 'FALSE
        End if
    End Function

    ' Reads and returns the strings
    Function initStrings(lang)
        Dim translation: Set translation = CreateObject("Scripting.Dictionary")
        
        Dim strings
        If (lang = "de") Then
            strings = translationDe
        Else
            strings = translationEn
        End If

        Dim i
        For i = 0 To UBound(strings)
            translation.Add strings(i)(0), strings(i)(1)
        Next

        Set initStrings = translation
    End Function

End class
