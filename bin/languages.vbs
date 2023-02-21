Class Language
    Function initLang(language)
        'Initializes the given language
        Select Case LCase(language)
            Case "de"
                Set initLang = initGerman
            Case Else
                Set initLang = initEnglish
        End Select
    End Function

    Function initGerman()
        'Initalizes the german language
        Dim strings: Set strings = CreateObject("Scripting.Dictionary")
        strings.add "0", "starten"
        strings.add "1", "pausieren"
        strings.add "2", "beenden"
        strings.add "3", "neustarten"
        strings.add "4", "Unbekannter Befehl."
        strings.add "5", "starten    - Startet den Server, nachdem dieser pausiert wurde."
        strings.add "6", "pausieren  - Pausiert den Server, schlie" & ChrW(223) & "t aber nicht das Programm."
        strings.add "7", "beenden    - F" & ChrW(228) & "hrt den Server herunter und beendet das Programm."
        strings.add "8", "neustarten - Startet den Server neu."
        strings.add "9", "Server wird auf Port "
        strings.add "10"," gestartet..."
        strings.add "11","Keine ausf" & ChrW(252) & "hrbare TinyWeb-Datei gefunden, unter: "
        strings.add "12","Setze den Heimatpfad auf: "
        strings.add "13","Server (HTTPS) wurde erfolgreich gestartet..."
        strings.add "14","Server (HTTP) wurde erfolgreich gestartet..."
        strings.add "15","Die Datei 'TinyWeb.exe' konnte im Ordner 'bin' nicht gefunden werden."
        strings.add "16","Die Datei ist jedoch wichtig f" & ChrW(252) & "r die Funktion des Servers."
        strings.add "17","Folgende Datei konnte nicht gefunden werden: "
        strings.add "18","Um den Server zu starten wird die Datei '"
        strings.add "19","' ben" & ChrW(246) & "tigt. "
        strings.add "20","Eine Internetverbindung f" & ChrW(252) & "r das Herunterladen gebraucht."
        strings.add "21","M" & ChrW(246) & "chten Sie den Download jetzt starten? ([J] Ja | [N] Nein)"
        strings.add "22","j"
        strings.add "23","ja"
        strings.add "24","Datei wurde erfolgreich heruntergeladen: "
        strings.add "25","Fehler beim Herunterladen der Datei (HTTP-Status "
        Set initGerman = strings
    End Function

    Function initEnglish()
        'Initializes the english language (standard)
        Dim strings: Set strings = CreateObject("Scripting.Dictionary")
        strings.add "0", "start"
        strings.add "1", "pause"
        strings.add "2", "stop"
        strings.add "3", "restart"
        strings.add "4", "Unknown command."
        strings.add "5", "start   - Starts the server, after it was paused."
        strings.add "6", "pause   - Stops the server, but does not close the program."
        strings.add "7", "stop    - Shutsdown the server and stops the program."
        strings.add "8", "restart - Restarts the server."
        strings.add "9", "Starting the Server on Port "
        strings.add "10", "..."
        strings.add "11", "TinyWeb executable not found at "
        strings.add "12", "Set homepath to: "
        strings.add "13","Server (HTTPS) has been started..."
        strings.add "14","Server (HTTP) has been started..."
        strings.add "15","The file 'TinyWeb.exe' could not be found in the folder 'bin'."
        strings.add "16","But the file is necessary to start the server."
        strings.add "17","Could not find the file: "
        strings.add "18","To start the server the file '"
        strings.add "19","' is needed. "
        strings.add "20","An internet connection is required for the download."
        strings.add "21","Do you want to start the download now? ([Y] Yes | [N] No)"
        strings.add "22","y"
        strings.add "23","yes"
        strings.add "24","File downloaded successfully to "
        strings.add "25","Error downloading file (HTTP status "
        Set initEnglish = strings
    End Function
End class