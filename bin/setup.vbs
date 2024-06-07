Class Setup
    ' Checks if a config.inf exists and creates one if no config.inf exists, otherwise it will first delete the old config
    Sub setConfig(homepath, port, ssl, language)
        If objFSO.FileExists("./bin/config.inf") Then
            objFSO.DeleteFile("./bin/config.inf") 
        End If
        
        Dim objFile
        Set objFile = objFSO.CreateTextFile("./bin/config.inf")
        objFile.Write "homepath:" + homepath + vbCrLf + "port:" + port + vbCrLf + "ssl:" + ssl + vbCrLf + "language:" + language
        objFile.Close
    End Sub

    Public Function getConfig()
        ' Create a dictionary to store configuration values
        Dim config: Set config = CreateObject("Scripting.Dictionary")
        Dim file

        ' Check if the config file exists and opens it, when it exists
        If objFSO.FileExists("./bin/config.inf") Then
            Set file = objFSO.OpenTextFile("./bin/config.inf")
        Else 
            ' If it doesn't exist, try to create a default config file and opens it
            setConfig "./www", "80", "no", "en"
            Set file = objFSO.OpenTextFile("./bin/config.inf")
        End if

        ' Read each line of the config file
        Do While Not file.AtEndOfStream
            line = file.ReadLine()
            ' Look for the colon separator
            index = InStr(line, ":")
            If index > 0 Then
                ' Extract the key and value from the line
                key = Trim(Left(line, index - 1))
                value = Trim(Mid(line, index + 1))
                config.add key,value
            End If
        Loop

        file.Close

        ' Set default values for any missing configuration values
        If (config.item("homepath") = "") Then
            config.add "homepath", "./www"
        End if

        If (config.item("port") = "") Then
            config.add "port", "80"
        End if

        If (config.item("ssl") = "") Then
            config.add "ssl", "no"
        End if

        If (config.item("language") = "") Then
            config.add "language", "en"
        End if

        ' Return the config dictionary
        Set getConfig = config
    End Function

    ' Checks if necessary files for HTTP and HTTPs Connections are installed in the bin folder
    Sub getServerExes()
        If (getFileSize("./bin/TinyWeb.exe") < 75000) Then
            downloadFile "https://www.marzeck.de/src/public/downloads/servervbs/", "TinyWeb.exe"
        End If
        If (getFileSize("./bin/TinySSL.exe") < 95000) Then
            downloadFile "https://www.marzeck.de/src/public/downloads/servervbs/", "TinySSL.exe"
        End If
        If (getFileSize("./bin/libeay32.dll") < 1330000) Then
            downloadFile "https://www.marzeck.de/src/public/downloads/servervbs/", "libeay32.dll"
        End If
        If (getFileSize("./bin/libssl32.dll") < 260000) Then
            downloadFile "https://www.marzeck.de/src/public/downloads/servervbs/", "libssl32.dll"
        End If
    End Sub

    Function getFileSize(file)
        If (objFSO.FileExists(file) = false) Then
            getFileSize = 0
        Else
            getFileSize = objFSO.GetFile(file).size
        End If
    End Function

    Sub downloadFile(url, filename)
        WScript.Echo objMain.strings.item("18") & filename & objMain.strings.item("19") & vbCrLf & objMain.strings.item("20") & vbCrLf & objMain.strings.item("21")
        Dim allowInstall: allowInstall = LCase(WScript.StdIn.ReadLine)
        If (allowInstall = objMain.strings.item("22") Or allowInstall = objMain.strings.item("23")) Then
            Dim http: Set http = CreateObject("MSXML2.XMLHTTP")
            Dim bStrm: Set bStrm = createobject("Adodb.Stream")

            ' Send a GET request to the URL to download the file
            http.open "GET", (url & filename), False
            http.send

            ' Check the response status code to ensure the request was successful
            If http.status = 200 Then
                Dim savePath: savePath = objFSO.GetAbsolutePathName("./bin/") &  "\" & filename

                with bStrm
                    .type = 1
                    .open
                    .write http.responseBody
                    .savetofile savepath, 2
                end with

                ' Prints out a message, when the file was successfully downloaded
                WScript.Echo objMain.strings.item("24") & savePath & vbCrLf
            Else
                ' Prints out Text if the File could not be downloaded
                WScript.Echo objMain.strings.item("25") & http.status & ")" & vbCrLf
            End If
        End If
    End Sub

    Sub addVbsToPathtext()
        ' Get the value of the PATHTEXT environment variable
        Dim envVar: envVar = CreateObject("WScript.Shell").Environment("SYSTEM").Item("PATHEXT")
        ' Check if the path to the script directory is already included in the PATHTEXT variable
        If (InStr(envVar, ".VBS") = 0) Then
            ' Add the path to the script directory to the PATHTEXT variable
            objShell.Environment("SYSTEM").Item("PATHEXT") = envVar & ";" & ".VBS"
        End If
    End Sub
End Class
