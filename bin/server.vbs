Class Server
    Private objShell
    Private objTinyWeb
    Private homePath
    Private strTinyWebPath
    
    Public Function startServer(homepath, port, ssl)
        objMain.showDivisionLine()
        
        ' Checks if the TinyWeb.exe and the TinySSL.exe are existing
        If (objFSO.FileExists("./bin/TinyWeb.exe") And objFSO.FileExists("./bin/TinySSL.exe")) Then
            WScript.Echo objMain.strings.item("server_start_port") & port & objMain.strings.item("server_started")
            Dim serverValues
            homePath = checkRootLocation(homePath)
            
            ' Building the Homepath
            serverValues = " " & chr(34) & checkRootLocation(homePath) & chr(34) & " " & chr(34) & port & chr(34)
            
            ' Checks if all necessary SSL Files are ready to start SSL-connection, otherwise start normal connection (HTTP)
            strTinyWebPath = "./bin/TinyWeb.exe"
            If (ssl = "yes") Then
                If (checkSslFiles) Then
                    strTinyWebPath = "./bin/TinySSL.exe"
                End If
            End If

            Set objShell = CreateObject("WScript.Shell")

            ' Check if the TinyWeb / TinySSL file exists
            If (objFSO.FileExists(strTinyWebPath) = False) Then
                WScript.Echo objMain.strings.item("no_executable") & strTinyWebPath
                WScript.Quit
            Else
                ' Use the shell object to start the TinyWeb process with the -c option to specify the config file
                Set objTinyWeb = objShell.Exec(strTinyWebPath & serverValues)
                WScript.Echo objMain.strings.item("set_start_directory") & homePath

                ' Wait a few seconds for the server to start up
                If (strTinyWebPath = "./bin/TinySSL.exe") Then
                    WScript.Echo objMain.strings.item("https_server_started")
                Else
                    WScript.Echo objMain.strings.item("http_server_started")
                End If
                'Return that the Server could be started
                startServer = true
            End If
        Else
            ' Prompt that the TinyWeb.exe was not found in the bin folder
            WScript.Echo objMain.strings.item("tinyweb_not_found") & vbCrLf & objMain.strings.item("tinyweb_needed")
            'Return that the Server could not be started
            startServer = false
        End If
    End Function

    Public Sub stopServer()
        ' Get the process ID of the TinyWeb executable
        Dim intProcessID: intProcessID = objTinyWeb.ProcessID

        ' Use the shell object to kill the TinyWeb process
        objShell.Run "taskkill /F /PID " & intProcessID, 0, True

        ' Cleanup
        Set objTinyWeb = Nothing
        Set objShell = Nothing
    End Sub 

    Public Function checkRootLocation(rootPath)
        ' Checks the root-location, when the rootPath is relative, it gets converted to absolute
        checkRootLocation = rootPath
        if(Left(rootPath, 1) = ".") Then
            checkRootLocation = objFSO.GetAbsolutePathName(rootPath)
        End If
    End Function

    Public Function checkSslFiles()
        ' Checks if all necessary SSL-Files are in the bin folder 
        checkSslFiles = True
        Dim files: files = Array("./bin/cert.pem", "./bin/key.pem", "./bin/realms.cfg", "./bin/libeay32.dll", "./bin/libssl32.dll")
        For Each file In files
            If (checkSslFile(file) = False) Then
                checkSslFiles = False
            End if
        Next
    End Function

    ' Checks, if the files needed for HTTPS mode are available locally
    Public Function checkSslFile(filename)
        checkSslFile = True
        If (objFSO.FileExists(filename) = false) Then
            checkSslFile = False
            WScript.Echo objMain.strings.item("file_not_found") & filename
        End if
    End Function
End Class