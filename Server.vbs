' ServerVBS by Maximilian Marzeck (2023)
' TinyWeb by Maxim Masiutin (1997 - 2023)

Dim objFS0: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objMain: Set objMain = New Main

objMain.includeFile "./bin/server.vbs"
objMain.includeFile "./bin/setup.vbs"
objMain.includeFile "./bin/languages.vbs"

Class Main
    Dim mySetup
    Dim myServer 
    Dim myConfig
    Dim serverMode
    Dim strings

    ' Initializes the Setup (with Config), the server and the language pack
    Sub main()
        forceConsoleMode

        Set mySetup = New Setup
        Set myServer = New Server
        Set myConfig = mySetup.getConfig()
        
        Set strings = New Language.initLang(myConfig.item("language"))
        serverMode = 0

        ' Checks if the necessary exe and dll files are in their right place
        mySetup.getServerExes
        mySetup.addVbsToPathtext

        ' Starts the server and prints the menu
        If myServer.startServer(myConfig.item("homepath"), myConfig.item("port"), myConfig.item("ssl")) Then
            Dim mainMenu:cmainMenu = ""

            serverMode = 1

            showDivisionLine

            ' Checks which command was called by the user
            Do While Not mainMenu = strings.item("2")
                printMenu
                mainMenu = LCase(WScript.StdIn.ReadLine)

                Select Case mainMenu
                    Case strings.item("0")
                        serverStart
                    Case strings.item("1")
                        serverStop
                    Case strings.item("2")
                        serverStop
                    Case strings.item("3")
                        serverStop
                        serverStart
                    Case strings.item("26")
                        showSettings
                    Case Else
                        WScript.Echo strings.item("4")
                End Select

                showDivisionLine
            Loop
        End If
    End Sub

    ' Includes a vbs file and executes the code
    Public Sub includeFile(file)
        With CreateObject("Scripting.FileSystemObject")
            executeGlobal .openTextFile(file).readAll()
        End With
    End Sub

    ' Creates a string by concatenating the str by the count times
    Function stringRepeat(str, count)
        stringRepeat = String(count, str)
    End Function

    ' Shows a division line in the command prompt
    Public Sub showDivisionLine()
        WScript.Echo vbCrLf & stringRepeat("#", 70) & vbCrLf
    End Sub

    ' Prints the menu for the user
    Public Sub printMenu()
        WScript.Echo strings.item("5") & vbCrLf & _
            strings.item("6") & vbCrLf & _
            strings.item("7") & vbCrLf & _
            strings.item("8") & vbCrLf & _
            strings.item("39") & vbCrLf
    End Sub

    ' Starting all the necessary steps to start the server if it is not started yet
    Public Sub serverStart()
        If serverMode = 0 Then
            Set myConfig = mySetup.getConfig()
            Set strings = New Language.initLang(myConfig.item("language"))
            myServer.startServer myConfig.item("homepath"), myConfig.item("port"), myConfig.item("ssl")
            serverMode = 1
        End If
    End Sub

    ' Starting all the necessary steps to stop the server if it is not stopped yet
    Public Sub serverStop()
        If serverMode = 1 Then
            myServer.stopServer
            serverMode = 0
        End If
    End Sub

    'Shows the user the possible settinsg for the server
    public Sub showSettings()
        showDivisionLine
        printSettings

        Select case LCase(WScript.StdIn.ReadLine)
            case strings.item("31")
                setSetting "35", strings.item("31")
            case strings.item("32")
                setSetting "36", strings.item("32")
            case strings.item("33")
                setSetting "37", strings.item("33")
            case strings.item("34")
                setSetting "38", strings.item("34")
            Case Else
                WScript.Echo strings.item("4")
        End Select
    End Sub

    ' Prints out the options for the settings
    public Sub printSettings()
        WScript.Echo strings.item("27") & vbCrLf & _
            strings.item("28") & vbCrLf & _
            strings.item("29") & vbCrLf & _
            strings.item("30") & vbCrLf 
    End Sub

    ' Sets the setting in the config file
    Public Sub setSetting(text_id, setting)
        WScript.Echo strings.item(text_id)

        Dim setting_content
        setting_content = LCase(WScript.StdIn.ReadLine)

        Select case setting
            case strings.item("31") ' Set_Homepath
                mySetup.setConfig setting_content, myConfig.item("port"), myConfig.item("ssl"), myConfig.item("language")
            case strings.item("32") ' Set_Port
                If IsNumeric(setting_content) Then
                    mySetup.setConfig myConfig.item("homepath"), setting_content, myConfig.item("ssl"), myConfig.item("language")
                End If
            case strings.item("33") ' Set_SSL
                If setting_content = strings.item("22") Or setting_content = strings.item("23") Then
                    mySetup.setConfig myConfig.item("homepath"), myConfig.item("port"), "yes", myConfig.item("language")
                Else
                    mySetup.setConfig myConfig.item("homepath"), myConfig.item("port"), "no", myConfig.item("language")
                End If
            case strings.item("34") ' Set_Language
                If New Language.languageExists(setting_content) Then
                    mySetup.setConfig myConfig.item("homepath"), myConfig.item("port"), myConfig.item("ssl"), setting_content
                End If
        End Select

        serverStop
        serverStart
    End Sub

    ' Forces the application to run in CScript.exe
    Public Sub forceConsoleMode()
        Dim strArgs, strCmd, strEngine, i, objDebug, wshShell

        Set wshShell = CreateObject( "WScript.Shell" )
        strEngine = UCase( Right( WScript.FullName, 12 ) )

        If strEngine <> "\CSCRIPT.EXE" Then
            strArgs = ""
            
            If WScript.Arguments.Count > 0 Then
                For i = 0 To WScript.Arguments.Count - 1
                    strArgs = strArgs & " " & WScript.Arguments(i)
                Next
            End If

            strCmd = "CSCRIPT.EXE //NoLogo """ & WScript.ScriptFullName & """" & strArgs
            Set objDebug = wshShell.Exec( strCmd )

            Do While objDebug.Status = 0
                WScript.Sleep 100
            Loop

            WScript.Quit objDebug.ExitCode
        End If
    End Sub
End Class

'Starting the main function
objMain.main