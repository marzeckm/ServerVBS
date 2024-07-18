' ServerVBS by Maximilian Marzeck (2023)
' TinyWeb by Maxim Masiutin (1997 - 2023)

Dim objFS0: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objMain: Set objMain = New Main

' Import the languages
objMain.includeFile "./src/constants/translation_en.vbs"
objMain.includeFile "./src/constants/translation_de.vbs"

' Import the needed classes
objMain.includeFile "./bin/server.vbs"
objMain.includeFile "./bin/setup.vbs"
objMain.includeFile "./src/services/translate.service.vbs"


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
        
        Set strings = New TranslateService.initLang(myConfig.item("language"))
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
            Do While Not mainMenu = strings.item("stop")
                printMenu
                mainMenu = LCase(WScript.StdIn.ReadLine)

                Select Case mainMenu
                    Case strings.item("start")
                        serverStart
                    Case strings.item("pause")
                        serverStop
                    Case strings.item("stop")
                        serverStop
                    Case strings.item("restart")
                        serverStop
                        serverStart
                    Case strings.item("settings")
                        showSettings
                    Case Else
                        WScript.Echo strings.item("unknown_command")
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
        WScript.Echo strings.item("start_description") & vbCrLf & _
            strings.item("pause_description") & vbCrLf & _
            strings.item("stop_description") & vbCrLf & _
            strings.item("restart_description") & vbCrLf & _
            strings.item("settings_description") & vbCrLf
    End Sub

    ' Starting all the necessary steps to start the server if it is not started yet
    Public Sub serverStart()
        If serverMode = 0 Then
            Set myConfig = mySetup.getConfig()
            Set strings = New TranslateService.initLang(myConfig.item("language"))
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
            case strings.item("start_directory")
                setSetting "set_start_directory_prompt", strings.item("start_directory")
            case strings.item("port")
                setSetting "set_port_prompt", strings.item("port")
            case strings.item("ssl")
                setSetting "set_ssl_prompt", strings.item("ssl")
            case strings.item("language")
                setSetting "set_language_prompt", strings.item("language")
            Case Else
                WScript.Echo strings.item("unknown_command")
        End Select
    End Sub

    ' Prints out the options for the settings
    public Sub printSettings()
        WScript.Echo strings.item("set_start_directory_description") & vbCrLf & _
            strings.item("set_port_description") & vbCrLf & _
            strings.item("set_ssl_description") & vbCrLf & _
            strings.item("set_language_description") & vbCrLf 
    End Sub

    ' Sets the setting in the config file
    Public Sub setSetting(text_id, setting)
        WScript.Echo strings.item(text_id)

        Dim setting_content
        setting_content = LCase(WScript.StdIn.ReadLine)

        Select case setting
            case strings.item("start_directory") ' Set_Homepath
                mySetup.setConfig setting_content, myConfig.item("port"), myConfig.item("ssl"), myConfig.item("language")
            case strings.item("port") ' Set_Port
                If IsNumeric(setting_content) Then
                    mySetup.setConfig myConfig.item("homepath"), setting_content, myConfig.item("ssl"), myConfig.item("language")
                End If
            case strings.item("ssl") ' Set_SSL
                If setting_content = strings.item("yes_short") Or setting_content = strings.item("yes_long") Then
                    mySetup.setConfig myConfig.item("homepath"), myConfig.item("port"), "yes", myConfig.item("language")
                Else
                    mySetup.setConfig myConfig.item("homepath"), myConfig.item("port"), "no", myConfig.item("language")
                End If
            case strings.item("language") ' Set_Language
                If New TranslateService.languageExists(setting_content) Then
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