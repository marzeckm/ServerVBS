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
            strings.item("8") & vbCrLf   
    End Sub

    ' Starting all the necessary steps to start the server if it is not started yet
    Public Sub serverStart()
        If serverMode = 0 Then
            Set myConfig = mySetup.getConfig()
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

    ' Forces the application to run in CScript.exe
    Public Sub forceConsoleMode()
        Dim strCmd, strEngine, wshShell

        Set wshShell = CreateObject( "WScript.Shell" )
        strEngine = UCase( Right( WScript.FullName, 12 ) )

        If strEngine <> "\CSCRIPT.EXE" Then

            strCmd = "cscript.exe " & chr(34) & WScript.ScriptFullName & chr(34) & strArgs
            wshShell.Run "cscript.exe Server.vbs"

            WScript.Quit
        End If
    End Sub
End Class

'Starting the main function
objMain.main