Imports PMP.PopMem
Imports Gma.UserActivityMonitor
Imports System.Windows.Forms
Imports System.Environment
Imports System.Configuration
Imports System.Net.Configuration

Module main
    Dim popTB As PopProcess
    Dim memory As PopulousMemory
    Dim actHook As GlobalEventProvider = New GlobalEventProvider
    Dim isRunning As Boolean = False
    Dim customKeys As Boolean = False
    Dim populousInstall As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Bullfrog Productions Ltd\Populous: The Beginning", "InstallPath", "")

    Private Class Controls
        Public UP As Integer = 87
        Public DOWN As Integer = 83
        Public ROTATE_LEFT As Integer = 65
        Public ROTATE_RIGHT As Integer = 68
    End Class
    Dim GameControls As New Controls

    Sub Main()
        If Not IsNothing(ReadSetting("UP")) Then
            GameControls.UP = ReadSetting("UP")
        End If
        If Not IsNothing(ReadSetting("DOWN")) Then
            GameControls.DOWN = ReadSetting("DOWN")
        End If
        If Not IsNothing(ReadSetting("ROTATE_LEFT")) Then
            GameControls.ROTATE_LEFT = ReadSetting("ROTATE_LEFT")
        End If
        If Not IsNothing(ReadSetting("ROTATE_RIGHT")) Then
            GameControls.ROTATE_RIGHT = ReadSetting("ROTATE_RIGHT")
        End If

        If Not IsNothing(ReadSetting("CUSTOM_KEYS")) Then
            customKeys = ReadSetting("CUSTOM_KEYS")
        End If

        If Not customKeys Then
            If Not My.Computer.FileSystem.DirectoryExists(populousInstall & "\SAVE\") Then
                My.Computer.FileSystem.CreateDirectory(populousInstall & "\SAVE\")
            End If
            My.Computer.FileSystem.WriteAllBytes(populousInstall & "\SAVE\key_def.dat", My.Resources.key_def, False)

            Dim appData As String = GetFolderPath(SpecialFolder.LocalApplicationData)

            If Not My.Computer.FileSystem.DirectoryExists(appData & "\VirtualStore\Program Files (x86)\Bullfrog\Populous\.SAVE\") Then
                My.Computer.FileSystem.CreateDirectory(appData & "\VirtualStore\Program Files (x86)\Bullfrog\Populous\.SAVE\")
            End If
            My.Computer.FileSystem.WriteAllBytes(appData & "\VirtualStore\Program Files (x86)\Bullfrog\Populous\.SAVE\key_def.dat", My.Resources.key_def, False)

            If Not My.Computer.FileSystem.DirectoryExists(appData & "\Populous Reincarnated\Populous\SAVE\") Then
                My.Computer.FileSystem.CreateDirectory(appData & "\Populous Reincarnated\Populous\SAVE\")
            End If
            My.Computer.FileSystem.WriteAllBytes(appData & "\Populous Reincarnated\Populous\SAVE\key_def.dat", My.Resources.key_def, False)
        End If

        AddHandler actHook.KeyDown, AddressOf KeyDown
        AddHandler actHook.KeyUp, AddressOf KeyUp

        Dim status As Boolean = True
        Console.Title = "Populous Memory Patcher"
        ConsoleTools.SetConsoleColors(ConsoleColor.aqua)
        Console.WriteLine("Populous Memory Patcher (KEY MAP EDITION) By Brandan Tyler Lasley 2015" & vbNewLine)
        ConsoleTools.SetConsoleColors(ConsoleColor.white)

        Do Until Nothing
            Try
                Application.DoEvents()
                If status Then
                    ConsoleTools.SetConsoleColors(ConsoleColor.white)
                    Console.WriteLine(Date.Now & " Searching For Populous (SOFTWARE)..." & vbNewLine)
                    status = False
                End If
                If Not isRunning Then
                    Dim processesByName As Process() = Process.GetProcessesByName("popTB")
                    If Not (processesByName.Length < 1) Then
                        ConsoleTools.SetConsoleColors(ConsoleColor.green)
                        Console.WriteLine(Date.Now & " Populous Found, patching!" & vbNewLine)
                        ConsoleTools.SetConsoleColors(ConsoleColor.white)
                        popTB = New PopProcess(processesByName(0))
                        memory = New PopulousMemory(popTB, 1)
                        isRunning = True

                    End If
                Else
                    If Not IsNothing(memory) Then
                        If Not memory.isRunning Then
                            ConsoleTools.SetConsoleColors(ConsoleColor.red)
                            Console.WriteLine(Date.Now & " Populous Exited :(" & vbNewLine)
                            status = True
                            isRunning = False
                            memory = Nothing
                            popTB = Nothing
                        End If
                    End If
                End If
                System.Threading.Thread.Sleep(1)
            Catch ex As Exception
                ConsoleTools.SetConsoleColors(ConsoleColor.red)
                Console.WriteLine(Date.Now & " Error: " & ex.Message & vbNewLine)
            End Try
        Loop
    End Sub
    Private Function ReadSetting(key As String)
        Try
            Dim appSettings = ConfigurationManager.AppSettings
            Dim result As String = appSettings(key)
            If IsNothing(result) Then
                Return Nothing
            End If
            Return (result)
        Catch e As ConfigurationErrorsException
            Return Nothing
        End Try
    End Function
    Private Sub AddUpdateAppSettings(key As String, value As String)
        Try
            Dim configFile = ConfigurationManager.OpenExeConfiguration(My.Application.Info.AssemblyName + ".exe")
            Dim settings = configFile.AppSettings.Settings
            If IsNothing(settings(key)) Then
                settings.Add(key, value)
            Else
                settings(key).Value = value
            End If
            configFile.Save(ConfigurationSaveMode.Modified)
            ConfigurationManager.RefreshSection(My.Application.Info.AssemblyName + ".exe")
        Catch e As ConfigurationErrorsException
            Console.WriteLine("Error writing app settings")
        End Try
    End Sub
    Private Sub KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If isRunning Then
            If memory.isRunning Then
                If e.KeyCode = GameControls.UP Then ' UP
                    Console.Write("UP" & vbNewLine)
                    memory.Movement_Key_Down(PopulousMemory.Movement.UP)
                ElseIf e.KeyCode = GameControls.DOWN Then ' DOWN
                    Console.Write("DOWN" & vbNewLine)
                    memory.Movement_Key_Down(PopulousMemory.Movement.DOWN)
                ElseIf e.KeyCode = GameControls.ROTATE_LEFT Then ' LEFT
                    Console.Write("LEFT" & vbNewLine)
                    memory.Movement_Key_Down(PopulousMemory.Movement.ROTATE_LEFT)
                ElseIf e.KeyCode = GameControls.ROTATE_RIGHT Then ' RIGHT
                    Console.Write("RIGHT" & vbNewLine)
                    memory.Movement_Key_Down(PopulousMemory.Movement.ROTATE_RIGHT)
                End If
            End If
        End If
    End Sub
    Private Sub KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs)
        If isRunning Then
            If memory.isRunning Then
                If e.KeyCode = GameControls.UP Then ' UP
                    memory.Movement_Key_Up(PopulousMemory.Movement.UP)
                ElseIf e.KeyCode = GameControls.DOWN Then ' DOWN
                    memory.Movement_Key_Up(PopulousMemory.Movement.DOWN)
                ElseIf e.KeyCode = GameControls.ROTATE_LEFT Then ' LEFT
                    memory.Movement_Key_Up(PopulousMemory.Movement.ROTATE_LEFT)
                ElseIf e.KeyCode = GameControls.ROTATE_RIGHT Then ' RIGHT
                    memory.Movement_Key_Up(PopulousMemory.Movement.ROTATE_RIGHT)
                End If
            End If
        End If
    End Sub
End Module
