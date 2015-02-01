Module ConsoleTools

    Private Declare Function SetConsoleTextAttribute Lib "kernel32" _
      (ByVal hConsoleOutput As Integer, _
       ByVal wScreenColors As Integer) As Long

    Private Declare Function GetStdHandle Lib "kernel32" _
        (ByVal nStdHandle As Long) As Integer

    Private Const STD_OUTPUT_HANDLE As Integer = -11

    Public Enum ConsoleColor
        black = 0
        darkblue = 1
        darkgreen = 2
        darkaqua = 3
        darkred = 4
        purple = 5
        brown = 6
        grey = 7
        darkgrey = 8
        blue = 9
        green = 10
        aqua = 11
        red = 12
        pink = 13
        yellow = 14
        white = 15
    End Enum

    Public Sub SetConsoleColors(ByVal forecolor As ConsoleColor, ByVal backcolor As ConsoleColor)
        Try
            Dim hConsole As Integer = GetStdHandle(STD_OUTPUT_HANDLE)
            backcolor = IIf(backcolor = 0, 256, backcolor * 16)
            SetConsoleTextAttribute(hConsole, forecolor Or backcolor)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub SetConsoleColors(ByVal forecolor As ConsoleColor)
        Try
            Dim hConsole As Integer = GetStdHandle(STD_OUTPUT_HANDLE)
            SetConsoleTextAttribute(hConsole, forecolor)
        Catch ex As Exception
        End Try
    End Sub

    Public Function wordquest(ByVal splitz, ByVal text, ByVal number) As String
        On Error Resume Next
        Dim Words() As String = Split(text, splitz)
        Return Words(number)
        On Error Resume Next
    End Function
End Module