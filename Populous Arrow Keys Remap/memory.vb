Imports System.Text
Imports System.Threading

Namespace PopMem
    Public Class PopulousMemory

        Public proc As PopProcess
        Public popRenderer As Integer = 0
        Public gameStarted As Boolean = False
        Public Function isRunning()
            If Not Me.proc.proc.HasExited Then
                Return True
            End If
            Return False
        End Function
        Enum Movement
            UP = &H686808
            DOWN = &H686810
            ROTATE_LEFT = &H68680D
            ROTATE_RIGHT = &H68680B
        End Enum
        Public Sub Movement_Key_Down(ByVal key As Movement)
            If key = Movement.UP Then
                proc.WriteByte(CType(Movement.UP, IntPtr), 1)
            ElseIf key = Movement.DOWN Then
                proc.WriteByte(CType(Movement.DOWN, IntPtr), 1)
            ElseIf key = Movement.ROTATE_LEFT Then
                proc.WriteByte(CType(Movement.ROTATE_RIGHT, IntPtr), 1)
            ElseIf key = Movement.ROTATE_RIGHT Then
                proc.WriteByte(CType(Movement.ROTATE_LEFT, IntPtr), 1)
            End If
        End Sub
        Public Sub Movement_Key_Up(ByVal key As Movement)
            If key = Movement.UP Then
                proc.WriteByte(CType(Movement.UP, IntPtr), 0)
            ElseIf key = Movement.DOWN Then
                proc.WriteByte(CType(Movement.DOWN, IntPtr), 0)
            ElseIf key = Movement.ROTATE_LEFT Then
                proc.WriteByte(CType(Movement.ROTATE_RIGHT, IntPtr), 0)
            ElseIf key = Movement.ROTATE_RIGHT Then
                proc.WriteByte(CType(Movement.ROTATE_LEFT, IntPtr), 0)
            End If
        End Sub
        Public Sub Change_Keys()
            ' SOFTWARE ONLY
            proc.WriteByte(CType(&H8853FA, IntPtr), 20)
            proc.WriteByte(CType(&H877598, IntPtr), 11)
            MsgBox("Don't close this dialog (don't press ""Ok"") until you're finished mapping keys. Go to the game. Press F6 to save the key mappings.")
            proc.WriteByte(CType(&H8853FA, IntPtr), &HFF)
        End Sub
        Public Sub New(ByVal TheProcess As PopProcess, ByVal TheRenderer As Integer)
            proc = TheProcess
            popRenderer = TheRenderer
        End Sub
    End Class
End Namespace
