Imports System.Text
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Diagnostics

Namespace PopMem
    Public Class PopProcess

        <DllImport("kernel32.dll")> _
        Public Shared Function WriteProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, <[In](), Out()> ByVal buffer As Byte(), ByVal size As UInt32, ByRef lpNumberOfBytesWritten As IntPtr) As Integer
        End Function
        <DllImport("kernel32.dll")> _
        Public Shared Function ReadProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, <[In](), Out()> ByVal buffer As Byte(), ByVal size As UInteger, ByRef lpNumberOfBytesRead As IntPtr) As Integer
        End Function

        Public proc As Process
        Public Event OnExit As EventHandler
        Public Sub New(ByVal TheProcess As Process)
            proc = TheProcess
        End Sub
        Private Sub ExitWatcher()
            While True
                Try
                    Thread.Sleep(1000)

                    If proc.HasExited Then
                        RaiseEvent OnExit(Me, EventArgs.Empty)

                        Thread.CurrentThread.Abort()
                    End If

                Catch
                End Try
            End While
        End Sub
        Public Sub WriteInt8(ByVal address As IntPtr, ByVal val As Byte)
            Dim bytes As Byte() = New Byte() {val}
            WriteProcessMemory(address, bytes)
        End Sub
        Public Function ReadInt8(ByVal address As IntPtr) As [Byte]
            Return ReadProcessMemory(address, 1)(0)
        End Function
        Public Sub WriteInt16(ByVal address As IntPtr, ByVal val As Short)
            Dim bytes As Byte() = BitConverter.GetBytes(val)
            WriteProcessMemory(address, bytes)
        End Sub
        Public Function ReadInt16(ByVal address As IntPtr) As Short
            Return BitConverter.ToInt16(ReadProcessMemory(address, 2), 0)
        End Function
        Public Sub WriteInt32(ByVal address As IntPtr, ByVal val As Integer)
            Dim bytes As Byte() = BitConverter.GetBytes(val)
            WriteProcessMemory(address, bytes)
        End Sub
        Public Function ReadInt32(ByVal address As IntPtr) As Integer
            Return BitConverter.ToInt32(ReadProcessMemory(address, 4), 0)
        End Function
        Public Sub WriteString(ByVal address As IntPtr, ByVal enc As Encoding, ByVal str As String)
            Dim bytes As Byte() = enc.GetBytes(str)
            Me.WriteProcessMemory(address, bytes)
        End Sub
        Public Sub WriteString(ByVal address As IntPtr, ByVal str As String)
            Me.WriteString(address, Encoding.Unicode, str)
        End Sub
        Public Sub WriteShort(ByVal address As IntPtr, ByVal val As Short)
            Dim bytes As Byte() = BitConverter.GetBytes(val)
            Me.WriteProcessMemory(address, bytes)
        End Sub
        Public Sub WriteProcessMemory(ByVal address As IntPtr, ByVal ByteToWrite As Byte)
            Me.WriteProcessMemory(address, New Byte() {ByteToWrite})
        End Sub
        Public Function WriteProcessMemory(ByVal address As IntPtr, ByVal bytes As Byte()) As IntPtr
            Dim lpNumberOfBytesWritten As New IntPtr(0)
            PopProcess.WriteProcessMemory(Me.proc.Handle, address, bytes, CType(bytes.Length, UInt32), lpNumberOfBytesWritten)
            Return lpNumberOfBytesWritten
        End Function
        Public Sub WriteInt(ByVal address As IntPtr, ByVal val As Integer)
            Dim bytes As Byte() = BitConverter.GetBytes(val)
            Me.WriteProcessMemory(address, bytes)
        End Sub
        Public Sub WriteByte(ByVal address As IntPtr, ByVal val As Byte)
            Dim bytes As Byte() = New Byte() {val}
            Me.WriteProcessMemory(address, bytes)
        End Sub
        Public Function ReadString(ByVal address As IntPtr, ByVal length As UInt32) As String
            Return Me.ReadString(address, Encoding.Unicode, length)
        End Function
        Public Function ReadShort(ByVal address As IntPtr) As Short
            Return BitConverter.ToInt16(Me.ReadProcessMemory(address, 2), 0)
        End Function
        Public Function ReadProcessMemory(ByVal MemoryAddress As IntPtr, ByVal bytesToRead As UInt32) As Byte()
            Dim zero As IntPtr = IntPtr.Zero
            Dim buffer As Byte() = New Byte((CInt((bytesToRead - 1)) + 1) - 1) {}
            PopProcess.ReadProcessMemory(Me.proc.Handle, MemoryAddress, buffer, bytesToRead, zero)
            Return buffer
        End Function
        Public Function ReadInt(ByVal address As IntPtr) As Integer
            Return BitConverter.ToInt32(Me.ReadProcessMemory(address, 4), 0)
        End Function
        Public Function ReadByte(ByVal address As IntPtr) As Byte
            Return Me.ReadProcessMemory(address, 1)(0)
        End Function
        Public Function ReadString(ByVal address As IntPtr, ByVal enc As Encoding, ByVal MaxLength As UInteger) As String
            Return enc.GetString(ReadProcessMemory(address, MaxLength))
        End Function
    End Class
End Namespace
