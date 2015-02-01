
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Forms
Imports System.ComponentModel

Namespace gma.System.Windows
	''' <summary>
	''' This class allows you to tap keyboard and mouse and / or to detect their activity even when an 
	''' application runes in background or does not have any user interface at all. This class raises 
	''' common .NET events with KeyEventArgs and MouseEventArgs so you can easily retrive any information you need.
	''' </summary>
	Public Class UserActivityHook
		#Region "Windows structure definitions"

		''' <summary>
		''' The POINT structure defines the x- and y- coordinates of a point. 
		''' </summary>
		''' <remarks>
		''' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdi/rectangl_0tiq.asp
		''' </remarks>
		<StructLayout(LayoutKind.Sequential)> _
		Private Class POINT
			''' <summary>
			''' Specifies the x-coordinate of the point. 
			''' </summary>
			Public x As Integer
			''' <summary>
			''' Specifies the y-coordinate of the point. 
			''' </summary>
			Public y As Integer
		End Class

		''' <summary>
		''' The MOUSEHOOKSTRUCT structure contains information about a mouse event passed to a WH_MOUSE hook procedure, MouseProc. 
		''' </summary>
		''' <remarks>
		''' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookstructures/cwpstruct.asp
		''' </remarks>
		<StructLayout(LayoutKind.Sequential)> _
		Private Class MouseHookStruct
			''' <summary>
			''' Specifies a POINT structure that contains the x- and y-coordinates of the cursor, in screen coordinates. 
			''' </summary>
			Public pt As POINT
			''' <summary>
			''' Handle to the window that will receive the mouse message corresponding to the mouse event. 
			''' </summary>
			Public hwnd As Integer
			''' <summary>
			''' Specifies the hit-test value. For a list of hit-test values, see the description of the WM_NCHITTEST message. 
			''' </summary>
			Public wHitTestCode As Integer
			''' <summary>
			''' Specifies extra information associated with the message. 
			''' </summary>
			Public dwExtraInfo As Integer
		End Class

		''' <summary>
		''' The MSLLHOOKSTRUCT structure contains information about a low-level keyboard input event. 
		''' </summary>
		<StructLayout(LayoutKind.Sequential)> _
		Private Class MouseLLHookStruct
			''' <summary>
			''' Specifies a POINT structure that contains the x- and y-coordinates of the cursor, in screen coordinates. 
			''' </summary>
			Public pt As POINT
			''' <summary>
			''' If the message is WM_MOUSEWHEEL, the high-order word of this member is the wheel delta. 
			''' The low-order word is reserved. A positive value indicates that the wheel was rotated forward, 
			''' away from the user; a negative value indicates that the wheel was rotated backward, toward the user. 
			''' One wheel click is defined as WHEEL_DELTA, which is 120. 
			'''If the message is WM_XBUTTONDOWN, WM_XBUTTONUP, WM_XBUTTONDBLCLK, WM_NCXBUTTONDOWN, WM_NCXBUTTONUP,
			''' or WM_NCXBUTTONDBLCLK, the high-order word specifies which X button was pressed or released, 
			''' and the low-order word is reserved. This value can be one or more of the following values. Otherwise, mouseData is not used. 
			'''XBUTTON1
			'''The first X button was pressed or released.
			'''XBUTTON2
			'''The second X button was pressed or released.
			''' </summary>
			Public mouseData As Integer
			''' <summary>
			''' Specifies the event-injected flag. An application can use the following value to test the mouse flags. Value Purpose 
			'''LLMHF_INJECTED Test the event-injected flag.  
			'''0
			'''Specifies whether the event was injected. The value is 1 if the event was injected; otherwise, it is 0.
			'''1-15
			'''Reserved.
			''' </summary>
			Public flags As Integer
			''' <summary>
			''' Specifies the time stamp for this message.
			''' </summary>
			Public time As Integer
			''' <summary>
			''' Specifies extra information associated with the message. 
			''' </summary>
			Public dwExtraInfo As Integer
		End Class


		''' <summary>
		''' The KBDLLHOOKSTRUCT structure contains information about a low-level keyboard input event. 
		''' </summary>
		''' <remarks>
		''' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookstructures/cwpstruct.asp
		''' </remarks>
		<StructLayout(LayoutKind.Sequential)> _
		Private Class KeyboardHookStruct
			''' <summary>
			''' Specifies a virtual-key code. The code must be a value in the range 1 to 254. 
			''' </summary>
			Public vkCode As Integer
			''' <summary>
			''' Specifies a hardware scan code for the key. 
			''' </summary>
			Public scanCode As Integer
			''' <summary>
			''' Specifies the extended-key flag, event-injected flag, context code, and transition-state flag.
			''' </summary>
			Public flags As Integer
			''' <summary>
			''' Specifies the time stamp for this message.
			''' </summary>
			Public time As Integer
			''' <summary>
			''' Specifies extra information associated with the message. 
			''' </summary>
			Public dwExtraInfo As Integer
		End Class
		#End Region

		#Region "Windows function imports"
		''' <summary>
		''' The SetWindowsHookEx function installs an application-defined hook procedure into a hook chain. 
		''' You would install a hook procedure to monitor the system for certain types of events. These events 
		''' are associated either with a specific thread or with all threads in the same desktop as the calling thread. 
		''' </summary>
		''' <param name="idHook">
		''' [in] Specifies the type of hook procedure to be installed. This parameter can be one of the following values.
		''' </param>
		''' <param name="lpfn">
		''' [in] Pointer to the hook procedure. If the dwThreadId parameter is zero or specifies the identifier of a 
		''' thread created by a different process, the lpfn parameter must point to a hook procedure in a dynamic-link 
		''' library (DLL). Otherwise, lpfn can point to a hook procedure in the code associated with the current process.
		''' </param>
		''' <param name="hMod">
		''' [in] Handle to the DLL containing the hook procedure pointed to by the lpfn parameter. 
		''' The hMod parameter must be set to NULL if the dwThreadId parameter specifies a thread created by 
		''' the current process and if the hook procedure is within the code associated with the current process. 
		''' </param>
		''' <param name="dwThreadId">
		''' [in] Specifies the identifier of the thread with which the hook procedure is to be associated. 
		''' If this parameter is zero, the hook procedure is associated with all existing threads running in the 
		''' same desktop as the calling thread. 
		''' </param>
		''' <returns>
		''' If the function succeeds, the return value is the handle to the hook procedure.
		''' If the function fails, the return value is NULL. To get extended error information, call GetLastError.
		''' </returns>
		''' <remarks>
		''' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookfunctions/setwindowshookex.asp
		''' </remarks>
		<DllImport("user32.dll", CharSet := CharSet.Auto, CallingConvention := CallingConvention.StdCall, SetLastError := True)> _
		Private Shared Function SetWindowsHookEx(idHook As Integer, lpfn As HookProc, hMod As IntPtr, dwThreadId As Integer) As Integer
		End Function

		''' <summary>
		''' The UnhookWindowsHookEx function removes a hook procedure installed in a hook chain by the SetWindowsHookEx function. 
		''' </summary>
		''' <param name="idHook">
		''' [in] Handle to the hook to be removed. This parameter is a hook handle obtained by a previous call to SetWindowsHookEx. 
		''' </param>
		''' <returns>
		''' If the function succeeds, the return value is nonzero.
		''' If the function fails, the return value is zero. To get extended error information, call GetLastError.
		''' </returns>
		''' <remarks>
		''' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookfunctions/setwindowshookex.asp
		''' </remarks>
		<DllImport("user32.dll", CharSet := CharSet.Auto, CallingConvention := CallingConvention.StdCall, SetLastError := True)> _
		Private Shared Function UnhookWindowsHookEx(idHook As Integer) As Integer
		End Function

		''' <summary>
		''' The CallNextHookEx function passes the hook information to the next hook procedure in the current hook chain. 
		''' A hook procedure can call this function either before or after processing the hook information. 
		''' </summary>
		''' <param name="idHook">Ignored.</param>
		''' <param name="nCode">
		''' [in] Specifies the hook code passed to the current hook procedure. 
		''' The next hook procedure uses this code to determine how to process the hook information.
		''' </param>
		''' <param name="wParam">
		''' [in] Specifies the wParam value passed to the current hook procedure. 
		''' The meaning of this parameter depends on the type of hook associated with the current hook chain. 
		''' </param>
		''' <param name="lParam">
		''' [in] Specifies the lParam value passed to the current hook procedure. 
		''' The meaning of this parameter depends on the type of hook associated with the current hook chain. 
		''' </param>
		''' <returns>
		''' This value is returned by the next hook procedure in the chain. 
		''' The current hook procedure must also return this value. The meaning of the return value depends on the hook type. 
		''' For more information, see the descriptions of the individual hook procedures.
		''' </returns>
		''' <remarks>
		''' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookfunctions/setwindowshookex.asp
		''' </remarks>
		<DllImport("user32.dll", CharSet := CharSet.Auto, CallingConvention := CallingConvention.StdCall)> _
		Private Shared Function CallNextHookEx(idHook As Integer, nCode As Integer, wParam As Integer, lParam As IntPtr) As Integer
		End Function

		''' <summary>
		''' The CallWndProc hook procedure is an application-defined or library-defined callback 
		''' function used with the SetWindowsHookEx function. The HOOKPROC type defines a pointer 
		''' to this callback function. CallWndProc is a placeholder for the application-defined 
		''' or library-defined function name.
		''' </summary>
		''' <param name="nCode">
		''' [in] Specifies whether the hook procedure must process the message. 
		''' If nCode is HC_ACTION, the hook procedure must process the message. 
		''' If nCode is less than zero, the hook procedure must pass the message to the 
		''' CallNextHookEx function without further processing and must return the 
		''' value returned by CallNextHookEx.
		''' </param>
		''' <param name="wParam">
		''' [in] Specifies whether the message was sent by the current thread. 
		''' If the message was sent by the current thread, it is nonzero; otherwise, it is zero. 
		''' </param>
		''' <param name="lParam">
		''' [in] Pointer to a CWPSTRUCT structure that contains details about the message. 
		''' </param>
		''' <returns>
		''' If nCode is less than zero, the hook procedure must return the value returned by CallNextHookEx. 
		''' If nCode is greater than or equal to zero, it is highly recommended that you call CallNextHookEx 
		''' and return the value it returns; otherwise, other applications that have installed WH_CALLWNDPROC 
		''' hooks will not receive hook notifications and may behave incorrectly as a result. If the hook 
		''' procedure does not call CallNextHookEx, the return value should be zero. 
		''' </returns>
		''' <remarks>
		''' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookfunctions/callwndproc.asp
		''' </remarks>
		Private Delegate Function HookProc(nCode As Integer, wParam As Integer, lParam As IntPtr) As Integer

		''' <summary>
		''' The ToAscii function translates the specified virtual-key code and keyboard 
		''' state to the corresponding character or characters. The function translates the code 
		''' using the input language and physical keyboard layout identified by the keyboard layout handle.
		''' </summary>
		''' <param name="uVirtKey">
		''' [in] Specifies the virtual-key code to be translated. 
		''' </param>
		''' <param name="uScanCode">
		''' [in] Specifies the hardware scan code of the key to be translated. 
		''' The high-order bit of this value is set if the key is up (not pressed). 
		''' </param>
		''' <param name="lpbKeyState">
		''' [in] Pointer to a 256-byte array that contains the current keyboard state. 
		''' Each element (byte) in the array contains the state of one key. 
		''' If the high-order bit of a byte is set, the key is down (pressed). 
		''' The low bit, if set, indicates that the key is toggled on. In this function, 
		''' only the toggle bit of the CAPS LOCK key is relevant. The toggle state 
		''' of the NUM LOCK and SCROLL LOCK keys is ignored.
		''' </param>
		''' <param name="lpwTransKey">
		''' [out] Pointer to the buffer that receives the translated character or characters. 
		''' </param>
		''' <param name="fuState">
		''' [in] Specifies whether a menu is active. This parameter must be 1 if a menu is active, or 0 otherwise. 
		''' </param>
		''' <returns>
		''' If the specified key is a dead key, the return value is negative. Otherwise, it is one of the following values. 
		''' Value Meaning 
		''' 0 The specified virtual key has no translation for the current state of the keyboard. 
		''' 1 One character was copied to the buffer. 
		''' 2 Two characters were copied to the buffer. This usually happens when a dead-key character 
		''' (accent or diacritic) stored in the keyboard layout cannot be composed with the specified 
		''' virtual key to form a single character. 
		''' </returns>
		''' <remarks>
		''' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/userinput/keyboardinput/keyboardinputreference/keyboardinputfunctions/toascii.asp
		''' </remarks>
		<DllImport("user32")> _
		Private Shared Function ToAscii(uVirtKey As Integer, uScanCode As Integer, lpbKeyState As Byte(), lpwTransKey As Byte(), fuState As Integer) As Integer
		End Function

		''' <summary>
		''' The GetKeyboardState function copies the status of the 256 virtual keys to the 
		''' specified buffer. 
		''' </summary>
		''' <param name="pbKeyState">
		''' [in] Pointer to a 256-byte array that contains keyboard key states. 
		''' </param>
		''' <returns>
		''' If the function succeeds, the return value is nonzero.
		''' If the function fails, the return value is zero. To get extended error information, call GetLastError. 
		''' </returns>
		''' <remarks>
		''' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/userinput/keyboardinput/keyboardinputreference/keyboardinputfunctions/toascii.asp
		''' </remarks>
		<DllImport("user32")> _
		Private Shared Function GetKeyboardState(pbKeyState As Byte()) As Integer
		End Function

		<DllImport("user32.dll", CharSet := CharSet.Auto, CallingConvention := CallingConvention.StdCall)> _
		Private Shared Function GetKeyState(vKey As Integer) As Short
		End Function

		#End Region

		#Region "Windows constants"

		'values from Winuser.h in Microsoft SDK.
		''' <summary>
		''' Windows NT/2000/XP: Installs a hook procedure that monitors low-level mouse input events.
		''' </summary>
		Private Const WH_MOUSE_LL As Integer = 14
		''' <summary>
		''' Windows NT/2000/XP: Installs a hook procedure that monitors low-level keyboard  input events.
		''' </summary>
		Private Const WH_KEYBOARD_LL As Integer = 13

		''' <summary>
		''' Installs a hook procedure that monitors mouse messages. For more information, see the MouseProc hook procedure. 
		''' </summary>
		Private Const WH_MOUSE As Integer = 7
		''' <summary>
		''' Installs a hook procedure that monitors keystroke messages. For more information, see the KeyboardProc hook procedure. 
		''' </summary>
		Private Const WH_KEYBOARD As Integer = 2

		''' <summary>
		''' The WM_MOUSEMOVE message is posted to a window when the cursor moves. 
		''' </summary>
		Private Const WM_MOUSEMOVE As Integer = &H200
		''' <summary>
		''' The WM_LBUTTONDOWN message is posted when the user presses the left mouse button 
		''' </summary>
		Private Const WM_LBUTTONDOWN As Integer = &H201
		''' <summary>
		''' The WM_RBUTTONDOWN message is posted when the user presses the right mouse button
		''' </summary>
		Private Const WM_RBUTTONDOWN As Integer = &H204
		''' <summary>
		''' The WM_MBUTTONDOWN message is posted when the user presses the middle mouse button 
		''' </summary>
		Private Const WM_MBUTTONDOWN As Integer = &H207
		''' <summary>
		''' The WM_LBUTTONUP message is posted when the user releases the left mouse button 
		''' </summary>
		Private Const WM_LBUTTONUP As Integer = &H202
		''' <summary>
		''' The WM_RBUTTONUP message is posted when the user releases the right mouse button 
		''' </summary>
		Private Const WM_RBUTTONUP As Integer = &H205
		''' <summary>
		''' The WM_MBUTTONUP message is posted when the user releases the middle mouse button 
		''' </summary>
		Private Const WM_MBUTTONUP As Integer = &H208
		''' <summary>
		''' The WM_LBUTTONDBLCLK message is posted when the user double-clicks the left mouse button 
		''' </summary>
		Private Const WM_LBUTTONDBLCLK As Integer = &H203
		''' <summary>
		''' The WM_RBUTTONDBLCLK message is posted when the user double-clicks the right mouse button 
		''' </summary>
		Private Const WM_RBUTTONDBLCLK As Integer = &H206
		''' <summary>
		''' The WM_RBUTTONDOWN message is posted when the user presses the right mouse button 
		''' </summary>
		Private Const WM_MBUTTONDBLCLK As Integer = &H209
		''' <summary>
		''' The WM_MOUSEWHEEL message is posted when the user presses the mouse wheel. 
		''' </summary>
		Private Const WM_MOUSEWHEEL As Integer = &H20a

		''' <summary>
		''' The WM_KEYDOWN message is posted to the window with the keyboard focus when a nonsystem 
		''' key is pressed. A nonsystem key is a key that is pressed when the ALT key is not pressed.
		''' </summary>
		Private Const WM_KEYDOWN As Integer = &H100
		''' <summary>
		''' The WM_KEYUP message is posted to the window with the keyboard focus when a nonsystem 
		''' key is released. A nonsystem key is a key that is pressed when the ALT key is not pressed, 
		''' or a keyboard key that is pressed when a window has the keyboard focus.
		''' </summary>
		Private Const WM_KEYUP As Integer = &H101
		''' <summary>
		''' The WM_SYSKEYDOWN message is posted to the window with the keyboard focus when the user 
		''' presses the F10 key (which activates the menu bar) or holds down the ALT key and then 
		''' presses another key. It also occurs when no window currently has the keyboard focus; 
		''' in this case, the WM_SYSKEYDOWN message is sent to the active window. The window that 
		''' receives the message can distinguish between these two contexts by checking the context 
		''' code in the lParam parameter. 
		''' </summary>
		Private Const WM_SYSKEYDOWN As Integer = &H104
		''' <summary>
		''' The WM_SYSKEYUP message is posted to the window with the keyboard focus when the user 
		''' releases a key that was pressed while the ALT key was held down. It also occurs when no 
		''' window currently has the keyboard focus; in this case, the WM_SYSKEYUP message is sent 
		''' to the active window. The window that receives the message can distinguish between 
		''' these two contexts by checking the context code in the lParam parameter. 
		''' </summary>
		Private Const WM_SYSKEYUP As Integer = &H105

		Private Const VK_SHIFT As Byte = &H10
		Private Const VK_CAPITAL As Byte = &H14
		Private Const VK_NUMLOCK As Byte = &H90

		#End Region

		''' <summary>
		''' Creates an instance of UserActivityHook object and sets mouse and keyboard hooks.
		''' </summary>
		''' <exception cref="Win32Exception">Any windows problem.</exception>
		Public Sub New()
			Start()
		End Sub

		''' <summary>
		''' Creates an instance of UserActivityHook object and installs both or one of mouse and/or keyboard hooks and starts rasing events
		''' </summary>
		''' <param name="InstallMouseHook"><b>true</b> if mouse events must be monitored</param>
		''' <param name="InstallKeyboardHook"><b>true</b> if keyboard events must be monitored</param>
		''' <exception cref="Win32Exception">Any windows problem.</exception>
		''' <remarks>
		''' To create an instance without installing hooks call new UserActivityHook(false, false)
		''' </remarks>
		Public Sub New(InstallMouseHook As Boolean, InstallKeyboardHook As Boolean)
			Start(InstallMouseHook, InstallKeyboardHook)
		End Sub

		''' <summary>
		''' Destruction.
		''' </summary>
		Protected Overrides Sub Finalize()
			Try
				'uninstall hooks and do not throw exceptions
				[Stop](True, True, False)
			Finally
				MyBase.Finalize()
			End Try
		End Sub

		''' <summary>
		''' Occurs when the user moves the mouse, presses any mouse button or scrolls the wheel
		''' </summary>
		Public Event OnMouseActivity As MouseEventHandler
		''' <summary>
		''' Occurs when the user presses a key
		''' </summary>
		Public Event KeyDown As KeyEventHandler
		''' <summary>
		''' Occurs when the user presses and releases 
		''' </summary>
		Public Event KeyPress As KeyPressEventHandler
		''' <summary>
		''' Occurs when the user releases a key
		''' </summary>
		Public Event KeyUp As KeyEventHandler


		''' <summary>
		''' Stores the handle to the mouse hook procedure.
		''' </summary>
		Private hMouseHook As Integer = 0
		''' <summary>
		''' Stores the handle to the keyboard hook procedure.
		''' </summary>
		Private hKeyboardHook As Integer = 0


		''' <summary>
		''' Declare MouseHookProcedure as HookProc type.
		''' </summary>
		Private Shared MouseHookProcedure As HookProc
		''' <summary>
		''' Declare KeyboardHookProcedure as HookProc type.
		''' </summary>
		Private Shared KeyboardHookProcedure As HookProc


		''' <summary>
		''' Installs both mouse and keyboard hooks and starts rasing events
		''' </summary>
		''' <exception cref="Win32Exception">Any windows problem.</exception>
		Public Sub Start()
			Me.Start(True, True)
		End Sub

		''' <summary>
		''' Installs both or one of mouse and/or keyboard hooks and starts rasing events
		''' </summary>
		''' <param name="InstallMouseHook"><b>true</b> if mouse events must be monitored</param>
		''' <param name="InstallKeyboardHook"><b>true</b> if keyboard events must be monitored</param>
		''' <exception cref="Win32Exception">Any windows problem.</exception>
		Public Sub Start(InstallMouseHook As Boolean, InstallKeyboardHook As Boolean)
			' install Mouse hook only if it is not installed and must be installed
			If hMouseHook = 0 AndAlso InstallMouseHook Then
				' Create an instance of HookProc.
				MouseHookProcedure = New HookProc(AddressOf MouseHookProc)
				'install hook
				hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseHookProcedure, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()(0)), 0)
				'If SetWindowsHookEx fails.
				If hMouseHook = 0 Then
					'Returns the error code returned by the last unmanaged function called using platform invoke that has the DllImportAttribute.SetLastError flag set. 
					Dim errorCode As Integer = Marshal.GetLastWin32Error()
					'do cleanup
					[Stop](True, False, False)
					'Initializes and throws a new instance of the Win32Exception class with the specified error. 
					Throw New Win32Exception(errorCode)
				End If
			End If

			' install Keyboard hook only if it is not installed and must be installed
			If hKeyboardHook = 0 AndAlso InstallKeyboardHook Then
				' Create an instance of HookProc.
				KeyboardHookProcedure = New HookProc(AddressOf KeyboardHookProc)
				'install hook
				hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyboardHookProcedure, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()(0)), 0)
				'If SetWindowsHookEx fails.
				If hKeyboardHook = 0 Then
					'Returns the error code returned by the last unmanaged function called using platform invoke that has the DllImportAttribute.SetLastError flag set. 
					Dim errorCode As Integer = Marshal.GetLastWin32Error()
					'do cleanup
					[Stop](False, True, False)
					'Initializes and throws a new instance of the Win32Exception class with the specified error. 
					Throw New Win32Exception(errorCode)
				End If
			End If
		End Sub

		''' <summary>
		''' Stops monitoring both mouse and keyboard events and rasing events.
		''' </summary>
		''' <exception cref="Win32Exception">Any windows problem.</exception>
		Public Sub [Stop]()
			Me.[Stop](True, True, True)
		End Sub

		''' <summary>
		''' Stops monitoring both or one of mouse and/or keyboard events and rasing events.
		''' </summary>
		''' <param name="UninstallMouseHook"><b>true</b> if mouse hook must be uninstalled</param>
		''' <param name="UninstallKeyboardHook"><b>true</b> if keyboard hook must be uninstalled</param>
		''' <param name="ThrowExceptions"><b>true</b> if exceptions which occured during uninstalling must be thrown</param>
		''' <exception cref="Win32Exception">Any windows problem.</exception>
		Public Sub [Stop](UninstallMouseHook As Boolean, UninstallKeyboardHook As Boolean, ThrowExceptions As Boolean)
			'if mouse hook set and must be uninstalled
			If hMouseHook <> 0 AndAlso UninstallMouseHook Then
				'uninstall hook
				Dim retMouse As Integer = UnhookWindowsHookEx(hMouseHook)
				'reset invalid handle
				hMouseHook = 0
				'if failed and exception must be thrown
				If retMouse = 0 AndAlso ThrowExceptions Then
					'Returns the error code returned by the last unmanaged function called using platform invoke that has the DllImportAttribute.SetLastError flag set. 
					Dim errorCode As Integer = Marshal.GetLastWin32Error()
					'Initializes and throws a new instance of the Win32Exception class with the specified error. 
					Throw New Win32Exception(errorCode)
				End If
			End If

			'if keyboard hook set and must be uninstalled
			If hKeyboardHook <> 0 AndAlso UninstallKeyboardHook Then
				'uninstall hook
				Dim retKeyboard As Integer = UnhookWindowsHookEx(hKeyboardHook)
				'reset invalid handle
				hKeyboardHook = 0
				'if failed and exception must be thrown
				If retKeyboard = 0 AndAlso ThrowExceptions Then
					'Returns the error code returned by the last unmanaged function called using platform invoke that has the DllImportAttribute.SetLastError flag set. 
					Dim errorCode As Integer = Marshal.GetLastWin32Error()
					'Initializes and throws a new instance of the Win32Exception class with the specified error. 
					Throw New Win32Exception(errorCode)
				End If
			End If
		End Sub


		''' <summary>
		''' A callback function which will be called every time a mouse activity detected.
		''' </summary>
		''' <param name="nCode">
		''' [in] Specifies whether the hook procedure must process the message. 
		''' If nCode is HC_ACTION, the hook procedure must process the message. 
		''' If nCode is less than zero, the hook procedure must pass the message to the 
		''' CallNextHookEx function without further processing and must return the 
		''' value returned by CallNextHookEx.
		''' </param>
		''' <param name="wParam">
		''' [in] Specifies whether the message was sent by the current thread. 
		''' If the message was sent by the current thread, it is nonzero; otherwise, it is zero. 
		''' </param>
		''' <param name="lParam">
		''' [in] Pointer to a CWPSTRUCT structure that contains details about the message. 
		''' </param>
		''' <returns>
		''' If nCode is less than zero, the hook procedure must return the value returned by CallNextHookEx. 
		''' If nCode is greater than or equal to zero, it is highly recommended that you call CallNextHookEx 
		''' and return the value it returns; otherwise, other applications that have installed WH_CALLWNDPROC 
		''' hooks will not receive hook notifications and may behave incorrectly as a result. If the hook 
		''' procedure does not call CallNextHookEx, the return value should be zero. 
		''' </returns>
		Private Function MouseHookProc(nCode As Integer, wParam As Integer, lParam As IntPtr) As Integer
			' if ok and someone listens to our events
			If (nCode >= 0) AndAlso (OnMouseActivity IsNot Nothing) Then
				'Marshall the data from callback.
				Dim mouseHookStruct As MouseLLHookStruct = DirectCast(Marshal.PtrToStructure(lParam, GetType(MouseLLHookStruct)), MouseLLHookStruct)

				'detect button clicked
				Dim button As MouseButtons = MouseButtons.None
				Dim mouseDelta As Short = 0
				Select Case wParam
					Case WM_LBUTTONDOWN
						'case WM_LBUTTONUP: 
						'case WM_LBUTTONDBLCLK: 
						button = MouseButtons.Left
						Exit Select
					Case WM_RBUTTONDOWN
						'case WM_RBUTTONUP: 
						'case WM_RBUTTONDBLCLK: 
						button = MouseButtons.Right
						Exit Select
					Case WM_MOUSEWHEEL
						'If the message is WM_MOUSEWHEEL, the high-order word of mouseData member is the wheel delta. 
						'One wheel click is defined as WHEEL_DELTA, which is 120. 
						'(value >> 16) & 0xffff; retrieves the high-order word from the given 32-bit value
						mouseDelta = CShort((mouseHookStruct.mouseData >> 16) And &Hffff)
						'TODO: X BUTTONS (I havent them so was unable to test)
						'If the message is WM_XBUTTONDOWN, WM_XBUTTONUP, WM_XBUTTONDBLCLK, WM_NCXBUTTONDOWN, WM_NCXBUTTONUP, 
						'or WM_NCXBUTTONDBLCLK, the high-order word specifies which X button was pressed or released, 
						'and the low-order word is reserved. This value can be one or more of the following values. 
						'Otherwise, mouseData is not used. 
						Exit Select
				End Select

				'double clicks
				Dim clickCount As Integer = 0
				If button <> MouseButtons.None Then
					If wParam = WM_LBUTTONDBLCLK OrElse wParam = WM_RBUTTONDBLCLK Then
						clickCount = 2
					Else
						clickCount = 1
					End If
				End If

				'generate event 
				Dim e As New MouseEventArgs(button, clickCount, mouseHookStruct.pt.x, mouseHookStruct.pt.y, mouseDelta)
				'raise it
				RaiseEvent OnMouseActivity(Me, e)
			End If
			'call next hook
			Return CallNextHookEx(hMouseHook, nCode, wParam, lParam)
		End Function

		''' <summary>
		''' A callback function which will be called every time a keyboard activity detected.
		''' </summary>
		''' <param name="nCode">
		''' [in] Specifies whether the hook procedure must process the message. 
		''' If nCode is HC_ACTION, the hook procedure must process the message. 
		''' If nCode is less than zero, the hook procedure must pass the message to the 
		''' CallNextHookEx function without further processing and must return the 
		''' value returned by CallNextHookEx.
		''' </param>
		''' <param name="wParam">
		''' [in] Specifies whether the message was sent by the current thread. 
		''' If the message was sent by the current thread, it is nonzero; otherwise, it is zero. 
		''' </param>
		''' <param name="lParam">
		''' [in] Pointer to a CWPSTRUCT structure that contains details about the message. 
		''' </param>
		''' <returns>
		''' If nCode is less than zero, the hook procedure must return the value returned by CallNextHookEx. 
		''' If nCode is greater than or equal to zero, it is highly recommended that you call CallNextHookEx 
		''' and return the value it returns; otherwise, other applications that have installed WH_CALLWNDPROC 
		''' hooks will not receive hook notifications and may behave incorrectly as a result. If the hook 
		''' procedure does not call CallNextHookEx, the return value should be zero. 
		''' </returns>
		Private Function KeyboardHookProc(nCode As Integer, wParam As Int32, lParam As IntPtr) As Integer
			'indicates if any of underlaing events set e.Handled flag
			Dim handled As Boolean = False
			'it was ok and someone listens to events
			If (nCode >= 0) AndAlso (KeyDown IsNot Nothing OrElse KeyUp IsNot Nothing OrElse KeyPress IsNot Nothing) Then
				'read structure KeyboardHookStruct at lParam
				Dim MyKeyboardHookStruct As KeyboardHookStruct = DirectCast(Marshal.PtrToStructure(lParam, GetType(KeyboardHookStruct)), KeyboardHookStruct)
				'raise KeyDown
				If KeyDown IsNot Nothing AndAlso (wParam = WM_KEYDOWN OrElse wParam = WM_SYSKEYDOWN) Then
					Dim keyData As Keys = DirectCast(MyKeyboardHookStruct.vkCode, Keys)
					Dim e As New KeyEventArgs(keyData)
					RaiseEvent KeyDown(Me, e)
					handled = handled OrElse e.Handled
				End If

				' raise KeyPress
				If KeyPress IsNot Nothing AndAlso wParam = WM_KEYDOWN Then
					Dim isDownShift As Boolean = (If((GetKeyState(VK_SHIFT) And &H80) = &H80, True, False))
					Dim isDownCapslock As Boolean = (If(GetKeyState(VK_CAPITAL) <> 0, True, False))

					Dim keyState As Byte() = New Byte(255) {}
					GetKeyboardState(keyState)
					Dim inBuffer As Byte() = New Byte(1) {}
					If ToAscii(MyKeyboardHookStruct.vkCode, MyKeyboardHookStruct.scanCode, keyState, inBuffer, MyKeyboardHookStruct.flags) = 1 Then
						Dim key As Char = CChar(inBuffer(0))
						If (isDownCapslock Xor isDownShift) AndAlso [Char].IsLetter(key) Then
							key = [Char].ToUpper(key)
						End If
						Dim e As New KeyPressEventArgs(key)
						RaiseEvent KeyPress(Me, e)
						handled = handled OrElse e.Handled
					End If
				End If

				' raise KeyUp
				If KeyUp IsNot Nothing AndAlso (wParam = WM_KEYUP OrElse wParam = WM_SYSKEYUP) Then
					Dim keyData As Keys = DirectCast(MyKeyboardHookStruct.vkCode, Keys)
					Dim e As New KeyEventArgs(keyData)
					RaiseEvent KeyUp(Me, e)
					handled = handled OrElse e.Handled

				End If
			End If

			'if event handled in application do not handoff to other listeners
			If handled Then
				Return 1
			Else
				Return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam)
			End If
		End Function
	End Class
End Namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
