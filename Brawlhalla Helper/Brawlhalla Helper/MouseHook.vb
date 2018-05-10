Imports System.ComponentModel
Imports System.Runtime.InteropServices

Public Class globalInputHook
    Inherits Component

#Region "     mouseHook"

    Private Declare Function GetDoubleClickTime Lib "user32" () As Integer

    Private hMouseHook As IntPtr = IntPtr.Zero
    Private disposedValue As Boolean = False        ' To detect redundant calls

    Public Event MouseDown As MouseEventHandler
    Public Event MouseUp As MouseEventHandler
    Public Event MouseMove As MouseEventHandler
    Public Event MouseDoubleclick As MouseEventHandler

    Private Structure API_POINT
        Public x As Integer
        Public y As Integer
    End Structure

    Private Structure MSLLHOOKSTRUCT
        Public pt As API_POINT
        Public mouseData As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    Private Const WM_MOUSEWHEEL As Integer = &H20A
    Private Const WM_MOUSEHWHEEL As Integer = &H20E
    Private Const WM_MOUSEMOVE As Integer = &H200
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const WM_MBUTTONDOWN As Integer = &H207
    Private Const WM_MBUTTONUP As Integer = &H208
    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_RBUTTONUP As Integer = &H205
    Private Const WH_MOUSE_LL As Integer = 14
    Private Const WM_XBUTTONDOWN As Integer = &H20B
    Private Const WM_XBUTTONUP As Integer = &H20C
    Private Const XBUTTON1 As Integer = &H1
    Private Const XBUTTON2 As Integer = &H2

    Private Declare Auto Function LoadLibrary Lib "kernel32" (ByVal _
                lpFileName As String) As IntPtr
    Private Declare Auto Function SetWindowsHookEx Lib "user32.dll" (ByVal _
                idHook As Integer, ByVal lpfn As LowLevelMouseHookProc, ByVal hInstance As _
                IntPtr, ByVal dwThreadId As Integer) As IntPtr
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hhk As _
                IntPtr, ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As _
                MSLLHOOKSTRUCT) As Integer
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As _
                IntPtr) As Boolean

    Private Delegate Function LowLevelMouseHookProc(ByVal nCode As Integer,
                ByVal wParam As Integer, ByRef lParam As MSLLHOOKSTRUCT) As Integer

    Private Sub HookMouse()
        Dim hInstance As IntPtr = LoadLibrary("User32")
        hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf Me.HookProc, hInstance, 0)
    End Sub

    Private Sub UnhookMouse()
        UnhookWindowsHookEx(hMouseHook)
    End Sub

    Private Function HookProc(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As MSLLHOOKSTRUCT) As Integer

        Try
            If nCode >= 0 Then
                Select Case wParam
                    Case WM_XBUTTONDOWN
                        RaiseEvent MouseDown(Me, New MouseEventArgs(MouseButtons.XButton1, 0, lParam.pt.x, lParam.pt.y, 0))
                    Case WM_LBUTTONDOWN
                        RaiseEvent MouseDown(Me, New MouseEventArgs(MouseButtons.Left, 0, lParam.pt.x, lParam.pt.y, 0))
                    Case WM_RBUTTONDOWN
                        RaiseEvent MouseDown(Me, New MouseEventArgs(MouseButtons.Right, 0, lParam.pt.x, lParam.pt.y, 0))
                    Case WM_MBUTTONDOWN
                        RaiseEvent MouseDown(Me, New MouseEventArgs(MouseButtons.Middle, 0, lParam.pt.x, lParam.pt.y, 0))
                    Case WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP
                        Dim index As Integer = If(wParam = WM_LBUTTONUP, 0, If(wParam = WM_RBUTTONUP, 1, If(wParam = WM_MBUTTONUP, 2, -1)))
                        Static lastMouseUp(2) As Integer
                        If Environment.TickCount - lastMouseUp(index) < GetDoubleClickTime Then
                            Dim buttons() As MouseButtons = {MouseButtons.Left, MouseButtons.Right, MouseButtons.Middle}
                            RaiseEvent MouseDoubleclick(Me, New MouseEventArgs(buttons(index), 0, lParam.pt.x, lParam.pt.y, 0))
                        End If
                        lastMouseUp(index) = Environment.TickCount
                        RaiseEvent MouseUp(Nothing, Nothing)
                    Case WM_MOUSEMOVE
                        RaiseEvent MouseMove(Nothing, Nothing)
                    Case WM_MOUSEWHEEL, WM_MOUSEHWHEEL
                    Case Else
                        Console.WriteLine(wParam)
                End Select
            End If
            Return CallNextHookEx(hMouseHook, nCode, wParam, lParam)

        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Function

    Private Function mouseHooked() As Boolean
        Return hMouseHook <> IntPtr.Zero
    End Function

#End Region

#Region "     keyboardHook"

    Private Declare Function SetWindowsHookEx Lib "user32" _
    Alias "SetWindowsHookExA" (ByVal idHook As Integer,
    ByVal lpfn As lowlevelKeyboardHookProc, ByVal hmod As IntPtr,
    ByVal dwThreadId As Integer) As IntPtr

    Private Delegate Function lowlevelKeyboardHookProc(
                              ByVal Code As Integer,
                              ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) _
                                           As IntPtr

    Private Declare Function CallNextHookEx Lib "user32" _
      (ByVal hHook As IntPtr,
      ByVal nCode As Integer,
      ByVal wParam As Integer,
      ByVal lParam As KBDLLHOOKSTRUCT) As IntPtr

    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As Integer
        Public scanCode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    ' Low-Level Keyboard Constants
    Private Const HC_ACTION As Integer = 0

    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_KEYUP As Integer = &H101
    Private Const WM_SYSKEYDOWN As Integer = &H104
    Private Const WM_SYSKEYUP As Integer = &H105

    Private Const WH_KEYBOARD_LL As Integer = 13

    Public hKeyboardHook As IntPtr

    Public Event KeyDown As KeyEventHandler
    Public Event KeyPress As KeyPressEventHandler
    Public Event KeyUp As KeyEventHandler

    Private Declare Function GetAsyncKeyState Lib "user32" _
    (ByVal vKey As Integer) As Integer

    Private Declare Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Integer) As Integer

    Private Const VK_ALT As Integer = &H12
    Private Const VK_CONTROL As Integer = &H11
    Private Const VK_SHIFT As Integer = 16

    <MarshalAs(UnmanagedType.FunctionPtr)> Private callback As lowlevelKeyboardHookProc

    Private Sub HookKeyboard()
        callback = New lowlevelKeyboardHookProc(AddressOf KeyboardCallback)

        Dim hInstance As IntPtr = LoadLibrary("User32")
        hKeyboardHook = SetWindowsHookEx(
          WH_KEYBOARD_LL, callback,
          hInstance, 0)

        CheckHooked()
    End Sub
    Private Function KeyboardCallback(ByVal Code As Integer,
      ByVal wParam As Integer,
      ByRef lParam As KBDLLHOOKSTRUCT) As IntPtr

        Try
            If (Code = HC_ACTION) Then
                Dim CapsLock As Boolean = GetKeyState(Keys.CapsLock) = 1
                Dim shifting As Boolean = False
                Dim modifiers As Keys
                If GetAsyncKeyState(VK_CONTROL) <> 0 Then
                    modifiers = modifiers Or Keys.Control
                End If
                If GetAsyncKeyState(VK_SHIFT) <> 0 Then
                    modifiers = modifiers Or Keys.Shift
                    shifting = True
                End If
                If GetAsyncKeyState(VK_ALT) <> 0 Then
                    modifiers = modifiers Or Keys.Alt
                End If
                Static lastKeys As Keys
                Select Case wParam
                    Case WM_KEYDOWN, WM_SYSKEYDOWN
                        RaiseEvent KeyDown(Me, New KeyEventArgs(DirectCast(Asc(Chr(lParam.vkCode)), Keys) Or modifiers))
                        If lastKeys <> (DirectCast(Asc(Chr(lParam.vkCode)), Keys) Or modifiers) Then
                            lastKeys = (DirectCast(Asc(Chr(lParam.vkCode)), Keys) Or modifiers)
                            If CapsLock AndAlso shifting Then
                                RaiseEvent KeyPress(Me, New KeyPressEventArgs(Char.ToLower(Chr(lParam.vkCode))))
                            ElseIf Not CapsLock AndAlso shifting Then
                                RaiseEvent KeyPress(Me, New KeyPressEventArgs(Char.ToUpper(Chr(lParam.vkCode))))
                            ElseIf Not shifting Then
                                If CapsLock Then
                                    RaiseEvent KeyPress(Me, New KeyPressEventArgs(Char.ToUpper(Chr(lParam.vkCode))))
                                Else
                                    RaiseEvent KeyPress(Me, New KeyPressEventArgs(Char.ToLower(Chr(lParam.vkCode))))
                                End If
                            End If
                        End If
                    Case WM_KEYUP, WM_SYSKEYUP
                        If CapsLock AndAlso shifting Then
                            RaiseEvent KeyUp(Me, New KeyEventArgs(DirectCast(Asc(Chr(lParam.vkCode)), Keys) Or modifiers))
                        ElseIf Not CapsLock AndAlso shifting Then
                            RaiseEvent KeyUp(Me, New KeyEventArgs(DirectCast(Asc(Chr(lParam.vkCode)), Keys) Or modifiers))
                        ElseIf Not shifting Then
                            If CapsLock Then
                                RaiseEvent KeyUp(Me, New KeyEventArgs(DirectCast(Asc(Chr(lParam.vkCode)), Keys) Or modifiers))
                            Else
                                RaiseEvent KeyUp(Me, New KeyEventArgs(DirectCast(Asc(Chr(lParam.vkCode)), Keys) Or modifiers))
                            End If
                        End If
                        lastKeys = Nothing
                End Select
            End If
            '   Return CallNextHookEx(hKeyboardHook, Code, wParam, lParam)
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Function

    Private Function keyboardHooked() As Boolean
        Return hKeyboardHook <> IntPtr.Zero
    End Function

    Private Sub UnhookKeyboard()
        If keyboardHooked() Then
            UnhookWindowsHookEx(hKeyboardHook)
        End If
    End Sub

#End Region

    Private Sub CheckHooked()
        If keyboardHooked() Then
            Debug.WriteLine("Keyboard hooked")
        Else
            Debug.WriteLine("Keyboard hook failed: " & Err.LastDllError)
        End If
        If mouseHooked() Then
            Debug.WriteLine("Mouse hooked")
        Else
            Debug.WriteLine("Mouse hook failed: " & Err.LastDllError)
        End If
    End Sub

    Public Sub hookInput()
        HookMouse()
        HookKeyboard()
    End Sub

    Public Sub unhookInput()
        UnhookMouse()
        UnhookKeyboard()
    End Sub
End Class