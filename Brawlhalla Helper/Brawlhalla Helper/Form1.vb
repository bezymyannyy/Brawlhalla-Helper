Option Explicit On

Imports System.Runtime.InteropServices
Imports System.IO


Public Class Form1
    WithEvents globalInputHook1 As globalInputHook

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        globalInputHook1 = New globalInputHook
        globalInputHook1.hookInput()
    End Sub

    Private Sub Form1_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        globalInputHook1.unhookInput()
    End Sub

    Private Sub GlobalInputHook1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles globalInputHook1.MouseDown
        If CheckBox1.Checked Then
            If CheckBox2.Checked Then
                If e.Button = MouseButtons.XButton1 Then
                    My.Computer.Keyboard.SendKeys(Keys.Z)
                End If
            End If
            If CheckBox3.Checked Then
                If e.Button = MouseButtons.Middle Then
                    For i = 0 To 20
                        My.Computer.Keyboard.SendKeys(Keys.MButton)
                    Next i
                End If
            End If
        End If
    End Sub
End Class
