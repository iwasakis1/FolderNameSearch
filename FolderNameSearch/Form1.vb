Imports System.IO
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

Public Class Form1
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function FindFirstFile(ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As IntPtr
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        TextBox1.Text = ""
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()

        Dim ds = Directory.EnumerateDirectories(TextBox1.Text, "*", SearchOption.TopDirectoryOnly).ToArray

        mainl(ds)


        'Directory.GetDirectories(TextBox1.Text, "*", SearchOption.AllDirectories)
        Label1.Text = "検索終了"

    End Sub

    Private Sub mainl(ByVal ss() As String)
        Dim atta = False
        For Each s In ss
            If SkipCheck(s) Then Continue For
            Dim bs = backupchk(s)
            If bs <> "" Then
                atta = True
                Exit For
            End If
        Next

        If atta Then
            For Each s In ss
                If SkipCheck(s) Then Continue For
                If chekstr(s) Then
                    ListBox1.Items.Add(s)
                    ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                End If
            Next
        Else
            For Each s In ss
                If SkipCheck(s) Then Continue For
                Try
                    Dim ds = Directory.EnumerateDirectories(s, "*", SearchOption.TopDirectoryOnly).ToArray
                    mainl(ds)
                Catch ex As Exception
                End Try
            Next
        End If
    End Sub

    Private Function chekstr(ByVal s As String) As Boolean
        If (Regex.IsMatch(s, $"{TextBox2.Text}", RegexOptions.IgnoreCase)) Then
            Return True
        End If
        Return False
    End Function

    Private Function SkipCheck(ByVal s As String) As Boolean
        'Debug.WriteLine(s)
        Label1.Text = s
        Application.DoEvents()

        If (Regex.IsMatch(s, ".git|.vs|obj|bin|my project", RegexOptions.IgnoreCase)) Then
            Return True
        End If
        Return False
    End Function


    Private Function backupchk(ByVal ss As String) As String
        Dim fs = IO.Path.GetFileName(ss)
        If fs.IndexOf("backup", StringComparison.CurrentCultureIgnoreCase) >= 0 Then
            Return ss
        End If
        Return ""
    End Function



    ' The CharSet must match the CharSet of the corresponding PInvoke signature
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Structure WIN32_FIND_DATA
        Public dwFileAttributes As UInteger
        Public ftCreationTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public ftLastAccessTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public ftLastWriteTime As System.Runtime.InteropServices.ComTypes.FILETIME
        Public nFileSizeHigh As UInteger
        Public nFileSizeLow As UInteger
        Public dwReserved0 As UInteger
        Public dwReserved1 As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> Public cFileName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)> Public cAlternateFileName As String
    End Structure

End Class
