Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Public Function CopyDirectory(ByVal SrcPath As String, ByVal DestPath As String, Optional ByVal bQuiet As Boolean = False) As Boolean
        If Not System.IO.Directory.Exists(SrcPath) Then
            Throw New System.IO.DirectoryNotFoundException("The directory " & SrcPath & " does not exists")
        End If
        If System.IO.Directory.Exists(DestPath) AndAlso Not bQuiet Then
            If MessageBox.Show("directory " & DestPath & " already exists." & vbCrLf &
            "If you continue, any files with the same name will be overwritten",
            "Continue?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Cancel Then Exit Function
        End If

        'add Directory Seperator Character (\) for the string concatenation shown later
        If DestPath.Substring(DestPath.Length - 1, 1) <> System.IO.Path.DirectorySeparatorChar Then
            DestPath += System.IO.Path.DirectorySeparatorChar
        End If
        If Not System.IO.Directory.Exists(DestPath) Then System.IO.Directory.CreateDirectory(DestPath)
        Dim Files As String()
        Files = System.IO.Directory.GetFileSystemEntries(SrcPath)
        Dim element As String
        For Each element In Files
            If System.IO.Directory.Exists(element) Then
                'if the current FileSystemEntry is a directory,
                'call this function recursively
                CopyDirectory(element, DestPath & System.IO.Path.GetFileName(element), True)
            Else
                'the current FileSystemEntry is a file so just copy it
                System.IO.File.Copy(element, DestPath & System.IO.Path.GetFileName(element), True)
            End If
        Next
        Return True
    End Function
End Class
