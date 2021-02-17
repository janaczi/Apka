Imports System.IO
Imports System.Security
Imports System.Runtime.InteropServices
Public Class Form1
    '1) Zamapowanie pierwszej wolnej literki
    '- zapis do logu wykonanych kroków
    '- 
    '2) Skopiowanie wszystkiego z podanej lokalizacji
    Private rStatus As Integer = 0, rKomunikat As String = ""
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim FromPath As String = "\\PIJT-CGCSNN2\test$\ja"
        Dim ToPath As String = "C:\Users\stav\Pictures\test"
        Dim Username As String = "TECHNOKABEL\JTruminski", Password As String = "Jana84ik"
        SaveToLog("FromPath: " + FromPath.ToString)
        SaveToLog("ToPath: " + ToPath.ToString)
        SaveToLog("sername: " + Username.ToString)
        SaveToLog("Password: " + Len(Password).ToString)

        Dim Network As String = ""
        If Mid(FromPath, 1, 2) = "\\" Then
            Network = FromPath
        ElseIf Mid(ToPath, 1, 2) = "\\" Then
            Network = ToPath
        End If
        SaveToLog("Network: " + Network.ToString)
        If Network <> "" Then
            Dim Letter As String = ""
            MapDrive(UNCPath:=Network, Username:=Username, Password:=Password, rStatus:=rStatus, rKomunikat:=rKomunikat, rDriveLetter:=Letter)
            If rStatus = 1 Then
                SaveToLog("Zamapowano pod: " + Letter.ToString + " " + If(rKomunikat, ""))
                Dim DriveLetter As String = Letter + ":\"

            Else
                SaveToLog("Błąd podczas mapowania: " + If(rKomunikat, ""))
            End If
        Else
            SaveToLog("Brak informacji o udziale.")
        End If

    End Sub
    Public Sub SaveToLog(ByVal Text As String)
        Dim path As String = "NetworkCopyFilesLog.txt"
        Dim Data As String = String.Format("yyyy-MM-dd HH:mm:ss", DateTime.Now)
        Text = Data + " " + Text
        If Not File.Exists(path) Then
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Text.ToString)
            End Using
        End If

        Using sw As StreamWriter = File.AppendText(path)
            sw.WriteLine(Text.ToString)
        End Using
        AddLine(Text)
    End Sub
    Private Sub AddLine(ByVal line As String)
        Me.TextBox1.Text = If(Me.TextBox1.Text = String.Empty, line, Me.TextBox1.Text & ControlChars.CrLf & line)
    End Sub
    Public Function CopyDirectory(ByVal SrcPath As String, ByVal DestPath As String, Optional ByVal bQuiet As Boolean = False) As Boolean
        If Not Directory.Exists(SrcPath) Then
            SaveToLog("The directory " & SrcPath & " does not exists")
            'Throw New System.IO.DirectoryNotFoundException("The directory " & SrcPath & " does not exists")
        End If
        If Directory.Exists(DestPath) AndAlso Not bQuiet Then
            SaveToLog("directory " & DestPath & " already exists." & vbCrLf &
            "If you continue, any files with the same name will be overwritten.")
            'If MessageBox.Show("directory " & DestPath & " already exists." & vbCrLf &
            '"If you continue, any files with the same name will be overwritten",
            '"Continue?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question,
            'MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Cancel Then Exit Function
        End If

        'add Directory Seperator Character (\) for the string concatenation shown later
        If DestPath.Substring(DestPath.Length - 1, 1) <> Path.DirectorySeparatorChar Then
            DestPath += Path.DirectorySeparatorChar
        End If
        If Not Directory.Exists(DestPath) Then
            Directory.CreateDirectory(DestPath)
            SaveToLog("Utworzono folder: " & DestPath)
        End If
        Dim Files As String()
        Files = Directory.GetFileSystemEntries(SrcPath)
        Dim element As String
        For Each element In Files
            If Directory.Exists(element) Then
                'if the current FileSystemEntry is a directory,
                'call this function recursively
                CopyDirectory(element, DestPath & System.IO.Path.GetFileName(element), True)
            Else
                'the current FileSystemEntry is a file so just copy it
                File.Copy(element, DestPath & System.IO.Path.GetFileName(element), True)
                SaveToLog(element + "-->" & DestPath & System.IO.Path.GetFileName(element))
            End If
        Next
        Return True
    End Function

    Public Declare Function WNetAddConnection2 _
        Lib "mpr.dll" Alias "WNetAddConnection2A" _
        (
            ByRef lpNetResource As NETRESOURCE,
            ByVal lpPassword As String,
            ByVal lpUserName As String,
            ByVal dwFlags As Integer) As Integer

    Public Declare Function WNetCancelConnection2 _
        Lib "mpr" Alias "WNetCancelConnection2A" _
        (
            ByVal lpName As String,
            ByVal dwFlags As Integer,
            ByVal fForce As Integer) As Integer

    <StructLayout(LayoutKind.Sequential)>
    Public Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String
    End Structure
    Public Const ForceDisconnect As Integer = 1
    Public Const RESOURCETYPE_DISK As Long = &H1
    Public Sub MapDrive(
                       ByVal UNCPath As String, ByVal Username As String, ByVal Password As String _
                       , ByRef rStatus As Integer, ByRef rKomunikat As String, ByRef rDriveLetter As String
                      )
        Try
            rStatus = 0
            rKomunikat = ""
            rDriveLetter = ""

            Dim LetterAscii As Integer = 69 '69=E
            Dim Letter As String = Chr(LetterAscii)

            Dim nr As NETRESOURCE
            Dim result As Integer = 1

            While LetterAscii <= 90 And result <> 0
                LetterAscii += 1
                Letter = Chr(LetterAscii)
                nr = New NETRESOURCE With {
                    .lpRemoteName = UNCPath,
                    .lpLocalName = Letter & ":",
                    .dwType = RESOURCETYPE_DISK
                }
                result = WNetAddConnection2(nr, Password, Username, 0)
            End While

            If result = 0 Then
                rStatus = 1
                rKomunikat = ""
                rDriveLetter = Letter
            Else
                If LetterAscii > 90 Then
                    rStatus = 0
                    rKomunikat = "Brak dostępnych liter do zamapowania udziału."
                    rDriveLetter = ""
                Else
                    rStatus = 1
                    rKomunikat = "Problem z mapowanie udziału."
                    rDriveLetter = Letter
                End If
            End If

        Catch ex As Exception
            rStatus = 0
            rKomunikat = "Problem: " + ex.Message.ToString
            rDriveLetter = ""
        End Try
    End Sub

    Public Sub UnMapDrive(ByVal DriveLetter As String, ByRef rStatus As Integer, ByRef rKomunikat As String)
        Dim rc As Integer
        rc = WNetCancelConnection2(DriveLetter & ":", 0, ForceDisconnect)
        If rc = 0 Then
            rStatus = 1
            rKomunikat = ""
        Else
            rStatus = 0
            rKomunikat = "Nieodmontowano dysku: " + DriveLetter + ". Błąd numer: " + rc.ToString
        End If
    End Sub

End Class
