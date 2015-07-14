Imports System.Security.Permissions
Imports System.Security.Policy
Imports System.Security.Principal
Imports System.Threading


Module MainModule

    Public Declare Function LoadCursorFromFile Lib "user32" (ByVal Filename As String) As Integer
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hWndLock As IntPtr) As Boolean

    Public Structure Node

        Dim PartNumber As String
        Dim Revision As String
        Dim Level As Integer
        Dim FirstLevelIndex As Integer
        Dim ParentIndex As Integer
        Dim Index As Integer
        Dim Description As String

    End Structure

    Public Const APP_PATH As String = "C:\AssemblyTreeBuddy\"
    Public gFilename As String

    Public Const GridX As Integer = 50
    Public Const GridY As Integer = 50

    Public gBitMask(31) As Integer

    Public gPrincipal As WindowsPrincipal = CType(Thread.CurrentPrincipal, WindowsPrincipal)
    Public gComputerName As String
    Public gAppPath As String

    Public gMainForm As frmMain
    Public gTreeForm As frmTree
    Public gAssemblyTreeDataForm As frmAssemblyTreeData

    Public gTotalLines As Integer
    Public gCurrentTopLevel As Integer
    Public gTopLevelBuffer As Integer

    Public gOutputValues() As Node

    Public Const MAX_STRING_LEN As Integer = 82
    

    Public Sub SaveFormWindowToFile(ByRef TheForm As Form)

        Dim sc As New ScreenShot.ScreenCapture

        If Not System.IO.Directory.Exists("c:\Temp") Then
            System.IO.Directory.CreateDirectory("c:\Temp")
        End If

        sc.CaptureWindowToFile(TheForm.Handle, "c:\temp\HMI screenshot - " & TheForm.Name & " " & Format(Now, "yyyy-MM-dd  hh-mm-ss") & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

        sc = Nothing

    End Sub


    Public Sub LogErrorMessage(ByVal Location As String, ByVal Message As String)

        Dim FileNumber As Integer
        Dim Timestamp As String
        Dim MessageForm As frmMessage

        On Error Resume Next

        MessageForm = New frmMessage
        MessageForm.SetMessage(Location & ", " & Message)
        MessageForm.Show()

        Debug.Write(Location & ", " & Message)

        FileNumber = FreeFile()

        Timestamp = GetTimestampString(Now)
        FileOpen(FileNumber, gAppPath & "\error log.txt", OpenMode.Append)
        PrintLine(FileNumber, Timestamp & "  " & Location & ", " & Message)
        FileClose(FileNumber)

    End Sub


    Public Function LogMessageToFile(ByVal sMessage As String, ByVal sFilePath As String) As Boolean

        Dim FileNumber As Integer
        Dim Timestamp As String

        On Error Resume Next

        FileNumber = FreeFile()

        Timestamp = GetTimestampString(Now)
        FileOpen(FileNumber, sFilePath, OpenMode.Append)
        PrintLine(FileNumber, Timestamp & ", " & sMessage)
        FileClose(FileNumber)

        LogMessageToFile = True

    End Function


    Public Function GetTimestampString(ByVal TheDate As Date) As String

        Dim Timestamp As String

        On Error Resume Next

        Timestamp = Format(TheDate, "yyyy-MM-dd") & "  " & Format(TheDate, "HH:mm:ss")

        Return Timestamp

    End Function


    Public Function Distance(ByVal StartX As Double, ByVal StartY As Double, ByVal EndX As Double, ByVal EndY As Double) As Double

        Distance = Math.Sqrt((EndX - StartX) ^ 2 + (EndY - StartY) ^ 2)

    End Function


    Public Function ConvertSecondsToTimeString(ByVal TheSeconds As Integer) As String

        Dim Hours As Integer
        Dim Minutes As Integer
        Dim Seconds As Integer

        Dim Hr As String
        Dim Min As String
        Dim Sec As String

        Hours = Int(TheSeconds / 3600)
        Seconds = TheSeconds - (Hours * 3600)
        Minutes = Int(Seconds / 60)
        Seconds = Seconds - (Minutes * 60)

        Hr = Format(Hours, "00")
        Min = Format(Minutes, "00")
        Sec = Format(Seconds, "00")

        ConvertSecondsToTimeString = Hr & ":" & Min & ":" & Sec

    End Function


    Public Function IsIntegerArrayEmpty(ByRef TheArray() As Integer) As Boolean

        Dim NumItems As Integer

        Try
            NumItems = UBound(TheArray)
            Return False
        Catch ex As Exception
            Return True
        End Try

    End Function


    Public Function IsStringArrayEmpty(ByRef TheArray() As String) As Boolean

        Dim NumItems As Integer

        Try
            NumItems = UBound(TheArray)
            Return False
        Catch ex As Exception
            Return True
        End Try

    End Function


    Public Function IsSingleArrayEmpty(ByRef TheArray() As Single) As Boolean

        Dim NumItems As Integer

        Try
            NumItems = UBound(TheArray)
            Return False
        Catch ex As Exception
            Return True
        End Try

    End Function


    Public Function IsBooleanArrayEmpty(ByRef TheArray() As Boolean) As Boolean

        Dim NumItems As Integer

        Try
            NumItems = UBound(TheArray)
            Return False
        Catch ex As Exception
            Return True
        End Try

    End Function


    Public Function StripNulls(ByVal TheString As String) As String

        If (Strings.InStr(TheString, Strings.Chr(0)) > 0) Then
            TheString = Strings.Left(TheString, Strings.InStr(TheString, Strings.Chr(0)) - 1)
        End If

        Return TheString

    End Function


End Module
