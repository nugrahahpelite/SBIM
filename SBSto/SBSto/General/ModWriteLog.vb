Module ModWriteLog
    Public Sub pbuCreateLogFile(vriFso As Scripting.FileSystemObject, ByRef vriTextStream As Scripting.TextStream, vriUserID As String, vriModuleName As String, vriProcessName As String, vriTransOID As String, ByRef vriLogFileNameOnly As String, ByRef vriLogFileName As String, ByRef vriLogFileNameError As String)
        Dim vnLogFolder As String
        Dim vnProcessDate As String

        vnLogFolder = HttpContext.Current.Server.MapPath("~") & "\" & stuFolderName.WebLog & "\"
        vnLogFolder = HttpContext.Current.Server.MapPath("~") & "\" & stuFolderName.WebLog & "\" & stuFolderName.WebLog & "_" & Format(Date.Now, "yyMMdd") & "\"

        If Dir(vnLogFolder) = "" Then
            MkDir(vnLogFolder)
        End If

        vnProcessDate = Format(Date.Now, "yyMMdd_HHmmss")
        vriLogFileNameOnly = vriModuleName & "_" & vriProcessName & "_" & vriUserID & "_" & vriTransOID & "_" & vnProcessDate
        vriLogFileName = vnLogFolder & vriLogFileNameOnly & ".log"
        vriLogFileNameError = vnLogFolder & vriLogFileNameOnly & "_ERROR.log"

        vriTextStream = Nothing
        vriFso = CreateObject("Scripting.FileSystemObject")
        vriTextStream = vriFso.OpenTextFile(vriLogFileName, Scripting.IOMode.ForWriting, True)

        vriTextStream.WriteLine("Module " & vriModuleName)
        vriTextStream.WriteLine("Proses " & vriProcessName)
        vriTextStream.WriteLine("Process Start      : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vriTextStream.WriteLine("")
    End Sub

    Public Sub pbuCreateDataFile(vriFso As Scripting.FileSystemObject, ByRef vriTextStream As Scripting.TextStream, vriUserID As String, vriModuleName As String, vriProcessName As String, vriTransOID As String, vriFileFolder As String, vriLogFileName As String, ByRef vriDataFileNameOnly As String, ByRef vriDataFileName As String)
        Dim vnDataFolder As String
        Dim vnProcessDate As String

        vnDataFolder = HttpContext.Current.Server.MapPath("~") & "\" & vriFileFolder & "\"
        vnDataFolder = HttpContext.Current.Server.MapPath("~") & "\" & vriFileFolder & "\" & vriFileFolder & "_" & Format(Date.Now, "yyMMdd") & "\"

        If Dir(vnDataFolder) = "" Then
            MkDir(vnDataFolder)
        End If

        vnProcessDate = Format(Date.Now, "yyMMdd_HHmmss")
        vriDataFileNameOnly = vriModuleName & "_" & vriProcessName & "_" & vriUserID & "_" & vriTransOID & "_" & vnProcessDate
        vriDataFileName = vnDataFolder & vriDataFileNameOnly & ".txt"

        vriTextStream = Nothing
        vriFso = CreateObject("Scripting.FileSystemObject")
        vriTextStream = vriFso.OpenTextFile(vriDataFileName, Scripting.IOMode.ForWriting, True)

        vriTextStream.WriteLine("LogFileName = " & vriLogFileName)
        vriTextStream.WriteLine("")
    End Sub
End Module
