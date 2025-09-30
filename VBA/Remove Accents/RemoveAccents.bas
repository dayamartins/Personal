Attribute VB_Name = "Módulo1"
Sub Button_Process_Accents()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim wbNew As Workbook
    Dim rngSource As Range
    Dim row As Long, col As Long
    Dim todayDate As String
    Dim basePath As String
    Dim fullPath As String
    
    ' Define the active sheet as the source
    Set wsSource = ActiveSheet
    
    ' Define the range from A to F up to the last filled row
    row = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row
    Set rngSource = wsSource.Range("A1:F" & row)
    
    ' Create a new workbook
    Set wbNew = Workbooks.Add
    Set wsTarget = wbNew.Sheets(1)
    
    ' Copy the data converting the accents
    For row = 1 To rngSource.Rows.Count
        For col = 1 To rngSource.Columns.Count
            wsTarget.Cells(row, col).Value = RemoveAccents(CStr(rngSource.Cells(row, col).Value))
        Next col
    Next row
    
    ' Get the base path from cell M1 of the active sheet
    basePath = wsSource.Range("M1").Value

    ' Ensure the path ends with "\"
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"
    
    ' Create the folder if it doesn’t exist
    If Dir(basePath, vbDirectory) = "" Then
        MkDir basePath
    End If

    ' Define the file name with today’s date
    todayDate = Format(Date, "yyyy-mm-dd")
    
    ' Combine the path from the cell with the file name
    fullPath = basePath & "blocked_areas_" & todayDate & ".csv"

    ' Save the new workbook as CSV UTF-8
    Application.DisplayAlerts = False ' Prevent overwrite alerts
    wbNew.SaveAs Filename:=fullPath, FileFormat:=xlCSVUTF8
    Application.DisplayAlerts = True
    
    MsgBox "File successfully saved at: " & fullPath
End Sub

Function RemoveAccents(CharText As String) As String
    Dim A As String
    Dim B As String
    Dim i As Integer

    Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
    Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"

    For i = 1 To Len(AccChars)
        A = Mid(AccChars, i, 1)
        B = Mid(RegChars, i, 1)
        CharText = Replace(CharText, A, B)
    Next

    RemoveAccents = CharText
End Function

