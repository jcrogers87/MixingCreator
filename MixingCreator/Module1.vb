'\\Plataine 2015
'\\searches a CSV for a user specified column inside it and adds a new column with a mixing group
Imports System.IO
Module Module1
    'readconfig globals
    Dim inputsheets() As String, sortField As Integer, headers As Boolean, maximumSize As Integer, delimiterCol As Integer
    'globals
    Dim columnHeaders As String
    Sub Main()
        ReadConfig()

        For Each input As String In inputsheets
            If Path.GetExtension(input) = ".csv" Then
                ReadInput(input)
            End If
        Next

    End Sub
    Public Sub ReadInput(file As String)
        Try
            Dim SR As StreamReader = New StreamReader(file)
            Dim line As String = SR.ReadLine()
            Dim strArray As String() = line.Split(",")
            Dim data As DataTable = New DataTable()
            Dim row As DataRow

            For Each s As String In strArray
                data.Columns.Add(New DataColumn())
            Next

            'write everything to a datatable
            Do
                If Not line = String.Empty Then
                    row = data.NewRow()
                    row.ItemArray = line.Split(",")
                    data.Rows.Add(row)
                Else
                    Exit Do
                End If
                line = SR.ReadLine
            Loop
            SR.Close()

            'delete the first row if headers and then store the header information in a new array called columnheaders
            If headers = True Then
                data.Rows(0).Delete()
                GetHeaders(file)
            Else
                columnHeaders = Nothing
            End If

            'Add unique objects in sortfield into an array called uniqueField
            Dim view As New DataView(data)
            view.Sort = data.Columns(sortField - 1).ToString
            Dim uniqueValues = view.ToTable(True, data.Columns(sortField - 1).ColumnName)
            data = view.ToTable()

            uniqueValues.Columns.Add("count")
            uniqueValues.Columns.Add("split")

            'count Number of entires per sortfield And calculate max
            Dim i As Integer
            For x As Integer = 0 To uniqueValues.Rows.Count - 1
                i = 0
                For y As Integer = 0 To data.Rows.Count - 1
                    If uniqueValues.Rows(x).Item(data.Columns(sortField - 1).ColumnName) = data.Rows(y).Item(sortField - 1) Then
                        i = i + 1
                        uniqueValues.Rows(x).Item("count") = i
                    End If
                Next
            Next

            'calculate the split number
            For x = 0 To uniqueValues.Rows.Count - 1
                Dim maxJobs As Integer = maximumSize
                If maxJobs = Nothing Or 0 Then maxJobs = uniqueValues.Rows(x).Item("count")
                If uniqueValues.Rows(x).Item("count") / maxJobs > CInt(uniqueValues.Rows(x).Item("count") / maxJobs) And uniqueValues.Rows(x).Item("count") > maxJobs Then
                    maxJobs = CInt(Math.Ceiling(uniqueValues.Rows(x).Item("count") / Math.Ceiling(uniqueValues.Rows(x).Item("count") / maxJobs)))
                End If
                uniqueValues.Rows(x).Item("split") = maxJobs
            Next

            'add the new mixing group
            data.Columns.Add("mixingGroup")
            Dim mixingString As String
            For x = 0 To uniqueValues.Rows.Count - 1
                i = 0
                Dim j As Integer = 0
                For y = 0 To data.Rows.Count - 1
                    If data.Rows(y).Item(data.Columns(sortField - 1).ColumnName) = uniqueValues.Rows(x).Item(data.Columns(sortField - 1).ColumnName) Then
                        j = j + 1
                        mixingString = data.Rows(y).Item(delimiterCol - 1) & "_" & uniqueValues.Rows(x).Item(data.Columns(sortField - 1).ColumnName) & "-" & i.ToString
                        data.Rows(y).Item("mixingGroup") = mixingString
                        If j = uniqueValues.Rows(x).Item("split") Then
                            i = i + 1
                            j = 0
                        End If
                    End If
                Next
            Next

            WriteMe(data, file)
        Catch ex As Exception
            Call MsgBox("An error has occurred. Contact Plataine" & Chr(13) & ex.ToString)
        End Try
    End Sub
    Public Sub GetHeaders(input As String)
        If Path.GetExtension(input) = ".csv" Then
            Dim headerString() As String = File.ReadAllLines(input)
            columnHeaders = headerString(0)
        End If
    End Sub
    Public Sub WriteMe(ByVal output As DataTable, ByVal fileName As String)
        If File.Exists(fileName) Then
            File.Delete(fileName)
        End If
        Dim delim As String = Nothing, first As Boolean
        first = True
        Dim sw As New StreamWriter(fileName, True)
        Dim builder As New System.Text.StringBuilder
        For Each row As DataRow In output.Rows
            delim = ""
            If Not first Then builder.AppendLine()
            first = False
            For Each col As DataColumn In output.Columns
                builder.Append(delim)
                delim = ","
                builder.Append(row(col.ColumnName))
            Next
        Next
        sw.WriteLine(columnHeaders & ",MixingGroup")
        sw.WriteLine(builder.ToString())
        sw.Close()
    End Sub
    Public Sub ReadConfig()
        If Not File.Exists("C:\ProgramData\Plataine\MixingCreator.config") Then
            BuildConfig()
        End If
        maximumSize = Nothing
        delimiterCol = 1
        Try
            headers = False
            Dim configFile() As String = File.ReadAllLines("C:\ProgramData\Plataine\MixingCreator.config")
            For Each line As String In configFile
                Dim setting() As String = Split(line, "=")
                If UCase(setting(0)) = "INPUTFOLDER" Then
                    inputsheets = Directory.GetFiles(setting(1))
                ElseIf UCase(setting(0)) = "HEADERS" Then
                    If UCase(setting(1).ToString) = "TRUE" Then headers = True Else headers = False
                ElseIf UCase(setting(0)) = "SORTCOLUMN" Then
                    sortField = CInt(setting(1).ToString)
                ElseIf UCase(setting(0)) = "MAX" Then
                    maximumSize = (setting(1).ToString)
                ElseIf UCase(setting(0)) = "DELIMETERCOLUMN" Then
                    delimiterCol = CInt(setting(1).ToString)
                End If
            Next
            If IsNothing(inputsheets) Or IsNothing(sortField) Or IsNothing(maximumSize) Then
                Call MsgBox("Your config file is invalid. Must be of the form:" _
                           & Chr(13) & "inputfile=pathtojobs" _
                           & Chr(13) & "SortColumn=integer" _
                           & Chr(13) & "max=integer" _
                           & Chr(13) & "delimetercolumn=integer" _
                           & Chr(13) & "Config location must be: C:\ProgramData\Plataine\MixingCreator.config")
                End
            End If
        Catch ex As Exception
            Call MsgBox("Your config file is missing, or missing required column mappings." _
                               & Chr(13) & "Config location must be: C:\ProgramData\Plataine\MixingCreator.config" _
                               & Chr(13) & "Delete existing config files to generate a fresh one.")
            End
        End Try
    End Sub
    Public Sub BuildConfig()
        If (Not Directory.Exists("C:\ProgramData\Plataine\")) Then
            Directory.CreateDirectory("C:\ProgramData\Plataine\")
        End If
        Dim sw As New StreamWriter("C:\ProgramData\Plataine\MixingCreator.config", False)
        sw.WriteLine("##Auto Config##")
        sw.WriteLine("inputfolder=C:\TPO\InputFile")
        sw.WriteLine("sortColumn=9")
        sw.WriteLine("headers=true")
        sw.WriteLine("max=500")
        sw.WriteLine("delimetercolumn=1")
        sw.Close()
    End Sub
End Module
