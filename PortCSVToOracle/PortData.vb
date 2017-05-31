﻿Imports Oracle.ManagedDataAccess.Client
Imports System.IO
Imports System.Windows.Forms

Module PortData

    Enum Protocols
        TCP
        UDP
    End Enum

    Sub Main()

        Dim HostName As String = "192.168.56.101"
        Dim Port As Integer = 1521
        Dim Protocol As Protocols = Protocols.TCP
        Dim Service As String = "xe"
        Dim Username As String = "hr"
        Dim Password As String = "oracle"
        Dim CSVPath As String = Nothing
        Dim TableName As String = Nothing
        Dim SkipFirstLine As Boolean = True
        Dim GenerateID As Boolean = True
        Dim PrintStatements As Boolean = False

        Dim i As Integer = 0
        Dim OurPath As String = Process.GetCurrentProcess.MainModule.FileName
        For Each CommandArg As String In Environment.GetCommandLineArgs
            i += 1
            Try
                If CommandArg = OurPath Then Continue For 'Fix to keep our CSVPath from being set to our self
                If File.Exists(CommandArg) Then
                    CSVPath = CommandArg
                    Continue For
                End If
                Dim LowerCommandArg As String = CommandArg.ToLower
                If LowerCommandArg.StartsWith("host:") Then
                    HostName = CommandArg.Split(":")(1)
                ElseIf LowerCommandArg.StartsWith("port:") Then
                    Port = CInt(CommandArg.Split(":")(1))
                ElseIf LowerCommandArg.StartsWith("protocol:") Then
                    Protocol = System.Enum.Parse(GetType(Protocols), CommandArg.Split(":")(1))
                ElseIf LowerCommandArg.StartsWith("service:") Then
                    Service = CommandArg.Split(":")(1)
                ElseIf LowerCommandArg.StartsWith("user:") Then
                    Username = CommandArg.Split(":")(1)
                ElseIf LowerCommandArg.StartsWith("password:") Then
                    Password = CommandArg.Split(":")(1)
                ElseIf LowerCommandArg.StartsWith("table:") Then
                    TableName = CommandArg.Split(":")(1)
                ElseIf LowerCommandArg.StartsWith("skipfirst:") Then
                    SkipFirstLine = Boolean.Parse(CommandArg.Split(":")(1))
                ElseIf LowerCommandArg.StartsWith("generateid:") Then
                    GenerateID = Boolean.Parse(CommandArg.Split(":")(1))
                ElseIf LowerCommandArg.StartsWith("print:") Then
                    PrintStatements = Boolean.Parse(CommandArg.Split(":")(1))
                End If
            Catch ex As Exception
                Console.WriteLine("Error! Could not parse parameter " & i & " (" & CommandArg & ")")
                Console.WriteLine(ex.Message)
            End Try
        Next

        If Not File.Exists(CSVPath) Then
            Dim CSVFinder As New OpenFileDialog With {
                .DefaultExt = ".csv",
                .Filter = "Comma Seperated Values|*.csv",
                .InitialDirectory = "%userprofile%",
                .FilterIndex = 0,
                .Title = "Select a CSV file to open"}
            If CSVFinder.ShowDialog = DialogResult.OK Then
                CSVPath = CSVFinder.FileName
            Else
                Exit Sub
            End If
        End If

        Console.WriteLine("The following CSV file will be used:")
        Console.WriteLine(CSVPath)
        Console.WriteLine()

        Console.WriteLine("The following database will be filled with test data:")
        Console.WriteLine("HostName: " & HostName)
        Console.WriteLine("Port: " & Port)
        Console.WriteLine("Protocol: " & Protocol.ToString)
        Console.WriteLine("Service: " & Service)
        Console.WriteLine("User: " & Username)
        Console.WriteLine("Password: " & Password)
        Console.WriteLine("Table: " & TableName)
        Console.WriteLine()
        Console.Write("Press any key to begin...")
        Console.ReadKey()

        Dim DatabaseInfo As String = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=" & Protocol.ToString & ")(HOST=" & HostName & ")(PORT=" & Port & ")))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=" & Service & ")));User Id=" & Username & ";Password=" & Password & ";"
        Dim Database As OracleConnection = Nothing
        Console.WriteLine()
        Console.WriteLine("Attempting to connect with following info:")
        Console.WriteLine(DatabaseInfo)
        Console.WriteLine()

        i = 0
        Dim MaxConnectionAttempts As Integer = 3
        Dim ConnectionSuccess As Boolean = False
        While i < MaxConnectionAttempts And Not ConnectionSuccess
            i += 1
            Try
                Console.Write("Connection Attempt " & i & "\" & MaxConnectionAttempts & " : ")
                Database = New OracleConnection(DatabaseInfo)
                Database.Open()
                ConnectionSuccess = True
                Console.WriteLine("Success :)")
            Catch OrEx As OracleException
                Console.WriteLine("Failed! (Oracle error)")
                PrintOracleException(OrEx)
            Catch ex As Exception
                Console.WriteLine("Failed! (Internal error)")
                Console.WriteLine(ex.Message)
            End Try
            Console.WriteLine()
        End While
        If Not ConnectionSuccess Then
            Console.WriteLine("No database connection could be established within " & MaxConnectionAttempts & " attempts!")
            Console.WriteLine("Press any key to close...")
            Console.ReadKey()
            Exit Sub
        End If

        Using InputReader As New StreamReader(CSVPath)
            If SkipFirstLine Then InputReader.ReadLine() 'Skip first line containing headers
            Dim RowNumber As Integer = 0
            While Not InputReader.EndOfStream
                Dim InsertQuery As String = "INSERT INTO " & TableName & " VALUES ("
                If GenerateID Then InsertQuery &= RowNumber & ","
                Dim LineSplit As String() = InputReader.ReadLine.Split(";")
                For i = 0 To LineSplit.Count - 1
                    If i = 0 Then
                        InsertQuery &= OracleDateTimeString(CSVDateTimeToLocatDateTime(LineSplit(i))) & ","
                        Continue For
                    End If
                    InsertQuery &= "'" & LineSplit(i) & "'"
                    If i < LineSplit.Count - 1 Then InsertQuery &= ","
                Next
                InsertQuery &= ")"
                If PrintStatements Then Console.WriteLine(InsertQuery)
                Try
                    Dim InsertCommand As New OracleCommand(InsertQuery, Database)
                    InsertCommand.ExecuteNonQuery()
                Catch OrEx As OracleException
                    Console.WriteLine("Failed! (Oracle error)")
                    PrintOracleException(OrEx)
                Catch ex As Exception
                    Console.WriteLine("Failed! (Internal error)")
                    Console.WriteLine(ex.Message)
                End Try
                RowNumber += 1
            End While
        End Using

        Console.WriteLine()
        Console.WriteLine("Done!")
        Console.WriteLine("Press enter to close...")
        Console.ReadLine()

    End Sub

    Sub PrintOracleException(OrEx As OracleException)
        Console.WriteLine(OrEx.ErrorCode & " - " & OrEx.Message)
        For Each OrErr As OracleError In OrEx.Errors
            Console.WriteLine(" " & OrErr.Number & " - " & OrErr.Message)
            Console.WriteLine(" " & OrErr.Source)
            Console.WriteLine(" " & OrErr.Procedure)
        Next
    End Sub

    Const OracleDateTimeFormatEquivalent As String = "yyyy/MM/dd HH:mm"
    Const OracleDateTimeFormat As String = "yyyy/mm/dd hh24:mi"

    Function OracleDateTimeString(Input As DateTime) As String
        Return "TO_DATE('" & Input.ToString(OracleDateTimeFormatEquivalent) & "', '" & OracleDateTimeFormat & "')"
    End Function

    Function CSVDateTimeToLocatDateTime(CSVDateTime As String) As DateTime
        Dim SignificantPart As String = CSVDateTime.Split(".")(0)
        Dim SignificantSplit() As String = SignificantPart.Split(" ")
        Dim DateSplit() As String = SignificantSplit(0).Split("-")
        Dim TimeSplit() As String = SignificantSplit(1).Split(":")
        Dim Year As Integer = CInt(DateSplit(0))
        Dim Month As Integer = CInt(DateSplit(1))
        Dim Day As Integer = CInt(DateSplit(2))
        Dim Hour As Integer = CInt(TimeSplit(0))
        Dim Minute As Integer = CInt(TimeSplit(1))
        Dim Second As Integer = CInt(TimeSplit(2))
        Return New DateTime(Year, Month, Day, Hour, Minute, Second)
    End Function

End Module
