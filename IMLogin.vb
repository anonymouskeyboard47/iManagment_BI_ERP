Option Explicit On 
Option Strict Off

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections
Imports System.Web

Public Class IMLogin
    Inherits System.Web.UI.Control
    Private strConnString As String
    Private cn As OleDbConnection = Nothing
    Private mySQL As String = Nothing
    Private myConnectFlag As Boolean = False
    Private SysTrans As OleDbTransaction

#Region "Connection Properties"

    Public Property TheSuccess() As String

        Get
            Return ReturnSuccess
        End Get

        Set(ByVal Value As String)
            ReturnSuccess = Value
        End Set

    End Property


    Public Property TheReturnError() As String

        Get
            Return ReturnError
        End Get

        Set(ByVal Value As String)
            ReturnError = Value
        End Set

    End Property

    Public Property AccessConnectionString() As String

        Get
            Return strAccessConnString
        End Get

        Set(ByVal Value As String)
            strAccessConnString = Value
        End Set

    End Property

    Public Property OrgConnectionString() As String

        Get
            Return strOrgAccessConnString
        End Get

        Set(ByVal Value As String)
            strOrgAccessConnString = Value
        End Set

    End Property

    Public Property OrgConnectionSQLServer() As String

        Get
            Return bOrgConnectionSQLServer
        End Get

        Set(ByVal Value As String)
            bOrgConnectionSQLServer = Value
        End Set

    End Property

    Public Property AccessConnStringADOX() As String

        Get
            Return strAccessConnStringADOX
        End Get

        Set(ByVal Value As String)
            strAccessConnStringADOX = Value
        End Set

    End Property

    Public Property OrgAccessConnStringADOX() As Boolean

        Get
            Return strOrgAccessConnStringADOX
        End Get

        Set(ByVal Value As Boolean)
            strOrgAccessConnStringADOX = Value
        End Set

    End Property

    Public Property SQLConnString() As String

        Get
            Return strSQLConnString
        End Get

        Set(ByVal Value As String)
            strSQLConnString = Value
        End Set

    End Property

    Public Property DBUserName() As String

        Get
            Return strDBUserName
        End Get

        Set(ByVal Value As String)
            strDBUserName = Value
        End Set

    End Property

    Public Property DBPassword() As String

        Get
            Return strDBPassword
        End Get

        Set(ByVal Value As String)
            strDBPassword = Value
        End Set

    End Property

    Public Property DBDatabase() As String

        Get
            Return strDBDatabase
        End Get

        Set(ByVal Value As String)
            strDBDatabase = Value
        End Set

    End Property

    Public Property DBDBPath() As String

        Get
            Return strDBDBPath
        End Get

        Set(ByVal Value As String)
            strDBDBPath = Value
        End Set

    End Property

#End Region

    Public Property ConnectString() As String

        'USED TO SET AND RETRIEVE THE CONNECTION STRING VALUE
        Get

            Return strConnString

        End Get

        Set(ByVal Value As String)
            strConnString = Value
        End Set

    End Property

    Public Sub New()
        MyBase.New()
        ReturnError = ""
    End Sub

    Public Sub New(ByRef strConnect As String)
        MyBase.New()
        strConnString = strConnect
        ReturnError = ""
    End Sub

    Public Function ConnectToDatabase() As Boolean


        Try

            cn = New OleDbConnection
            cn.ConnectionString = strConnString
            cn.Open()
            myConnectFlag = True
            Return True

        Catch ex As Exception
            ReturnError += " -- Cannot Connect to System Data. Invalid Data Presented for Login!" & vbCrLf & ex.Message
            Return False
        End Try

    End Function

    Public Sub CloseDb()
        'Close the database connection
        If myConnectFlag Then
            cn.Close()
        End If

    End Sub

    Private Function ExecuteSQL(ByRef mySQL As String, _
            ByRef myDataSet As DataSet, _
                ByRef mySQLError As String) As Boolean

        Dim mylocalDataSet As DataSet = New DataSet
        Dim mylocalError As String = Nothing
        Dim mySuccessFlag As Boolean = False
        Dim myCommand As OleDbCommand = Nothing
        Dim myDataSetCommand As OleDbDataAdapter = Nothing

        Try

            If (myConnectFlag) Then

                myCommand = New OleDbCommand(mySQL, cn)

                'If Not SysTrans Is Nothing Then
                '    myCommand.Transaction = SysTrans
                'End If

                myDataSetCommand = New OleDbDataAdapter(myCommand)
                myDataSetCommand.Fill(mylocalDataSet, "temp")
                mySuccessFlag = True

            Else
                CloseDb()
                ReturnError += " -- iManagement Connection: " & _
                    "[Connection Failed]: Invalid Connection parameters. " & _
                    cn.Provider & Chr(10) & _
                    cn.State & Chr(13) & _
                    cn.Database & Chr(13) & _
                    cn.DataSource & Chr(13)

            End If

        Catch e As OleDbException
            CloseDb()
            Dim myError As OleDbError

            For Each myError In e.Errors
                ReturnError += " -- iManagement Connection Execution :" & Chr(10) & _
                    " [Error Source]: " & myError.Source & Chr(10) & _
                        " [Error Message]: " & myError.Message & Chr(10) & _
                                myCommand.CommandText & Chr(10) & _
                                myCommand.CommandType & _
                    cn.Provider & Chr(10) & _
                    cn.State & Chr(13) & _
                    cn.Database & Chr(13) & _
                    cn.DataSource & Chr(13)


            Next

            SysTrans.Rollback()

        Finally

            mySQLError = mylocalError
            myDataSet = mylocalDataSet

        End Try

        Return mySuccessFlag

    End Function

    Public Function ExecuteQuery _
      (ByVal connString As String, _
          ByVal strQuery As String, _
              ByRef retDataSet As DataSet) As Boolean
        Dim mySuccessflsg As Boolean
        Dim myDataSet As DataSet = Nothing
        Dim strSQL As String
        Dim mySQLError As String
        Dim myQueryflag As Boolean
        Dim mylocalError As String = Nothing

        Try

            ReturnError = ""

            strSQL = strQuery
            myQueryflag = ExecuteSQL(strSQL, myDataSet, mySQLError)

            'ReturnError = "AccessCompleted"

        Catch e As OleDbException
            CloseDb()
            Dim myError As OleDbError

            For Each myError In e.Errors
                ReturnError += " -- iManagement Execution : [Error Source]: " + _
                    myError.Source + "[Error Message]: " + _
                        myError.Message

            Next

        Finally
            If (myQueryflag) Then

                retDataSet = myDataSet
                ExecuteQuery = True

            End If
        End Try

    End Function

    Public Function BeginTheTrans() As Boolean
        bCommadTransactionInitiate = True
        bCommadTransactionStartedState = False

    End Function

    Public Function CommitTheTrans() As Boolean
        If Not SysTrans.Connection Is Nothing Then
            SysTrans.Commit()
            bCommadTransactionStartedState = False
            bCommandTransactionCompleteState = True
        End If

    End Function

    Public Function RollbackTheTrans() As Boolean
        If Not SysTrans.Connection Is Nothing Then
            SysTrans.Rollback()
        End If

    End Function

    '[Used to write sequential files as a stream
    Private Function WriteSequentialFile _
        (ByVal strFileString As String, _
            ByVal strFileName As String, _
                ByVal DisplayErrorMessages As Boolean) _
                    As Boolean
        Try

            'Write the file to disk sequentially
            Dim sr As StreamWriter
            Dim Contents As String
            Dim flHandle As Long

            flHandle = FreeFile()

            If flHandle = 0 Then
                Return False
                Exit Function
            End If

            sr = New StreamWriter(strFileName)
            sr.Write(strFileString)
            sr.Close()

            'FileOpen(flHandle, strFileName, OpenMode.Output)
            'Print(flHandle, strFileString)

            'FileClose(flHandle)
            ''MsgBox("File closed")

            Return True

        Catch ex As Exception

            If DisplayErrorMessages = True Then
                MsgBox(ex.Message)

            End If

        End Try

    End Function

    Public Function ReadSequentialFile _
       (ByVal strFileName As String, _
               ByVal DisplayErrorMessages As Boolean) _
                   As String
        Try

            'Write the file to disk sequentially
            Dim sr As StreamReader
            Dim Contents As String
            Dim flHandle As Long

            flHandle = FreeFile()

            If flHandle = 0 Then
                Return False
                Exit Function
            End If

            sr = New StreamReader(strFileName)
            Contents = sr.ReadToEnd()
            sr.Close()

            Return Contents

        Catch ex As Exception

            If DisplayErrorMessages = True Then
                MsgBox(ex.Message)

            End If

        End Try

    End Function

    Public Function FillArray(ByVal strFillConnString As String, _
            ByVal strTSQL As String, ByVal strValueField As String, _
                ByVal strTextField As String, _
                    Optional ByVal lNumOfColumns As Long = 0, _
    Optional ByVal bUseRecordset _
                            As Boolean = False) As Object

        Dim datFillData As DataSet
        Dim bReturnedSuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim strTextFieldData() As String
        Dim i As Integer
        Dim objLogin As IMLogin = New IMLogin
        Dim strTextFieldData2(,) As String
        Dim strTextFieldData3(,,) As String
        Dim strTextfieldData6(,,,,,) As String

        Try

            ReturnError = ""

            datFillData = New DataSet

            objLogin.ConnectString = strFillConnString
            objLogin.ConnectToDatabase()

            'The db is okay now get the recordset
            bReturnedSuccess = objLogin.ExecuteQuery(strFillConnString, _
                strTSQL, datFillData)

            objLogin.CloseDb()

            If datFillData Is Nothing Then
                Exit Function
            End If

            If bUseRecordset = True Then
                Return datFillData
                Exit Function
            End If

            For Each myDataTables In datFillData.Tables

                'Check if there is any data. If not exit
                If myDataTables.Rows.Count = 0 Then
                    datFillData = Nothing
                    objLogin = Nothing
                    Return Nothing

                    Exit Function
                Else

                    'Resize the array
                    If lNumOfColumns = 0 Then
                        ReDim strTextFieldData(myDataTables.Rows.Count)

                    ElseIf lNumOfColumns = 2 Then
                        ReDim strTextFieldData2 _
                            (myDataTables.Rows.Count, myDataTables.Rows.Count)

                    ElseIf lNumOfColumns = 3 Then
                        ReDim strTextFieldData3 _
                            (myDataTables.Rows.Count, myDataTables.Rows.Count, _
                                 myDataTables.Rows.Count)

                    ElseIf lNumOfColumns = 6 Then
                        ReDim strTextfieldData6 _
                            (myDataTables.Rows.Count, _
                                myDataTables.Rows.Count, _
                                    myDataTables.Rows.Count, _
                                        myDataTables.Rows.Count, _
                                    myDataTables.Rows.Count, _
                                myDataTables.Rows.Count)

                    End If

                End If

                i = 0
                For Each myDataRows In myDataTables.Rows
                    If lNumOfColumns = 0 Then
                        strTextFieldData(i) = myDataRows(0).ToString()

                    ElseIf lNumOfColumns = 2 Then
                        strTextFieldData2(i, 0) = myDataRows(0).ToString()
                        strTextFieldData2(i, 1) = myDataRows(1).ToString()

                    ElseIf lNumOfColumns = 3 Then
                        strTextFieldData3(i, 0, 0) = myDataRows(0).ToString()
                        strTextFieldData3(i, 1, 0) = myDataRows(1).ToString()
                        strTextFieldData3(i, 0, 1) = myDataRows(2).ToString()

                    ElseIf lNumOfColumns = 6 Then
                        strTextfieldData6(i, 0, 0, 0, 0, 0) = myDataRows(0).ToString()
                        strTextfieldData6(i, 1, 0, 0, 0, 0) = myDataRows(1).ToString()
                        strTextfieldData6(i, 0, 1, 0, 0, 0) = myDataRows(2).ToString()
                        strTextfieldData6(i, 0, 0, 1, 0, 0) = myDataRows(3).ToString()
                        strTextfieldData6(i, 0, 0, 0, 1, 0) = myDataRows(4).ToString()
                        strTextfieldData6(i, 0, 0, 0, 0, 1) = myDataRows(5).ToString()

                    End If

                    i = i + 1

                Next
            Next

            myDataTables = Nothing
            myDataRows = Nothing
            myDataColumns = Nothing

            objLogin = Nothing
            datFillData.Dispose()

            If lNumOfColumns = 0 Then
                Return strTextFieldData

            ElseIf lNumOfColumns = 2 Then
                Return strTextFieldData2

            ElseIf lNumOfColumns = 3 Then
                Return strTextFieldData3

            ElseIf lNumOfColumns = 6 Then
                Return strTextfieldData6
            End If



        Catch ex As Exception
            ReturnError += " -- TripodSystems iManagement Data Access Error. " & _
                "Please contact your administrator (" & _
                    ex.Message & ")"

        End Try

    End Function

    Public Function FillDataset(ByVal strFillConnString As String, _
         ByVal strTSQL As String, ByVal strValueField As String, _
             ByVal strTextField As String, _
                 Optional ByVal lNumOfColumns As Long = 0) As Object

        Dim datFillData As DataSet
        Dim bReturnedSuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim strTextFieldData() As String
        Dim i As Integer
        Dim objLogin As IMLogin = New IMLogin
        Dim strTextFieldData2(,) As String
        Dim strTextFieldData3(,,) As String

        Try

            datFillData = New DataSet

            objLogin.ConnectString = strFillConnString
            objLogin.ConnectToDatabase()

            'The db is okay now get the recordset
            bReturnedSuccess = objLogin.ExecuteQuery(strFillConnString, _
                strTSQL, datFillData)

            objLogin.CloseDb()

            If datFillData Is Nothing Then
                Exit Function
            End If

            Return datFillData

            datFillData = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Function

    'Used to return the Max value of a query. The return value is a
    'long value
    Public Function ReturnMaxLongValue _
        (ByVal strValConnString As String, _
            ByVal strQuery As String) As Long

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin
        Dim MaxValue As Long

        objLogin.ConnectString = strValConnString
        objLogin.ConnectToDatabase()

        bQuerySuccess = objLogin.ExecuteQuery(strValConnString, strQuery, _
                                                datRetData)

        objLogin.CloseDb()

        If datRetData Is Nothing Then
            Exit Function
        End If

        If bQuerySuccess = True Then

            Dim i As Integer

            For Each myDataTables In datRetData.Tables

                'Check if there is any data. If not exit
                If myDataTables.Rows.Count = 0 Then
                    datRetData = Nothing
                    objLogin = Nothing
                    'Return a value indicating that the search was not successful
                    Return 0
                    Exit Function

                End If

                For Each myDataRows In myDataTables.Rows

                    If IsDBNull(myDataRows(0)) = False Then
                        MaxValue = myDataRows(0)
                    Else
                        Return 0
                    End If

                Next

            Next

            Return MaxValue
        Else
            Return MaxValue
        End If


    End Function

    Public Function FindHasRecords(ByVal strValConnecString As String, _
        ByVal strQuery As String) As Boolean

        Try

            Dim datRetData As DataSet = New DataSet
            Dim bQuerySuccess As Boolean
            Dim myDataTables As DataTable
            Dim myDataColumns As DataColumn
            Dim myDataRows As DataRow
            Dim objLogin As IMLogin = New IMLogin

            objLogin.ConnectString = strValConnecString
            objLogin.ConnectToDatabase()

            'ReturnError = "AccessDone" & strValConnecString

            bQuerySuccess = objLogin.ExecuteQuery(strValConnecString, _
                            strQuery, datRetData)

            'ReturnError = "AccessInSession" & strAccessConnString

            'If ReturnError = "" Then
            'ReturnError = "Access completed and not successful"
            'End If


            objLogin.CloseDb()

            If datRetData Is Nothing Then
                Exit Function
            End If

            If bQuerySuccess = True Then

                Dim i As Integer

                For Each myDataTables In datRetData.Tables

                    'Check if there is any data. If not exit
                    If myDataTables.Rows.Count = 0 Then
                        datRetData = Nothing
                        objLogin = Nothing
                        'Return a value indicating that the search was not successful
                        Return False
                        Exit Function

                    End If

                    Return True

                Next
            Else
                Return False
            End If

        Catch ex As Exception
            ReturnError += " -- " & ex.Message & "-" & ex.Source
        End Try

    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        SysTrans = Nothing
    End Sub

    Private Sub objTimer_Tick(ByVal sender As Object, _
        ByVal e As System.EventArgs)

    End Sub

End Class

