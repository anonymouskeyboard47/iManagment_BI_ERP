Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMTransactionTypes

#Region "PrivateVariables"

    Private strTransactionType As String
    Private strDebitCredit As String
    Private strSequenceGroupID As String
    Private MaxValue As Long

#End Region

#Region "Properties"

    Public Property TransactionType() As String

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return strTransactionType
        End Get

        Set(ByVal Value As String)
            strTransactionType = Value
        End Set

    End Property

    Public Property DebitCredit() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strDebitCredit
        End Get

        Set(ByVal Value As String)
            strDebitCredit = Value
        End Set

    End Property

    Public Property SequenceGroupID() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strSequenceGroupID
        End Get

        Set(ByVal Value As String)
            strSequenceGroupID = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Sub NewRecord()

        strTransactionType = ""
        strDebitCredit = ""
        strSequenceGroupID = ""

    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Sub Save()
        'Saves a new sequence group details

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        Try

            If Trim(strTransactionType) = "" And _
                        Trim(strSequenceGroupID) = "" _
                                                        Then

                MsgBox("Please provide the Transaction Type Name" & _
                            " and the Sequence Group ID to save", _
                                    MsgBoxStyle.Critical, "Save Action Failed")
                Exit Sub

            Else

                '[Check if the SequenceGroupID exists

                If Find("SELECT TransactionType FROM SequenceTransaction " & _
                                        " WHERE TransactionType = '" & _
                                                Trim(strTransactionType) & _
                                        "'", False) = False Then

                    'If the sequncegroupID does not exist in sequencemaster then
                    strInsertInto = "INSERT INTO SequenceTransaction (" & _
                        "TransactionType," & _
                        "DebitCredit," & _
                        "SequenceGroupID" & _
                            ") VALUES "

                    strSaveQuery = strInsertInto & _
                            "'" & strTransactionType & _
                            "'," & strDebitCredit & _
                            ",'" & strSequenceGroupID & _
                                    "')"

                    objLogin.connectString = strAccessConnString
                    objLogin.ConnectToDatabase()

                    bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                    strSaveQuery, _
                    datSaved)

                    objLogin.CloseDb()

                    If bSaveSuccess = True Then
                        MsgBox("Transaction Type and accompanying Sequence ID Saved Successfully", _
                            MsgBoxStyle.Information, _
                        "iManagement - Record Saved Successfully")

                    Else

                        MsgBox("'Save Transaction Type' action failed." & _
                            " Make sure all mandatory details are entered", _
                                MsgBoxStyle.Exclamation, _
                                    "iManagement - Sequence Group " & _
                                            "Addition Failed")

                        Exit Sub

                    End If

                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Exclamation, "iManagement System Error")

        End Try

    End Sub

    Public Function Find(ByVal strQuery As String, _
                        ByVal ReturnStatus As Boolean) As Boolean

        Try

            Dim datRetData As DataSet = New DataSet
            Dim bQuerySuccess As Boolean
            Dim myDataTables As DataTable
            Dim myDataColumns As DataColumn
            Dim myDataRows As DataRow
            Dim objLogin As IMLogin = New IMLogin

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bQuerySuccess = objLogin.ExecuteQuery(strAccessConnString, strQuery, _
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

                        'Return a value indicating that the search was not successful
                        Return False
                        Exit Function

                    End If
                    If ReturnStatus = True Then


                        For Each myDataRows In myDataTables.Rows

                            strTransactionType = myDataRows("TransactionType").ToString()
                            strDebitCredit = myDataRows("DebitCredit").ToString()
                            strSequenceGroupID = myDataRows("SequenceGroupID").ToString()

                        Next

                    End If
                Next

                Return True

            Else

                Return False

            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Exclamation, "iManagement System Error")
        End Try

    End Function

    Public Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strDeleteQuery = strDelQuery

        If Trim(strTransactionType) <> "" _
                            Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Sequence Details Deleted", MsgBoxStyle.Information, _
                    "iManagement - Record Deleted Successfully")
            Else
                MsgBox("'Delete Sequence' action failed", _
                    MsgBoxStyle.Exclamation, "Deletion failed")
            End If
        Else
            MsgBox("Cannot Delete. Please select an existing Sequence's Detail", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

        End If

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If Trim(strTransactionType) <> "" _
                             Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Sequence Details Updated")
            End If

        End If

    End Sub

    Public Function FillControl(ByVal strFillConnString As String, _
            ByVal strTSQL As String, ByVal strValueField As String, _
                                ByVal strTextField As String) As String()

        Dim datFillData As DataSet
        Dim bReturnedSuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim strTextFieldData() As String
        Dim i As Integer
        Dim objLogin As IMLogin = New IMLogin

        Try

            datFillData = New DataSet

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            'The db is okay now get the recordset
            bReturnedSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                strTSQL, datFillData)

            objLogin.CloseDb()

            If datFillData Is Nothing Then
                Exit Function
            End If

            For Each myDataTables In datFillData.Tables

                'Check if there is any data. If not exit
                If myDataTables.Rows.Count = 0 Then
                    'Return an empty array
                    ReDim strTextFieldData(1)
                    strTextFieldData(0) = ""
                    Return strTextFieldData

                    Exit Function
                Else
                    'Resize the array
                    ReDim strTextFieldData(myDataTables.Rows.Count)

                End If

                i = 0
                For Each myDataRows In myDataTables.Rows
                    strTextFieldData(i) = myDataRows(0).ToString()
                    i = i + 1
                Next

            Next

            Return strTextFieldData
            datFillData.Dispose()

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnMaxValue(ByVal strQuery As String) As Boolean
        'Query must contain at least rows from Sequence

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.connectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bQuerySuccess = objLogin.ExecuteQuery(strAccessConnString, strQuery, _
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

                    'Return a value indicating that the search was not successful
                    Return False
                    Exit Function

                End If

                For Each myDataRows In myDataTables.Rows

                    MaxValue = myDataRows(0)

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

#End Region

End Class
