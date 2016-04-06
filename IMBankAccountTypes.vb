Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMBankAccountTypes


#Region "PrivateVariables"
    Private lAccountTypeID As Long
    Private strAccountTypeTitle As String
    Private strAccountTypeDescription As String

#End Region


#Region "Properties"

    Public Property AccountTypeID() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lAccountTypeID
        End Get

        Set(ByVal Value As Long)
            lAccountTypeID = Value
        End Set

    End Property

    Public Property AccountTypeTitle() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strAccountTypeTitle
        End Get

        Set(ByVal Value As String)
            strAccountTypeTitle = Value
        End Set

    End Property

    Public Property AccountTypeDescription() As String

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return strAccountTypeDescription
        End Get

        Set(ByVal Value As String)
            strAccountTypeDescription = Value
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

        lAccountTypeID = 0
        strAccountTypeTitle = ""
        strAccountTypeDescription = ""

    End Sub

#End Region


#Region "DatabaseProcedures"

    Public Function Save() As Boolean
        'Saves a new country name

        Try

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        If Trim(strAccountTypeTitle) = 0 Then
            MsgBox("Please provide the Bank Account Type in order to save it." _
            , MsgBoxStyle.Exclamation, _
                "iManagement - invalid or incomplete information")

            datSaved = Nothing
            objLogin = Nothing
            Exit Function

        End If


        If Find("SELECT * FROM BankAccountTypes " & _
        " WHERE AccountTypeID = " & lAccountTypeID, False) = True Then



            datSaved = Nothing
            objLogin = Nothing
            Exit Function
        End If

        strInsertInto = "INSERT INTO BankAccountTypes (" & _
            "AccountTypeTitle," & _
            "AccountTypeDescription," & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                "'" & Trim(strAccountTypeTitle) & _
                "', '" & Trim(strAccountTypeDescription) & _
                        "')"

        objLogin.ConnectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
        strSaveQuery, _
        datSaved)

        objLogin.CloseDb()

        If bSaveSuccess = True Then
            MsgBox("Bank Account Types Lookup Details Saved", MsgBoxStyle.Information, _
            "iManagement - Record Saved Successfully")
                Return True
        Else

            MsgBox("'Save Bank Account Type action failed." & _
                " Make sure all mandatory details are entered", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Addition of Bank account Type Failed")

        End If

        datSaved = Nothing
        objLogin = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function Find(ByVal strQuery As String, _
        ByVal bReturnValues As Boolean) As Boolean

        Try
            Dim datRetData As DataSet = New DataSet
            Dim bQuerySuccess As Boolean
            Dim myDataTables As DataTable
            Dim myDataColumns As DataColumn
            Dim myDataRows As DataRow
            Dim objLogin As IMLogin = New IMLogin

            objLogin.ConnectString = strAccessConnString
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
                        datRetData = Nothing
                        objLogin = Nothing

                        Exit Function

                    End If

                    If bReturnValues = True Then
                        For Each myDataRows In myDataTables.Rows

                            lAccountTypeID = myDataRows("AccountTypeID")
                            strAccountTypeTitle = myDataRows _
                                    ("AccountTypeTitle").ToString()
                            strAccountTypeDescription = myDataRows _
                                    ("AccountTypeDescription").ToString()

                        Next
                    End If

                Next

                Return True
            Else
                Return False
        End If

        Catch ex As Exception

        End Try
    End Function

    Public Function Delete() As Boolean
        Try
            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = "DELETE * FROM BankAccountTypes WHERE AccounTypeID = " & lAccountTypeID

            If lAccountTypeID = 0 Then
                MsgBox("Cannot Delete. Please select an existing Bank Account Type Detail", _
                         MsgBoxStyle.Exclamation, _
                         "iManagement - invalid or incomplete Information")
                datDelete = Nothing
                objLogin = Nothing
                Exit Function

            End If
            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Bank Account Type Details Deleted")
                Return True
            Else
                MsgBox("'Delete Bank Account Type' action failed", _
                    MsgBoxStyle.Exclamation, " Bank Account Type Deletion failed")
            End If

            datDelete = Nothing
            objLogin = Nothing


        Catch ex As Exception

        End Try

    End Function

    Public Function Update(ByVal strUpQuery As String) As Boolean

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery



            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Bank Account Type Details Updated")
            End If
        Catch ex As Exception

        End Try

    End Function

#End Region


End Class



