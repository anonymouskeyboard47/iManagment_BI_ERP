Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMEmployerAccounts

#Region "PrivateVariables"

    Private lEmployerID As Long
    Private lBankID As Long
    Private strAccountNumber As String
    Private lAccountTypeID As Long


#End Region

#Region "Properties"

    Public Property EmployerID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lEmployerID
        End Get

        Set(ByVal Value As Long)
            lEmployerID = Value
        End Set

    End Property

    Public Property BankID() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lBankID
        End Get

        Set(ByVal Value As Long)
            lBankID = Value
        End Set

    End Property

    Public Property AccountNumber() As String

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return strAccountNumber
        End Get

        Set(ByVal Value As String)
            strAccountNumber = Value
        End Set

    End Property

    Public Property AccountTypeID() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lAccountTypeID
        End Get

        Set(ByVal Value As Long)
            lAccountTypeID = Value
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

        lEmployerID = 0
        lBankID = 0
        strAccountNumber = ""
        lAccountTypeID = 0

    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Sub Save()
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        If lEmployerID <> 0 And _
                lBankID <> 0 And _
                    Trim(strAccountNumber) <> "" _
                            Then

            strInsertInto = "INSERT INTO EmployerBankAccount (" & _
                "BankID," & _
                "AccountNumber," & _
                "AccountTypeID" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                               ", " & BankID & _
                               ", '" & AccountNumber & _
                               "', " & AccountTypeID & _
                                       ")"


            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Customer's Other Earnings Details Saved")

            Else

                MsgBox("'Save Customer Other Earnings action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Customer's Other Earnings Details Addition Failed")

            End If

        End If

    End Sub

    Public Function Find(ByVal strQuery As String) As Boolean

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

                    lEmployerID = myDataRows("EmployerID").ToString()
                    lBankID = myDataRows("BankID").ToString()
                    strAccountNumber = myDataRows("AccountNumber").ToString()
                    lAccountTypeID = myDataRows("AccountTypeID").ToString()

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strDeleteQuery = strDelQuery

        If lEmployerID = 0 And _
                lBankID <> 0 And _
                 Trim(strAccountNumber) <> "" _
                            Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Employer's Company Bank Account Details Deleted", MsgBoxStyle.Information, _
                    "iManagement - Record Deleted Successfully")
            Else
                MsgBox("'Delete Employer Bank Company Account' action failed", _
                    MsgBoxStyle.Exclamation, " Employer Account Deletion failed")
            End If
        Else
            MsgBox("Cannot Delete. Please select an existing Employer Bank account", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

        End If

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If lEmployerID <> 0 And _
                 lBankID <> 0 And _
                    strAccountNumber _
                        Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Employer's Bank Account Details Updated")
            End If

        End If

    End Sub


#End Region

End Class
