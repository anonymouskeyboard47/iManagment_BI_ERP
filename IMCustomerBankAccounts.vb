Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCustomerBankAccounts
    Inherits IMBank


#Region "PrivateVariables"

    Private lCustomerNo As Long
    Private strAccountNumber As String
    Private strAccountType As String
    Private lAccountID As Long

#End Region

#Region "Properties"

    Public Property AccountID() As Long

        Get
            Return lAccountID
        End Get

        Set(ByVal Value As Long)
            lAccountID = Value
        End Set

    End Property

    Public Property CustomerNo() As Long

        Get
            Return lCustomerNo
        End Get

        Set(ByVal Value As Long)
            lCustomerNo = Value
        End Set

    End Property

    Public Property AccountNumber() As String

        Get
            Return Trim(strAccountNumber)
        End Get

        Set(ByVal Value As String)
            strAccountNumber = Value
        End Set

    End Property

    Public Property AccountType() As String

        Get
            Return Trim(strAccountType)
        End Get

        Set(ByVal Value As String)
            strAccountType = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    'Public Sub New()
    '    MyBase.New()
    'End Sub

#End Region

#Region "GeneralProcedures"

    Public Sub New()
        MyBase.New()
    End Sub


#End Region

#Region "DatabaseProcedures"

    Public Shadows Sub Save()
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        If lCustomerNo <> 0 Or _
            Trim(strAccountNumber) = "" Or _
                BankID = 0 _
                            Then

            MsgBox("Please provide an existing" & _
            Chr(10) & "1. Customer Number" & _
            Chr(10) & "2. Account number" & _
            Chr(10) & "3. Existing Bank." _
                           , MsgBoxStyle.Exclamation, _
                           "iManagement - invalid or incomplete information")

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub

        End If

        'Check if there is an existing series with this name
        If Find("SELECT * FROM CustomerBankAccounts WHERE  BankID = " _
                    & BankID & " AND CustomerNo = " & _
                        lCustomerNo & " AND AccountNumber = '" & _
                            Trim(strAccountNumber) & "'", False) = True Then

            If MsgBox("The Customer's Bank Account Details already exists." & _
            Chr(10) & "Do you want to update the details?", _
                    MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                            MsgBoxResult.Yes Then

                Update("UPDATE CustomerBankAccounts SET " & _
                    "CustomerNo = " & lCustomerNo & _
                            " AND BankID = " & BankID & _
                            " AND AccountNumber = '" & Trim(strAccountNumber) & _
                                "' WHERE  BankID = " _
                                    & BankID & " AND CustomerNo = " & _
                                        lCustomerNo & " AND AccountNumber = '" & _
                                            Trim(strAccountNumber) & "'")

            End If

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If

        'Check if there is an existing series with this name
        If Find("SELECT * FROM CustomerBankAccounts WHERE AccountID = " & _
                    lAccountID, False) = True Then

            If MsgBox("The Customer's Bank Account Details already exists." & _
            Chr(10) & "Do you want to update the details?", _
                    MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                            MsgBoxResult.Yes Then

                Update("UPDATE CustomerBankAccounts SET " & _
                    "CustomerNo = " & lCustomerNo & _
                            ", BankID = " & BankID & _
                            ", AccountNumber = '" & Trim(strAccountNumber) & _
                                "' WHERE AccountID = " & _
                    lAccountID)

            End If

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If


        strInsertInto = "INSERT INTO CustomerBankAccounts (" & _
            "CustomerNo," & _
            "BankID," & _
            "AccountNumber," & _
            "AccountType" & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                "(" & lCustomerNo & _
                "," & BankID & _
                ",'" & strAccountNumber & _
                "','" & strAccountType & _
                        "')"

        objLogin.connectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
        strSaveQuery, _
        datSaved)

        objLogin.CloseDb()

        If bSaveSuccess = True Then
            MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
            "iManagement - Customer's Bank Account Details Saved")

        Else

            MsgBox("'Save Customer Bank Account' action failed." & _
                " Make sure all mandatory details are entered", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Customer's Bank Account Addition Failed")

        End If


    End Sub

    Public Shadows Function Find(ByVal strQuery As String, _
            ByVal bReturnValues As Boolean) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.connectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bQuerySuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strQuery, _
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

                For Each myDataRows In myDataTables.Rows
                    If bReturnValues = True Then
                        lAccountID = myDataRows("AccountID")
                        lCustomerNo = myDataRows("CustomerNo")
                        BankID = myDataRows("BankID")
                        strAccountNumber = myDataRows("AccountNumber").ToString()
                        strAccountType = myDataRows("AccountType").ToString()
                    End If

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Shadows Sub Delete()

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        If lAccountID = 0 Then
            MsgBox("Cannot Delete due to missing information. Please provide an existing" & _
            Chr(10) & "1. Customer Bank Account by selecting the Account ID." _
                           , MsgBoxStyle.Exclamation, _
                           "iManagement - invalid or incomplete information")
            objLogin = Nothing
            datDelete = Nothing

            Exit Sub

        End If

        strDeleteQuery = "DELETE * FROM CustomerBankAccounts WHERE AccountID = " & _
                    lAccountID

        objLogin.connectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
        datDelete)

        objLogin.CloseDb()

        If bDelSuccess = True Then
            MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                "iManagement - Customer's Bank Account Details Deleted")
        Else
            MsgBox("'Delete Customer's Bank Account' action failed", _
                MsgBoxStyle.Exclamation, " Customer Bank Account Deletion failed")
        End If

    End Sub

    Public Shadows Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If (lCustomerNo <> 0 And _
                 Trim(strAccountNumber) <> 0 And _
                    BankID <> 0) Or (lAccountID <> 0) Then

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Customer's Bank Account Details Updated")
            End If

        End If

    End Sub


#End Region



End Class
