Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMEmployeeBankAccounts
    Inherits IMBank


#Region "PrivateVariables"

    Private lEmployerID As Long
    Private lAccountID As Long
    Private strAccountNumber As String
    Private strAccountType As String

#End Region

#Region "Properties"

    Public Property EmployerID() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lEmployerID
        End Get

        Set(ByVal Value As Long)
            lEmployerID = Value
        End Set

    End Property

    Public Property AccountID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lAccountID
        End Get

        Set(ByVal Value As Long)
            AccountID = Value
        End Set

    End Property

    Public Property AccountNumber() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strAccountNumber
        End Get

        Set(ByVal Value As String)
            strAccountNumber = Value
        End Set

    End Property

    Public Property AccountType() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strAccountType
        End Get

        Set(ByVal Value As String)
            strAccountType = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    'Public Sub New()
    'MyBase.New()
    'End Sub

#End Region

#Region "GeneralProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

    Public Shadows Sub NewRecord()

        lEmployerID = 0
        AccountID = 0
        BankID = 0
        AccountNumber = ""
        AccountType = ""

    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Shadows Function Save() As Boolean
        'Saves a new country name
        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String

            If lEmployerID = 0 Or _
                Trim(strAccountNumber) = "" Or _
                BankID = 0 _
                                Then
                MsgBox("Cannot save the details due to missing information." & _
                Chr(10) & "1. Select an existing employer" & _
                Chr(10) & "2. The text for the Bank Account Number" & _
                Chr(10) & "3. Select an existing employer" & _
                    MsgBoxStyle.Exclamation, _
                    "iManagement - invalid or incomplete information")

                datSaved = Nothing
                objLogin = Nothing
                Exit Function
            End If

            If Find("SELECT * FROM EmployerSalaryAccounts WHERE " & _
            " (EmployerID = " & lEmployerID & _
                " AND AccountNr = '" & strAccountNumber & _
                "' AND BankID = " & BankID & _
                ") OR (AccountID = " & lAccountID & ")", False) = True Then

                If MsgBox("The Employer's Account Number exists. Do you wish" & _
                " to update it?", _
                    MsgBoxStyle.YesNo, _
                    "iManagement - Update Record?") = MsgBoxResult.Yes Then

                    Return Update("UPDATE EmployerBankAccounts SET " & _
                    " AccountNr = '" & strAccountNumber & "'" & _
                    ", AccountType = '" & strAccountType & "'" & _
                    " WHERE AccountNr = '" & strAccountNumber & _
                    "' AND BankID = " & BankID & _
                    " AND EmployerID = " & lEmployerID)

                End If

                datSaved = Nothing
                objLogin = Nothing
                Exit Function
            End If

            strInsertInto = "INSERT INTO EmployerSalaryAccount (" & _
                "EmployerID," & _
                "AccountID," & _
                "BankID," & _
                "AccountNr," & _
                "AccountType" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lEmployerID & _
                    "," & lAccountID & _
                    "," & BankID & _
                    ",'" & Trim(strAccountNumber) & _
                    "','" & strAccountType & _
                            "')"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Customer's Bank Account Details Saved")
                Return True
            Else

                MsgBox("'Save Customer Bank Account' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Customer's Bank Account Addition Failed")

            End If

        Catch ex As Exception

        End Try

    End Function

    Public Shadows Function Find(ByVal strQuery As String, _
            ByVal bReturnDetails As Boolean) As Boolean

        Try

            Dim datRetData As DataSet = New DataSet
            Dim bQuerySuccess As Boolean
            Dim myDataTables As DataTable
            Dim myDataColumns As DataColumn
            Dim myDataRows As DataRow
            Dim objLogin As IMLogin = New IMLogin

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bQuerySuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strQuery, _
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

                    If bReturnDetails = True Then
                        For Each myDataRows In myDataTables.Rows

                            lEmployerID = myDataRows("EmployerID")
                            lAccountID = myDataRows("AccountID")
                            BankID = myDataRows("BankID")
                            strAccountNumber = myDataRows("AccountNumber").ToString()
                            strAccountType = myDataRows("AccountType").ToString()


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

    Public Shadows Function Delete(ByVal strDelQuery As String) As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin



            If lEmployerID = 0 Or _
                    Trim(strAccountNumber) = "" Or _
                        BankID = 0 _
                                Then

                MsgBox("Cannot Delete. Please select an existing " & _
                "Employer's Bank Account Details", _
                      MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete Information")

                datDelete = Nothing
                objLogin = Nothing
                Exit Function

            End If


            'Confirm Deletion
            If MsgBox("Do you want to delete this Employer's Bank Account?", _
            MsgBoxStyle.YesNo, "iManagement - Delete Records?") = _
                MsgBoxResult.No Then

                datDelete = Nothing
                objLogin = Nothing
                Exit Function
            End If

            strDeleteQuery = "DELETE * FROM EmployerBankAccount WHERE " & _
            " BankID = " & BankID & " AND AccountNr = '" & strAccountNumber & _
            "' AND EmployerId = " & lEmployerID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strDeleteQuery, _
                    datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Employer's Bank Account Record Deleted Successfully", _
                    MsgBoxStyle.Information, _
                        "iManagement - Record Deleted")
                Return True
            Else
                MsgBox("'Delete Employer's Bank Account' action failed", _
                    MsgBoxStyle.Exclamation, " Deletion failed")
            End If

            datDelete = Nothing
            objLogin = Nothing


        Catch ex As Exception

        End Try

    End Function

    Public Shadows Function Update(ByVal strUpQuery As String) As Boolean

        Try


            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If lEmployerID <> 0 Or _
                     lAccountID <> 0 Or _
                        BankID <> 0 _
                            Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                        "iManagement -  Customer's Bank Account Details Updated")
                    Return True
                End If

            End If


        Catch ex As Exception

        End Try

    End Function


#End Region

End Class
