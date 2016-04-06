Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMEmployeeMonetaryActivities

#Region "PrivateVariables"

    Private lCustomerNo As Long
    Private lMonetaryActivityID As Long
    Private lCompanyActivityID As Long
    Private strActivityDescription As String
    Private strPhysicalAddress As String
    Private strPostalAddress As String
    Private bActivityStatus As Boolean
    Private dbMonthlyReturns As Decimal

#End Region

#Region "Properties"

    Public Property CustomerNo() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lCustomerNo
        End Get

        Set(ByVal Value As Long)
            lCustomerNo = Value
        End Set

    End Property

    Public Property ActivityID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lMonetaryActivityID
        End Get

        Set(ByVal Value As Long)
            lMonetaryActivityID = Value
        End Set

    End Property

    Public Property CompanyActivityID() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return CompanyActivityID
        End Get

        Set(ByVal Value As Long)
            CompanyActivityID = Value
        End Set

    End Property

    Public Property ActivityDescription() As String

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return strActivityDescription
        End Get

        Set(ByVal Value As String)
            strActivityDescription = Value
        End Set

    End Property

    Public Property PhysicalAddress() As String

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return strPhysicalAddress
        End Get

        Set(ByVal Value As String)
            strPhysicalAddress = Value
        End Set

    End Property

    Public Property PostalAddress() As String

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return strPostalAddress
        End Get

        Set(ByVal Value As String)
            strPostalAddress = Value
        End Set

    End Property

    Public Property ActivityStatus() As Boolean

        Get
            Return bActivityStatus
        End Get

        Set(ByVal Value As Boolean)
            bActivityStatus = Value
        End Set

    End Property

    Public Property MonthlyReturns() As Decimal

        Get
            Return dbMonthlyReturns
        End Get

        Set(ByVal Value As Decimal)
            dbMonthlyReturns = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
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

        If lCustomerNo <> 0 And _
            lCompanyActivityID <> 0 _
                            Then

            strInsertInto = "INSERT INTO CustomerSalaries (" & _
                "CustomerNo," & _
                "ActivityID," & _
                "ActivityTitle," & _
                "ActivityDescription," & _
                "PhysicalAddress," & _
                "PostalAddress," & _
                "ActivityStatus," & _
                "MonthlyReturns" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lCustomerNo & _
                    " , " & lMonetaryActivityID & _
                    " , " & lCompanyActivityID & _
                    " , '" & strActivityDescription & _
                    "', '" & strPhysicalAddress & _
                    "', '" & strPostalAddress & _
                    "', " & bActivityStatus & _
                    ", " & dbMonthlyReturns & _
                            ")"

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
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
                    Exit Function

                End If

                For Each myDataRows In myDataTables.Rows

                    lCustomerNo = myDataRows("CustomerNo")
                    lMonetaryActivityID = myDataRows("MonetaryActivityID")
                    lCompanyActivityID = myDataRows("CompanyActivityID")
                    strActivityDescription = myDataRows("ActivityDescription").ToString()
                    strPhysicalAddress = myDataRows("PhysicalAddress").ToString()
                    strPostalAddress = myDataRows("PostalAddress").ToString()
                    bActivityStatus = myDataRows("ActivityStatus")

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

        If lCustomerNo <> 0 And _
                lMonetaryActivityID <> 0 _
                            Then

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Customer's Other Earnings Details Deleted")
            Else
                MsgBox("'Delete Other Earnings' action failed", _
                    MsgBoxStyle.Exclamation, " Customer Other Earnings Deletion failed")
            End If
        Else
            MsgBox("Cannot Delete. Please select an existing Customer's Other Earnings Detail", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

        End If

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If lCustomerNo <> 0 And _
                 lMonetaryActivityID <> 0 _
                        Then

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Customer's Other Earnings Details Updated")
            End If

        End If

    End Sub

  
#End Region


End Class
