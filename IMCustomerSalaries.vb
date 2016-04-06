Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCustomerSalaries

#Region "PrivateVariables"

    Private lSalaryID As Long
    Private lCustomerNo As Long
    Private dtStartDate As Date
    Private dbSalaryAmount As Decimal
    Private lEmployerID As Long
    Private lSalaryTypeID As Long
    Private bSalaryStatus As Boolean

#End Region

#Region "Properties"


    Public Property SalaryID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lSalaryID
        End Get

        Set(ByVal Value As Long)
            lSalaryID = Value
        End Set

    End Property


    Public Property SalaryTypeID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lSalaryTypeID
        End Get

        Set(ByVal Value As Long)
            lSalaryTypeID = Value
        End Set

    End Property

    Public Property CustomerNo() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lCustomerNo
        End Get

        Set(ByVal Value As Long)
            lCustomerNo = Value
        End Set

    End Property

    Public Property StartDate() As Date

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return dtStartDate
        End Get

        Set(ByVal Value As Date)
            dtStartDate = Value
        End Set

    End Property

    Public Property SalaryAmount() As Decimal

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return dbSalaryAmount
        End Get

        Set(ByVal Value As Decimal)
            dbSalaryAmount = Value
        End Set

    End Property

    Public Property EmployerID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lEmployerID
        End Get

        Set(ByVal Value As Long)
            lEmployerID = Value
        End Set

    End Property

    Public Property SalaryStatus() As Boolean

        Get
            Return bSalaryStatus
        End Get

        Set(ByVal Value As Boolean)
            bSalaryStatus = Value
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

        lCustomerNo = ""
        dtStartDate = Now
        dbSalaryAmount = 0
        lEmployerID = 0

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

        'If bSalaryStatus = True Then
        '    If Find("SELECT * FROM CustomerSalaries WHERE CustomerNo =" & _
        '        lCustomerNo & " AND EmployerID = " & lEmployerID & _
        '        " AND SalaryStatus = TRUE") = True Then

        '        If MsgBox("This customer already has another salary that is enabled." & _
        '            " Do you want to set this new salary" & _
        '                Chr(10) & " as the current salary for the customer with the employer" & _
        '                    " and disable the previous one?", _
        '                        MsgBoxStyle.YesNo, _
        '                            "IManagement - Record Exists") = MsgBoxResult.No Then

        '            Exit Sub
        '        Else

        '            Update("UPDATE CustomerSalaries SET " & _
        '                " SalaryStatus = FALSE" & _
        '                    " WHERE CustomerNo = " & _
        '                        lCustomerNo & " EmployerID = " & _
        '                            lEmployerID & " AND SalaryStatus = TRUE")


        '            If Find("SELECT * FROM CustomerSalaries WHERE CustomerNo =" & _
        '                    lCustomerNo & " AND EmployerID = " & lEmployerID & _
        '                        " AND SalaryStatus = TRUE OR (StartDate => '" & _
        '                            dtStartDate & "' AND EndDate <= '" & _
        '                                dtEndDate & "') OR (StartDate => '" & _
        '                                    dtEndDate & "' AND EndDate <= '" & _
        '                                        dtEndDate & "')") = True Then

        '                If MsgBox("The Dates specified fall in between another date" & _
        '                    " for the same customer and " & _
        '                        Chr(10) & "same employee. Do you want to update the details?", _
        '                            MsgBoxStyle.Critical, _
        '                                "iManagement - Update the details?") = MsgBoxResult.No Then
        '                    Exit Sub

        '                Else

        '                    Update("UPDATE CustomerSalaries SET " & _
        '                        " EndDate = '" & dtEndDate & "'" & _
        '                            " StartDate = '" & dtStartDate & "'" & _
        '                                " WHERE CustomerNo = " & _
        '                                    lCustomerNo & " EmployerID = " & _
        '                                        lEmployerID & " SalaryStatus = TRUE")

        '                End If
        '            End If
        '        End If
        '    End If
        'End If


        If lCustomerNo = 0 Or _
            dbSalaryAmount = 0 Or _
                    lEmployerID = 0 Or _
                        lSalaryTypeID = 0 _
                                    Then

            MsgBox("To save the Customer's Salary you must provide:" & _
                Chr(10) & "1 : The Customer Number" & _
                    Chr(10) & "2: The Salary Amount" & _
                        Chr(10) & "3: The Customer's Employer" & _
                            Chr(10) & "4: The Customer's Salary Type.", _
                                MsgBoxStyle.Exclamation, _
                                    "iManagement - invalid or incomplete information")

            datSaved = Nothing
            objLogin = Nothing

            Exit Sub

        End If


        'Check if there is an existing series with this name
        If Find("SELECT * FROM CustomerSalaries WHERE  SalaryID = " _
                    & lSalaryID, False) = True Then

            If MsgBox("The NOK Details already exists." & _
            Chr(10) & "Do you want to update the details?", _
                    MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                            MsgBoxResult.Yes Then

                Update("UPDATE CustomerSalaries SET " & _
                            "CustomerNo = " & lCustomerNo & _
                            " AND StartDate = '" & dtStartDate & _
                            "' AND SalaryAmount = " & dbSalaryAmount & _
                            " AND EmployerID = " & lEmployerID & _
                            " AND SalaryStatus = " & bSalaryStatus & _
                                " WHERE  SalaryTypeID = " _
                                    & lSalaryTypeID)

            End If

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If


        strInsertInto = "INSERT INTO CustomerSalaries (" & _
            "CustomerNo," & _
            "StartDate," & _
            "SalaryAmount," & _
            "EmployerID," & _
            "SalaryTypeID," & _
            "SalaryStatus" & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                "(" & lCustomerNo & _
                " ,' " & dtStartDate & _
                " ', " & dbSalaryAmount & _
                " , " & lEmployerID & _
                " , " & lSalaryTypeID & _
                " , " & bSalaryStatus & _
                        ")"

        objLogin.connectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
        strSaveQuery, _
        datSaved)

        objLogin.CloseDb()

        If bSaveSuccess = True Then
            MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
            "iManagement - Customer's Salary Details Saved")

        Else

            MsgBox("'Save Customer Salary action failed." & _
                " Make sure all mandatory details are entered", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Customer's Salary Details Addition Failed")

        End If


    End Sub

    Public Function Find(ByVal strQuery As String, ByVal bReturnValues As Boolean) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.connectString = strAccessConnString
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
                        lCustomerNo = myDataRows("CustomerNo")
                        dtStartDate = myDataRows("StartDate")
                        dbSalaryAmount = myDataRows("SalaryAmount")
                        lEmployerID = myDataRows("EmployerID")
                        lSalaryTypeID = myDataRows("SalaryTypeID")

                    End If
                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Sub Delete()

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin


        If lSalaryID = 0 Then
            MsgBox("Please select an existing Salary Type identifier", _
            MsgBoxStyle.Exclamation, "iManagement - invalid or incomplete information")

            datDelete = Nothing
            objLogin = Nothing

            Exit Sub

        End If

        strDeleteQuery = "DELETE * FROM CustomerSalaries WHERE SalaryID = " & lSalaryID

        objLogin.connectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
        datDelete)

        objLogin.CloseDb()

        If bDelSuccess = True Then
            MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                "iManagement - Customer's Salary Details Deleted")
        Else
            MsgBox("'Delete Customer Salary' action failed", _
                MsgBoxStyle.Exclamation, " Customer Salary Deletion failed")
        End If

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If lCustomerNo <> 0 _
                        Then

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Customer's Salary Details Updated")
            End If

        End If

    End Sub

    


#End Region


End Class
