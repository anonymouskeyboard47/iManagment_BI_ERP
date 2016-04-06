Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCustomerEmployment

#Region "PrivateVariables"
    Private lJobID As Long
    Private lCustomerNo As Long
    Private lEmployerID As Long
    Private lEmploymentTypeID As Long
    Private dtContractCommencmentDate As Date
    Private dtContractExpiryDate As Date
    Private strJobPosition As String
    Private strJobTitle As String
    Private strPhysicalAddress As String
    Private strNSSFNo As String
    Private lPaymentSchemeID As Long

#End Region

#Region "Properties"

    Public Property JobID() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lJobID
        End Get

        Set(ByVal Value As Long)
            lJobID = Value
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

    Public Property EmployerID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lEmployerID
        End Get

        Set(ByVal Value As Long)
            lEmployerID = Value
        End Set

    End Property

    Public Property EmploymentTypeID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lEmploymentTypeID
        End Get

        Set(ByVal Value As Long)
            lEmploymentTypeID = Value
        End Set

    End Property

    Public Property ContractCommencmentDate() As Date

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return dtContractCommencmentDate
        End Get

        Set(ByVal Value As Date)
            dtContractCommencmentDate = Value
        End Set

    End Property

    Public Property ContractExpiryDate() As Date

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return dtContractExpiryDate
        End Get

        Set(ByVal Value As Date)
            dtContractExpiryDate = Value
        End Set

    End Property

    Public Property JobPosition() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strJobPosition
        End Get

        Set(ByVal Value As String)
            strJobPosition = Value
        End Set

    End Property

    Public Property JobTitle() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strJobTitle
        End Get

        Set(ByVal Value As String)
            strJobTitle = Value
        End Set

    End Property

    Public Property PhysicalAddress() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhysicalAddress
        End Get

        Set(ByVal Value As String)
            strPhysicalAddress = Value
        End Set

    End Property


    Public Property NSSFNo() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strNSSFNo
        End Get

        Set(ByVal Value As String)
            strNSSFNo = Value
        End Set

    End Property

    Public Property PaymentSchemeID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lPaymentSchemeID
        End Get

        Set(ByVal Value As Long)
            lPaymentSchemeID = Value
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
        ljobid = 0
        lCustomerNo = ""
        lEmployerID = 0
        lEmploymentTypeID = 0
        dtContractCommencmentDate = Now
        dtContractExpiryDate = Now
        strJobPosition = ""
        strJobTitle = ""
        strPhysicalAddress = ""
        strNSSFNo = ""
        lPaymentSchemeID = 0


    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Sub Save()
        'Saves a new country name
        Try


            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String

            If lCustomerNo = 0 And _
                lEmployerID = 0 _
                                Then


            Else

                If Find("SELECT * FROM CustomerMaster WHERE CustomerNo = " & _
                    lCustomerNo, False) = False Then

                    MsgBox("The Customer Number provided does not exist." & _
                    " Cannot add this record.", _
                        MsgBoxStyle.Critical, _
                            "iManagement - invalid or incomplete information")

                    datSaved = Nothing
                    objLogin = Nothing

                    Exit Sub

                End If

                If Find("SELECT * FROM Employers WHERE EmployerID = " & _
                    lEmployerID, False) = False Then

                    MsgBox("The Employer provided does not exist." & _
                    " Cannot add this record", _
                        MsgBoxStyle.Critical, _
                            "iManagement - invalid or incomplete information")

                    datSaved = Nothing
                    objLogin = Nothing

                    Exit Sub

                End If

                'Check if there is an existing series with this name
                If Find("SELECT * FROM CustomerEmployment WHERE JobID = " & _
                    lJobID, False) = True Then

                    If MsgBox("The Employment Details already exists." & _
                    Chr(10) & "Do you want to update the details?", _
                            MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                    MsgBoxResult.Yes Then

                        Update("UPDATE MeetingMaster SET " & _
                            "EmployerID = " & lEmployerID & _
                                    " AND EmploymentTypeID = " & lEmploymentTypeID & _
                                    " AND ContractCommencmentDate = '" & dtContractCommencmentDate & _
                                    "' AND ContractExpiryDate = '" & dtContractExpiryDate & _
                                    "' AND JobPosition = '" & Trim(strJobPosition) & _
                                        "' AND JobTitle = '" & Trim(strJobTitle) & _
                                        "' AND PhysicalAddress = '" & Trim(strPhysicalAddress) & _
                                        "' AND NSSFNo = '" & Trim(strNSSFNo) & _
                                        "' AND PaymentSchemeID = " & lPaymentSchemeID & _
                                        " WHERE  JobID = " _
                                            & lJobID)

                    End If

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Sub
                End If


                strInsertInto = "INSERT INTO CustomerEmployment (" & _
                    "CustomerNo," & _
                    "EmployerID," & _
                    "EmploymentTypeID," & _
                    "ContractCommencmentDate," & _
                    "ContractExpiryDate," & _
                    "JobPosition," & _
                    "JobTitle," & _
                    "PhysicalAddress," & _
                    "NSSFNo," & _
                    "PaymentSchemeID" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "(" & lCustomerNo & _
                        "," & lEmployerID & _
                        "," & lEmploymentTypeID & _
                        ",'" & dtContractCommencmentDate & _
                        "','" & dtContractExpiryDate & _
                        "','" & Trim(strJobPosition) & _
                        "','" & Trim(strJobTitle) & _
                        "','" & Trim(strPhysicalAddress) & _
                        "','" & Trim(strNSSFNo) & _
                        "'," & lPaymentSchemeID & _
                                ")"

                objLogin.connectString = strAccessConnString
                objLogin.ConnectToDatabase()

                bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                strSaveQuery, _
                datSaved)

                objLogin.CloseDb()

                If bSaveSuccess = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - Customer's Employment Details Saved")

                Else

                    MsgBox("'Save Customer' action failed." & _
                        " Make sure all mandatory details are entered", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Customer's Employment Details Addition Failed")

                End If

            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Function Find(ByVal strQuery As String, ByVal bReturnValues As Boolean) As Boolean
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
                        datRetData = Nothing
                        objLogin = Nothing
                        Exit Function

                    End If

                    For Each myDataRows In myDataTables.Rows

                        If bReturnValues = True Then

                            lCustomerNo = myDataRows("CustomerNo")
                            lEmployerID = myDataRows("EmployerID")
                            lEmploymentTypeID = myDataRows("EmploymentTypeID")
                            dtContractCommencmentDate = myDataRows("ContractCommencmentDate")
                            dtContractExpiryDate = myDataRows("ContractExpiryDate")
                            strJobPosition = myDataRows("JobPosition").ToString()
                            strJobTitle = myDataRows("JobTitle").ToString()
                            strPhysicalAddress = myDataRows("PhysicalAddress").ToString()
                            strNSSFNo = myDataRows("NSSFNo").ToString()
                            lPaymentSchemeID = myDataRows("PaymentSchemeID")

                        End If

                    Next

                Next

                Return True
            Else
                Return False
            End If

            datRetData = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Sub Delete()

        Try
            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            If lJobID = 0 Then
                MsgBox("You have not selected an existing Job ID to delete.", _
                    MsgBoxStyle.Critical, "iManagement - invalid or incomplete information")

                objLogin = Nothing
                datDelete = Nothing

                Exit Sub

            Else

                If Find("SELECT * FROM CustomerEmployment WHERE JobID = " & lJobID, False) = False Then

                    MsgBox("You have not selected an existing Job ID to delete.", _
                                      MsgBoxStyle.Critical, _
                                        "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datDelete = Nothing

                    Exit Sub
                End If

                strDeleteQuery = "DELETE * FROM CustomerEmployement WHERE JobId = " & lJobID

                objLogin.connectString = strAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - Customer's Employment Details Deleted")
                Else
                    MsgBox("'Delete Customer's Employment' action failed", _
                        MsgBoxStyle.Exclamation, " Customer Employment Deletion failed")
                End If


            End If

            datDelete = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Update(ByVal strUpQuery As String)
        Try
            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If lCustomerNo = 0 And _
                    lEmployerID = 0 _
                            Then

                MsgBox("Please select an existing customer, employer, and a Job", _
                    MsgBoxStyle.Information, _
                        "iManagement - invalid or missing information")
                Exit Sub

            Else

                objLogin.connectString = strAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                        "iManagement -  Customer's Employment Details Updated")
                End If

            End If

            datUpdated = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Sub


#End Region


End Class
