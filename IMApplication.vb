Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMApplication
    Inherits IMApplicationCustomer


#Region "PrivateVariables"

    Private lApplicationID As Long
    Private strApplicationSerial As String
    Private lApplicationTypeID As Long
    Private bApplicationStatus As Boolean 'If enabled or disabled Enabled or Disabled
    Private strApplicationStatusDescription As String
    Private dtApplicationDate As Date
    Private dtApplicationExpiryDate As Date
    Private strApplicationPosition As String 'Text description of Where it has reached

#End Region

#Region "Properties"

    Public Property ApplicationSerial() As String

        Get
            Return strApplicationSerial
        End Get

        Set(ByVal Value As String)
            strApplicationSerial = Value
        End Set

    End Property

    Public Property ApplicationID() As Long

        Get
            Return lApplicationID
        End Get

        Set(ByVal Value As Long)
            lApplicationID = Value
        End Set

    End Property

    Public Property ApplicationTypeID() As Long

        Get
            Return lApplicationTypeID
        End Get

        Set(ByVal Value As Long)
            lApplicationTypeID = Value
        End Set

    End Property

    Public Property ApplicationStatus() As Boolean

        Get
            Return bApplicationStatus
        End Get

        Set(ByVal Value As Boolean)
            bApplicationStatus = Value
        End Set

    End Property

    Public Property ApplicationStatusDescription() As String

        Get
            Return strApplicationStatusDescription
        End Get

        Set(ByVal Value As String)
            strApplicationStatusDescription = Value
        End Set

    End Property

    Public Property ApplicationDate() As Date

        Get
            Return dtApplicationDate
        End Get

        Set(ByVal Value As Date)
            dtApplicationDate = Value
        End Set

    End Property

    Public Property ApplicationExpiryDate() As Date

        Get
            Return dtApplicationExpiryDate
        End Get

        Set(ByVal Value As Date)
            dtApplicationExpiryDate = Value
        End Set

    End Property

    Public Property ApplicationPosition() As String

        Get
            Return strApplicationPosition
        End Get

        Set(ByVal Value As String)
            strApplicationPosition = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function CalculateNextApplNo() As String
        Try

            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String
            Dim strProposedApplNo As String

            Dim objLogin As IMLogin = New IMLogin

                MyMaxValue = objLogin.FillArray(strOrgAccessConnString, _
                                "SELECT COUNT(*) AS TotalRecords FROM" & _
                                    " ApplicationMaster " & _
                                    " WHERE " & _
                    " Day(ApplicationDate) = Day(Now()) " & _
                    " AND " & _
                    " Year(ApplicationDate) = Year(Now()) " & _
                    " AND " & _
                    " Month(ApplicationDate)=Month(Now())", "", "")


                If Not MyMaxValue Is Nothing Then
                    For Each strItem In MyMaxValue
                        If Not strItem Is Nothing Then

                            MaxValue = CLng(Val(strItem))


                        End If
                    Next
                End If

                MaxValue = MaxValue + 1

            strProposedApplNo = "Appl" & Now.Day.ToString _
                & Now.Month.ToString & _
                    Now.Year.ToString & _
                            MaxValue.ToString

            objLogin = Nothing

                Return strProposedApplNo

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, _
                "iManagement - System Error")
        End Try

    End Function

#End Region

#Region "DatabaseProcedures"

    'Save informaiton
    Public Function Save(ByVal bDisplayErrorMessages As Boolean, _
            ByVal bDisplaySuccessMessages As Boolean, _
                ByVal bDisplayFailureMessages As Boolean, _
                    ByVal bSaveApplicationCustomer As Boolean) As Boolean

        'Saves a new base organization
        Try

            Dim trTransDB As OleDbTransaction



            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String
            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String

            If Trim(strOrganizationName) = "" Then

                MsgBox("Please open an existing company.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Select an existing company")
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Make sure the customer number has been entered
            If ApplCustCustomerNo = 0 Then
                MsgBox("Please provide an existing customer number", _
                    MsgBoxStyle.Critical, _
                        "iManagement - invalid or incomplete information")
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Make sure the Application Type has been entered
            If lApplicationTypeID = 0 Then
                MsgBox("You must provide an appropriate Application Type." _
                                , MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if there is an existing series with this name
            If Find("SELECT * FROM ApplicationMaster " & _
                "WHERE ApplicationSerial = '" & _
                    Trim(strApplicationSerial) & "'", _
                        False, True, True) = True Then

                If MsgBox("The Application Number already exists." & _
                    Chr(10) & "Do you want to update the details?", _
                            MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                    MsgBoxResult.Yes Then

                    Update("UPDATE ApplicationMaster SET " & _
                        " ApplicationTypeID = " & lApplicationTypeID & _
                        " , ApplicationStatus = " & bApplicationStatus & _
                        " , ApplicationStatusDescription = '" & _
                        strApplicationStatusDescription & _
                        "' , ApplicationDate = '" & dtApplicationDate & _
                        "' , ApplicationExpiryDate = '" & _
                        dtApplicationExpiryDate & _
                        "' , ApplicationPosition = '" & _
                        strApplicationPosition & _
                        "' WHERE  strProductName = '" & _
                        strApplicationSerial & "'")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If

            strApplicationSerial = CalculateNextApplNo()

            strInsertInto = "INSERT INTO ApplicationMaster (" & _
                "ApplicationTypeID," & _
                "ApplicationStatus," & _
                "ApplicationSerial," & _
                "ApplicationStatusDescription," & _
                "ApplicationDate," & _
                "ApplicationExpiryDate," & _
                "ApplicationPosition" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lApplicationTypeID & _
                    "," & bApplicationStatus & _
                    ",'" & Trim(strApplicationSerial) & _
                    "','" & Trim(strApplicationStatusDescription) & _
                    "','" & dtApplicationDate & _
                    "','" & dtApplicationExpiryDate & _
                    "','" & Trim(strApplicationPosition) & _
                    "')"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()
            objLogin.BeginTheTrans()


            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            If bSaveApplicationCustomer = True Then
                Find("SELECT * FROM AppicationMaster WHERE " & _
                " ApplicationSerial = '" & strApplicationSerial & "'", _
                True, True, False)


                'Save Customer
                bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                "INSERT INTO ApplicationCustomer (CustomerNo,ApplicationID ) " & _
                " VALUES (" & ApplCustCustomerNo & "," & lApplicationID & ")", _
                datSaved)

               
                objLogin.CloseDb()
            End If



            If bSaveSuccess = True Then
                If bDisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", _
                    MsgBoxStyle.Information, _
                    "iManagement - New Application Saved")

                End If

            Else

                If bDisplayFailureMessages = True Then
                    MsgBox("'Save New Application' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save New Application Failed")
                End If

            End If

            objLogin = Nothing
            datSaved = Nothing


        Catch ex As Exception

            If bDisplayErrorMessages = True Then
                MsgBox(ex.Source, MsgBoxStyle.Critical, _
                    "iManagement - Database or system error")

            End If
        End Try

    End Function

    'Find Information
    Public Function Find(ByVal strQuery As String, _
        ByVal bReturnValues As Boolean, _
            ByVal bReturnApplicationDetails As Boolean, _
                ByVal bReturnApplicationCustomer As Boolean) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.ConnectString = strOrgAccessConnString
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
                    objLogin = Nothing
                    datRetData = Nothing


                    Exit Function

                End If


                If bReturnValues = True Then
                    For Each myDataRows In myDataTables.Rows

                        If bReturnApplicationDetails = True Then
                            lApplicationID = _
                                    myDataRows("ApplicationID")
                            strApplicationSerial = _
                                    myDataRows("ApplicationSerial").ToString
                            lApplicationTypeID = _
                                    myDataRows("ApplicationTypeID")
                            bApplicationStatus = _
                                    myDataRows("ApplicationStatus")
                            strApplicationStatusDescription = _
                                    myDataRows("ApplicationStatusDescription").ToString
                            dtApplicationDate = _
                                    myDataRows("ApplicationDate")
                            dtApplicationExpiryDate = _
                                    myDataRows("ApplicationExpiryDate")
                            strApplicationPosition = _
                                    myDataRows("ApplicationPosition").ToString

                        End If

                        'Do you return values from application customer
                        If bReturnApplicationCustomer Then
                            ApplCustFind("SELECT * FROM ApplicationCustomer " & _
                            " WHERE ApplicationID = " & lApplicationID, True)

                        End If

                    Next
                End If
            Next

            objLogin = Nothing
            datRetData = Nothing

            Return True
        Else

            objLogin = Nothing
            datRetData = Nothing
            Return False

        End If

        objLogin = Nothing
        datRetData = Nothing

    End Function

    'Delete data
    Public Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try


            strDeleteQuery = strDelQuery

            If strApplicationSerial = "" _
                 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - Application Deleted")
                Else
                    MsgBox("'Application delete' action failed", _
                        MsgBoxStyle.Exclamation, "Application Deletion failed")
                End If
            Else

                MsgBox("Cannot Delete. Please select an existing Application.", _
                        MsgBoxStyle.Exclamation, "iManagement - invalid or incomplete information")

            End If

            objLogin = Nothing
            datDelete = Nothing

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

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Application Details Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub


#End Region


End Class
