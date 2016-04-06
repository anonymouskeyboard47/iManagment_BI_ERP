Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMEmployeeNextOFKin

    Inherits IMCustomer

#Region "PrivateVariables"

    Private lNOKID As Long
    Private strSurname As String
    Private strFirstName As String
    Private strMiddleName As String
    Private strOtherName As String
    Private strSex As String
    Private dtDateOfBirth As Date
    Private strCountryOfBirth As String
    Private strCountryofCitizenship As String
    Private strCountryOfResidence As String
    Private strPhysicalAddress As String
    Private strPostalAddress As String
    Private strPostCode As String
    Private strPostCountryCode As String
    Private strPostCityCode As String
    Private strPostTown As String
    Private strPhone1 As String
    Private strPhone2 As String
    Private strPhone3 As String
    Private strEmailAddress As String
    Private strPINNo As String

#End Region

#Region "Properties"


    Public Shadows Property Surname() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strSurname
        End Get

        Set(ByVal Value As String)
            strSurname = Value
        End Set

    End Property

    Public Shadows Property MiddleName() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strMiddleName
        End Get

        Set(ByVal Value As String)
            strMiddleName = Value
        End Set

    End Property

    Public Shadows Property FirstName() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strFirstName
        End Get

        Set(ByVal Value As String)
            strFirstName = Value
        End Set

    End Property

    Public Shadows Property OtherName() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strOtherName
        End Get

        Set(ByVal Value As String)
            strOtherName = Value
        End Set

    End Property

    Public Shadows Property Sex() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strSex
        End Get

        Set(ByVal Value As String)
            strSex = Value
        End Set

    End Property

    Public Shadows Property DateOfBirth() As Date

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return dtDateOfBirth
        End Get

        Set(ByVal Value As Date)
            dtDateOfBirth = Value
        End Set

    End Property

    Public Shadows Property CountryOfBirth() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCountryOfBirth
        End Get

        Set(ByVal Value As String)
            strCountryOfBirth = Value
        End Set

    End Property

    Public Shadows Property CountryOfCitizenship() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCountryofCitizenship
        End Get

        Set(ByVal Value As String)
            strCountryofCitizenship = Value
        End Set

    End Property

    Public Shadows Property CountryOfResidence() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCountryOfResidence
        End Get

        Set(ByVal Value As String)
            strCountryOfResidence = Value
        End Set

    End Property

    Public Shadows Property PhysicalAddress() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhysicalAddress
        End Get

        Set(ByVal Value As String)
            strPhysicalAddress = Value
        End Set

    End Property

    Public Shadows Property PostalAddress() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostalAddress
        End Get

        Set(ByVal Value As String)
            strPostalAddress = Value
        End Set

    End Property

    Public Shadows Property PostalCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostCode
        End Get

        Set(ByVal Value As String)
            strPostCode = Value
        End Set

    End Property

    Public Shadows Property PostCountryCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostCountryCode
        End Get

        Set(ByVal Value As String)
            strPostCountryCode = Value
        End Set

    End Property

    Public Shadows Property PostCityCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostCityCode
        End Get

        Set(ByVal Value As String)
            strPostCityCode = Value
        End Set

    End Property

    Public Shadows Property PostTown() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostTown
        End Get

        Set(ByVal Value As String)
            strPostTown = Value
        End Set

    End Property

    Public Shadows Property Phone1() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhone1
        End Get

        Set(ByVal Value As String)
            strPhone1 = Value
        End Set

    End Property

    Public Shadows Property Phone2() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhone2
        End Get

        Set(ByVal Value As String)
            strPhone2 = Value
        End Set

    End Property

    Public Shadows Property Phone3() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhone3
        End Get

        Set(ByVal Value As String)
            strPhone3 = Value
        End Set

    End Property

    Public Shadows Property EmailAddress() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strEmailAddress
        End Get

        Set(ByVal Value As String)
            strEmailAddress = Value
        End Set

    End Property

    Public Shadows Property PINNo() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPINNo
        End Get

        Set(ByVal Value As String)
            strPINNo = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Shadows Sub NewRecord()

        lNOKID = 0
        strSurname = ""
        strFirstName = ""
        strMiddleName = ""
        strSex = ""
        dtDateOfBirth = Now
        strCountryOfBirth = ""
        strCountryofCitizenship = ""
        strCountryOfResidence = ""
        strPhysicalAddress = ""
        strPostalAddress = ""
        strPostCode = ""
        strPostCountryCode = ""
        strPostCityCode = ""
        strPostTown = ""
        strPhone1 = ""
        strPhone2 = ""
        strPhone3 = ""
        strEmailAddress = ""
        strPINNo = ""

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

        If Trim(strFirstName) <> "" And _
            Trim(strSurname) <> "" And _
                dtDateOfBirth < Now() _
                            Then

            strInsertInto = "INSERT INTO NextOfKin (" & _
                "CustomerNo," & _
                "Surname " & _
                "FirstName," & _
                "MiddleName," & _
                "OtherName," & _
                "Sex," & _
                "DateOfBirth," & _
                "CountryOfBirth," & _
                "CountryOfCitizenship," & _
                "CountryOfResidence," & _
                "PhysicalAddress," & _
                "PostalAddress," & _
                "PostCode," & _
                "PostCountryCode," & _
                "PostCityCode," & _
                "PostTown," & _
                "Phone," & _
                "Phone2," & _
                "Phone3," & _
                "EmailAddress," & _
                "PINNo" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & CustomerNo & _
                    "', '" & strSurname & _
                    "', '" & strFirstName & _
                    "', '" & strMiddleName & _
                    "', '" & strOtherName & _
                    "', '" & strSex & _
                    "', '" & dtDateOfBirth & _
                    "', '" & strCountryOfBirth & _
                    "', '" & strCountryofCitizenship & _
                    "', '" & strCountryOfResidence & _
                    "', '" & strPhysicalAddress & _
                    "', '" & strPostalAddress & _
                    "', '" & strPostCode & _
                    "', '" & strPostCountryCode & _
                    "', '" & strPostCityCode & _
                    "', '" & strPostTown & _
                    "', '" & strPhone1 & _
                    "', '" & strPhone2 & _
                    "', '" & strPhone3 & _
                    "', '" & strEmailAddress & _
                    "', '" & strPINNo & _
                    "')"

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Customer Saved")

            Else

                MsgBox("'Save Customer' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Customer Addition Failed")

            End If

        End If

    End Sub

    Public Shadows Function Find(ByVal strQuery As String) As Boolean

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

                    CustomerNo = myDataRows("CustomerNo").ToString()
                    strSurname = myDataRows("Surname").ToString()
                    strFirstName = myDataRows("FirstName").ToString()
                    strMiddleName = myDataRows("MiddleName").ToString()
                    strOtherName = myDataRows("OtherName").ToString()
                    strSex = myDataRows("Sex").ToString()
                    dtDateOfBirth = myDataRows("DateOfBirth")
                    strCountryOfBirth = myDataRows("CountryOfBirth").ToString()
                    strCountryofCitizenship = myDataRows("CountryofCitizenship").ToString()
                    strCountryOfResidence = myDataRows("CountryOfResidence").ToString()
                    strPhysicalAddress = myDataRows("PhysicalAddress").ToString()
                    strPostalAddress = myDataRows("PostalAddress").ToString()
                    strPostCode = myDataRows("PostCode").ToString()
                    strPostCountryCode = myDataRows("PostCountryCode").ToString()
                    strPostCityCode = myDataRows("PostCityCode").ToString()
                    strPostTown = myDataRows("PostTown").ToString()
                    strPhone1 = myDataRows("Phone1").ToString()
                    strPhone2 = myDataRows("Phone2").ToString()
                    strPhone3 = myDataRows("Phone3").ToString()
                    strEmailAddress = myDataRows("EmailAddress").ToString()
                    strPINNo = myDataRows("PINNo").ToString()

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Shadows Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strDeleteQuery = strDelQuery

        If Trim(CustomerNo) <> "" Or _
                Trim(strFirstName) <> "" Or _
                    Trim(strSurname) <> "" _
                            Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Customer Details Deleted")
            Else
                MsgBox("'Delete Customer' action failed", _
                    MsgBoxStyle.Exclamation, " Customer Deletion failed")
            End If
        Else
            MsgBox("Cannot Delete. Please select an existing Activity", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

        End If

    End Sub

    Public Shadows Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If Trim(CustomerNo) <> "" Or _
                 Trim(strFirstName) <> "" Or _
                    Trim(strSurname) <> "" _
                        Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Customer Details Updated")
            End If

        End If

    End Sub


#End Region

End Class
