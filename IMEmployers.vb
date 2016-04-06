Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMEmployers


#Region "Connection Properties"

    Public Property AccessConnectionString() As String

        Get
            Return strAccessConnString
        End Get

        Set(ByVal Value As String)
            strAccessConnString = Value
        End Set

    End Property

    Public Property OrgConnectionString() As String

        Get
            Return strOrgAccessConnString
        End Get

        Set(ByVal Value As String)
            strOrgAccessConnString = Value
        End Set

    End Property

    Public Property OrgConnectionSQLServer() As String

        Get
            Return bOrgConnectionSQLServer
        End Get

        Set(ByVal Value As String)
            bOrgConnectionSQLServer = Value
        End Set

    End Property

    Public Property AccessConnStringADOX() As String

        Get
            Return strAccessConnStringADOX
        End Get

        Set(ByVal Value As String)
            strAccessConnStringADOX = Value
        End Set

    End Property

    Public Property OrgAccessConnStringADOX() As Boolean

        Get
            Return strOrgAccessConnStringADOX
        End Get

        Set(ByVal Value As Boolean)
            strOrgAccessConnStringADOX = Value
        End Set

    End Property

    Public Property SQLConnString() As String

        Get
            Return strSQLConnString
        End Get

        Set(ByVal Value As String)
            strSQLConnString = Value
        End Set

    End Property

    Public Property DBUserName() As String

        Get
            Return strDBUserName
        End Get

        Set(ByVal Value As String)
            strDBUserName = Value
        End Set

    End Property

    Public Property DBPassword() As String

        Get
            Return strDBPassword
        End Get

        Set(ByVal Value As String)
            strDBPassword = Value
        End Set

    End Property

    Public Property DBDatabase() As String

        Get
            Return strDBDatabase
        End Get

        Set(ByVal Value As String)
            strDBDatabase = Value
        End Set

    End Property

    Public Property DBDBPath() As String

        Get
            Return strDBDBPath
        End Get

        Set(ByVal Value As String)
            strDBDBPath = Value
        End Set

    End Property

#End Region


#Region "PrivateBankVariables"

    Private lEmployerID As Long
    Private strEmployerName As String
    Private strPhysicalAddress As String
    Private strPostalAddress As String
    Private strPostalCode As String
    Private strPostalCountryCode As String
    Private strPostalCityCode As String
    Private strPostalTown As String
    Private lCompanyTypeID As Long
    Private lActivityID As Long
    Private lNoOfEmployees As Long
    Private strFaxNumber As String
    Private strHumanResourceFax As String
    Private lMainEmploymentTypeID As Long
    Private strHREmailAddress As String

    Private strTaxInformationNumber As String
    Private strCompanyRegistrationNumber As String
    Private dtCompanyRegistrationDate As Date
    Private dtDateCreated As Date

    Private strCountryCode As String
    Private strCityCode As String
    Private strTown As String
    Private strPhone1 As String
    Private strPhone2 As String
    Private bVATRegistered As Boolean

#End Region

#Region "Properties"

    Public Property CountryCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCountryCode
        End Get

        Set(ByVal Value As String)
            strCountryCode = Value
        End Set

    End Property

    Public Property CityCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeDescription (STRING)
        Get
            Return strCityCode
        End Get

        Set(ByVal Value As String)
            strCityCode = Value
        End Set

    End Property

    Public Property Town() As String

        Get
            Return strTown
        End Get

        Set(ByVal Value As String)
            strTown = Value
        End Set

    End Property

    'TaxInformationNumber
    Public Property TaxInformationNumber() As String

        Get
            Return strTaxInformationNumber
        End Get

        Set(ByVal Value As String)
            strTaxInformationNumber = Value
        End Set

    End Property

    'CompRegNum
    Public Property CompanyRegistrationNumber() As String

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return strCompanyRegistrationNumber
        End Get

        Set(ByVal Value As String)
            strCompanyRegistrationNumber = Value
        End Set

    End Property

    'CompanyRegDate
    Public Property CompanyRegistrationDate() As Date

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return dtCompanyRegistrationDate
        End Get

        Set(ByVal Value As Date)
            dtCompanyRegistrationDate = Value
        End Set

    End Property

    Public Property DateCreated() As Date

        'USED TO SET AND RETRIEVE THE SalaryTypeDescription (STRING)
        Get
            Return dtDateCreated
        End Get

        Set(ByVal Value As Date)
            dtDateCreated = Value
        End Set

    End Property

    'VATRegistered
    Public Property VATRegistered() As Boolean

        Get
            Return bVATRegistered
        End Get

        Set(ByVal Value As Boolean)
            bVATRegistered = Value
        End Set

    End Property


    'Phone1
    Public Property Phone1() As String

        Get
            Return strPhone1
        End Get

        Set(ByVal Value As String)
            strPhone1 = Value
        End Set

    End Property

    'Phone2
    Public Property Phone2() As String

        Get
            Return strPhone2
        End Get

        Set(ByVal Value As String)
            strPhone2 = Value
        End Set

    End Property


    Public Property EmployerID() As Long

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return lEmployerID
        End Get

        Set(ByVal Value As Long)
            lEmployerID = Value
        End Set

    End Property

    Public Property EmployerName() As String

        'USED TO SET AND RETRIEVE THE BANK NAME (STRING)
        Get
            Return strEmployerName
        End Get

        Set(ByVal Value As String)
            strEmployerName = Value
        End Set

    End Property

    Public Property PhysicalAddress() As String

        'USED TO SET AND RETRIEVE THE PHYSICAL ADDRESS (STRING)
        Get
            Return strPhysicalAddress
        End Get

        Set(ByVal Value As String)
            strPhysicalAddress = Value
        End Set

    End Property

    Public Property PostalAddress() As String

        'USED TO SET AND RETRIEVE THE POSTAL ADDRESS (STRING)
        Get
            Return strPostalAddress
        End Get

        Set(ByVal Value As String)
            strPostalAddress = Value
        End Set

    End Property

    Public Property PostCode() As String

        'USED TO SET AND RETRIEVE THE POSTCODE (STRING)
        Get
            Return strPostalCode
        End Get

        Set(ByVal Value As String)
            strPostalCode = Value
        End Set

    End Property

    Public Property PostCountryCode() As String

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return strPostalCountryCode
        End Get

        Set(ByVal Value As String)
            strPostalCountryCode = Value
        End Set

    End Property

    Public Property PostalCityCode() As String

        'USED TO SET AND RETRIEVE THE POST CITY CODE (STRING)
        Get
            Return strPostalCityCode
        End Get

        Set(ByVal Value As String)
            strPostalCityCode = Value
        End Set

    End Property

    Public Property PostTown() As String

        'USED TO SET AND RETRIEVE THE POST TOWN CODE (STRING)
        Get
            Return strPostalTown
        End Get

        Set(ByVal Value As String)
            strPostalTown = Value
        End Set

    End Property

    Public Property CompanyTypeID() As Long

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return lCompanyTypeID
        End Get

        Set(ByVal Value As Long)
            lCompanyTypeID = Value
        End Set

    End Property

    Public Property ActivityID() As Long

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return lActivityID
        End Get

        Set(ByVal Value As Long)
            lActivityID = Value
        End Set

    End Property

    Public Property NoOfEmployees() As Long

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return lNoOfEmployees
        End Get

        Set(ByVal Value As Long)
            lNoOfEmployees = Value
        End Set

    End Property

    Public Property FaxNumber() As String

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return strFaxNumber
        End Get

        Set(ByVal Value As String)
            strFaxNumber = Value
        End Set

    End Property

    Public Property MainEmploymentTypeID() As Long

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return lMainEmploymentTypeID
        End Get

        Set(ByVal Value As Long)
            lMainEmploymentTypeID = Value
        End Set

    End Property

    Public Property HREmailAddress() As String

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return strHREmailAddress
        End Get

        Set(ByVal Value As String)
            strHREmailAddress = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function ReturnEmployerIdFromEmployerName( _
         ByVal strValEmployerString As String) As Long

        Try

            Dim objEmployer As IMEmployers = New IMEmployers

            With objEmployer
                .Find("SELECT * FROM Employers WHERE Name = '" _
                & strValEmployerString & "'", True)

                Return .EmployerID

            End With

            objEmployer = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnEmployerNameFromEmployerID( _
            ByVal lValEmployerID As Long) As String

        Try

            Dim objEmployer As IMEmployers = New IMEmployers

            With objEmployer
                .Find("SELECT * FROM Employers WHERE EmployerID = " _
                & lValEmployerID, True)

                Return .EmployerName

            End With

            objEmployer = Nothing

        Catch ex As Exception

        End Try


    End Function


    Public Sub NewRecord()

        lEmployerID = 0
        strEmployerName = ""
        strPhysicalAddress = ""
        strPostalAddress = ""
        strPostalCode = ""
        strPostalCountryCode = ""
        strPostalCityCode = ""
        strPostalTown = ""
        lCompanyTypeID = 0
        lActivityID = 0
        lNoOfEmployees = 0
        strFaxNumber = ""
        strHumanResourceFax = ""
        lMainEmploymentTypeID = 0
        strHREmailAddress = ""


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


        If Trim(strEmployerName) = "" Or _
            Trim(PhysicalAddress) = "" Or _
                Trim(lCompanyTypeID) = 0 Or _
                Trim(lActivityID) = 0 _
                             Then

            MsgBox("Cannot save Employer Information. Missing information." & _
            Chr(10) & "1: Employer Name" & _
            Chr(10) & "2: Company Type" & _
            Chr(10) & "3: Company Activity" _
                    , MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete Information")

            datSaved = Nothing
            objLogin = Nothing
            Exit Sub

        End If


        If Find("SELECT * FROM Employers WHERE Name = '" _
               & Trim(strEmployerName) & "'", _
                   False) = True Then

            If MsgBox("Do you want to update the details?" & _
            " Remember updating is amending details of an existing record," _
             & Chr(10) & "that is, this organization name already exists.", _
                MsgBoxStyle.YesNo, "iManagement - Update existing company details?") _
                    = MsgBoxResult.Yes Then

                Update("UPDATE Employers SET" & _
                    " CompRegNum = '" & Trim(strCompanyRegistrationNumber) & _
                    "' , TaxInformationNumber = '" & Trim(strTaxInformationNumber) & _
                    "' , CompanyRegDate = '" & dtCompanyRegistrationDate & _
                    "' , PhysicalAddress = '" & Trim(strPhysicalAddress) & _
                    "' , CountryCode = '" & Trim(strCountryCode) & _
                    "' , CityCode = '" & Trim(strCityCode) & _
                    "' , Town = '" & Trim(strTown) & _
                    "' , CompanyTypeID = " & lCompanyTypeID & _
                    " , ActivityID = " & lActivityID & _
                    " , PostalAddress = '" & Trim(strPostalAddress) & _
                    "' , PostalCode = '" & Trim(strPostalCode) & _
                    "' , PostalCountryCode = '" & Trim(strPostalCountryCode) & _
                    "' , PostalCityCode = '" & Trim(strPostalCityCode) & _
                    "' , VATRegistered = " & bVATRegistered & _
                    " , Phone1 = '" & Trim(strPhone1) & _
                    "' , Phone2 = '" & Trim(strPhone1) & _
                    "' , FaxNumber = '" & Trim(strFaxNumber) & _
                    "' WHERE Name = '" & Trim(strEmployerName) & "'")

            End If

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If


        strInsertInto = "INSERT INTO Employers (" & _
                "Name," & _
                "PhysicalAddress," & _
                "PostalAddress," & _
                "PostalCode," & _
                "PostalCountryCode," & _
                "PostalCityCode," & _
                "PostalTown," & _
                "CompanyTypeID," & _
                "ActivityID," & _
                "NoOfEmployees," & _
                "FaxNumber," & _
                "HumanResourceFax," & _
                "MainEmploymentTypeID," & _
                "HREmailAddress," & _
                "TaxInformationNumber," & _
                "CompRegNum," & _
                "CompanyRegDate," & _
                "DateCreated," & _
                "CountryCode," & _
                "CityCode," & _
                "Town," & _
                "VATRegistered," & _
                "Phone1," & _
                "Phone2" & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                    "(" & _
                "'" & Trim(strEmployerName) & _
                "', '" & Trim(strPhysicalAddress) & _
                "', '" & Trim(strPostalAddress) & _
                "', '" & Trim(strPostalCode) & _
                "', '" & Trim(strPostalCountryCode) & _
                "', '" & Trim(strPostalCityCode) & _
                "', '" & Trim(strPostalTown) & _
                "', " & lCompanyTypeID & _
                ", " & lActivityID & _
                ", " & lNoOfEmployees & _
                ", '" & Trim(strFaxNumber) & _
                "', '" & Trim(strHumanResourceFax) & _
                "', " & lMainEmploymentTypeID & _
                ", '" & Trim(strHREmailAddress) & _
                "' , '" & Trim(strTaxInformationNumber) & _
                "' , '" & Trim(strCompanyRegistrationNumber) & _
                "' , '" & dtCompanyRegistrationDate & _
                "' , '" & dtDateCreated & _
                "' , '" & Trim(strCountryCode) & _
                "' , '" & Trim(strCityCode) & _
                "' , '" & Trim(strTown) & _
                "' , " & bVATRegistered & _
                " , '" & Trim(strPhone1) & _
                "' , '" & Trim(strPhone2) & _
                    "')"

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                            strSaveQuery, _
                                    datSaved)

        objLogin.CloseDb()

        If bSaveSuccess = True Then
            MsgBox("Record Saved Successfully", _
                MsgBoxStyle.Information, _
                    "iManagement - Employer Details Saved")

        Else

            MsgBox("'Save Employer' action failed." & _
                " Make sure all mandatory details are entered", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Save Employer Details Failed")

        End If



    End Sub

    Public Function Find(ByVal strQuery As String, _
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
                        datRetData = Nothing
                        objLogin = Nothing
                        Return False
                        Exit Function

                    End If


                    For Each myDataRows In myDataTables.Rows

                        If bReturnDetails = True Then
                            lEmployerID = myDataRows("EmployerID").ToString()
                            strEmployerName = myDataRows("Name").ToString()
                            strPhysicalAddress = myDataRows("PhysicalAddress").ToString()
                            strPostalAddress = myDataRows("PostalAddress").ToString()
                            strPostalCode = myDataRows("PostalCode").ToString()
                            strPostalCountryCode = myDataRows("PostalCountryCode").ToString()
                            strPostalCityCode = myDataRows("PostalCityCode").ToString()
                            strPostalTown = myDataRows("PostalTown").ToString()
                            lCompanyTypeID = myDataRows("CompanyTypeID").ToString()
                            lActivityID = myDataRows("ActivityID").ToString()
                            lNoOfEmployees = myDataRows("NoOfEmployees").ToString()
                            strFaxNumber = myDataRows("FaxNumber").ToString()
                            strHumanResourceFax = myDataRows("HumanResourceFax").ToString()
                            lMainEmploymentTypeID = myDataRows("MainEmploymentTypeID").ToString()
                            strHREmailAddress = myDataRows("HREmailAddress").ToString()

                            strTaxInformationNumber = myDataRows("TaxInformationNumber").ToString()
                            strCompanyRegistrationNumber = myDataRows("CompRegNum").ToString()
                            dtCompanyRegistrationDate = myDataRows("CompanyRegDate")
                            dtDateCreated = myDataRows("DateCreated")
                            strCountryCode = myDataRows("CountryCode").ToString()
                            strCityCode = myDataRows("CityCode").ToString()
                            strTown = myDataRows("Town").ToString()
                            strPhone1 = myDataRows("Phone1").ToString()
                            strPhone2 = myDataRows("Phone2").ToString()
                            bVATRegistered = myDataRows("VATRegistered")

                        End If
                    Next
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


            'Deletes the country details of the country with the country code
            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin



            If Trim(strEmployerName) = "" And lEmployerID = 0 _
                         Then

                MsgBox("Cannot delete the Employer's details. Please" & _
                " select an existing Employer.", _
            MsgBoxStyle.Exclamation, _
            "iManagement - Invalid or incomplete Information")

                datDelete = Nothing
                objLogin = Nothing
                Exit Function
            End If

            If MsgBox("Are you sure you want to delete the Employer's" & _
                " Details?. " & _
                    Chr(10) & "The Deletion will include the Employer's " & _
                        "related Bank Accounts!", _
                            MsgBoxStyle.YesNo, _
                "iManagement - Delete Records?") Then

                datDelete = Nothing
                objLogin = Nothing
                Exit Function
            End If

            strDeleteQuery = "DELETE * FROM Employers " & _
            " LEFT JOIN EmployerSalaryAccount ON " & _
            " EmployerSalaryAccount.EmployerID = " & _
            " Employers.EmployerID " & _
            " WHERE Name = '" & strEmployerName & "'"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strDeleteQuery, _
                            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Employer Details Deleted")

                Return True
            Else
                MsgBox("'Employer delete' action failed", _
                    MsgBoxStyle.Exclamation, " Employer Deletion failed")
            End If


        Catch ex As Exception

        End Try

    End Function

    Public Function Update(ByVal strUpQuery As String) As Boolean
        'Updates country details of the country with the country code

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If Trim(strEmployerName) = "" And lEmployerID = 0 _
                     Then

            MsgBox("Cannot update the Employer's details. Please select an existing Employer.", _
        MsgBoxStyle.Exclamation, _
        "iManagement - Invalid or incomplete Information")

            datUpdated = Nothing
            objLogin = Nothing
            Exit Function
        End If


        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
        strUpdateQuery, _
        datUpdated)

        objLogin.CloseDb()

        If bUpdateSuccess = True Then

            MsgBox("Record Updated Successfully", _
                MsgBoxStyle.Information, _
                "iManagement - Employer Details Updated")
            Return True
        Else

            MsgBox("Update of employer details failed", MsgBoxStyle.Information, _
                "iManagement - Data update failed")

        End If

    End Function


#End Region

End Class
