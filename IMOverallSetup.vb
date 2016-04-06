Option Explicit On 
'Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMOverallSetup

#Region "PrivateOverallSetupVariables"

    Private lOrganizationID As Long
    Private strSysOrganizationName As String
    Private strTaxInformationNumber As String
    Private strCompanyRegistrationNumber As String
    Private dtCompanyRegistrationDate As Date
    Private dtSysCreationDate As Date
    Private strPhysicalAddress As String
    Private strCountryCode As String
    Private strCityCode As String
    Private strTown As String
    Private strPostAddress As String
    Private strPostCode As String
    Private strPostCountry As String
    Private strPostCity As String
    Private lCompanyTypeID As Long
    Private bVATRegistered As Boolean
    Private strPhone1 As String
    Private strPhone2 As String
    Private strFaxNumber As String
    Private strDefaultCurrency As String
    Private strSysCompanyStatus As String
    Private lRoundingOff As Long

    Private strSysOrganizationFileName As String
    Private strSysOrganizationPath As String

    Private bAutomatedCustNum As Boolean

#End Region

#Region "Properties"

    'Automated Customer Number
    Public Property AutomatedCustNum() As Boolean

        Get
            Return bAutomatedCustNum
        End Get

        Set(ByVal Value As Boolean)
            bAutomatedCustNum = Value
        End Set

    End Property

    'OrganizationFileName
    Public Property OrganizationFileName() As String

        Get
            Return strSysOrganizationFileName
        End Get

        Set(ByVal Value As String)
            strSysOrganizationFileName = Value
        End Set

    End Property

    'OrganizationPath
    Public Property OrganizationPath() As String

        Get
            Return strSysOrganizationPath
        End Get

        Set(ByVal Value As String)
            strSysOrganizationPath = Value
        End Set

    End Property
    'OrganizationID
    Public Property OrganizationID() As Long

        Get
            Return lOrganizationID
        End Get

        Set(ByVal Value As Long)
            lOrganizationID = Value
        End Set

    End Property

    'OrganizationName
    Public Property OrganizationName() As String

        Get
            Return strSysOrganizationName
        End Get

        Set(ByVal Value As String)
            strSysOrganizationName = Value
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

    Public Property SysCreationDate() As Date

        'USED TO SET AND RETRIEVE THE SalaryTypeDescription (STRING)
        Get
            Return dtSysCreationDate
        End Get

        Set(ByVal Value As Date)
            dtSysCreationDate = Value
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

    Public Property PostAddress() As String

        Get
            Return strPostAddress
        End Get

        Set(ByVal Value As String)
            strPostAddress = Value
        End Set

    End Property

    Public Property PostCode() As String

        Get
            Return strPostCode
        End Get

        Set(ByVal Value As String)
            strPostCode = Value
        End Set

    End Property

    'PostCountry
    Public Property PostCountry() As String

        Get
            Return strPostCountry
        End Get

        Set(ByVal Value As String)
            strPostCountry = Value
        End Set

    End Property

    'PostCity
    Public Property PostCity() As String

        Get
            Return strPostCity
        End Get

        Set(ByVal Value As String)
            strPostCity = Value
        End Set

    End Property

    'CompanyTypeID
    Public Property CompanyTypeID() As Long

        Get
            Return lCompanyTypeID
        End Get

        Set(ByVal Value As Long)
            lCompanyTypeID = Value
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

    'FaxNumber
    Public Property FaxNumber() As String

        Get
            Return strFaxNumber
        End Get

        Set(ByVal Value As String)
            strFaxNumber = Value
        End Set

    End Property

    'DefaultCurrency
    Public Property DefaultCurrency() As String

        Get
            Return strDefaultCurrency
        End Get

        Set(ByVal Value As String)
            strDefaultCurrency = Value
        End Set

    End Property

    'DefaultCurrency
    Public Property RoundingOff() As Long

        Get
            Return lRoundingOff
        End Get

        Set(ByVal Value As Long)
            lRoundingOff = Value
        End Set

    End Property

    'SysCompanyStatus
    Public Property SysCompanyStatus() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strSysCompanyStatus
        End Get

        Set(ByVal Value As String)
            strSysCompanyStatus = Value
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

    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Function SaveDefaultCurrency(ByVal DisplayErrorMessages As Boolean, _
            ByVal DisplaySuccessMessages As Boolean, _
                ByVal DisplayFailureMessages As Boolean) As Boolean

        'Saves a new base organization
        Try


            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

         
            If lOrganizationID = 0 Then
                MsgBox("You must provide an existing organization.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Invalid or incomplete information")

                objLogin = Nothing
                datSaved = Nothing
                Exit Function

            End If

            If Find("SELECT * FROM CompanyMaster WHERE OrganizationID = " _
                & lOrganizationID, _
                    False, False, False) = False Then

                MsgBox("Please select an existing company name (Default Currency Value Not Updated)", _
                MsgBoxStyle.Critical, "iManagement - missing company details")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            Else

                Update("UPDATE CompanyMaster SET" & _
                    " DefaultCurrency = '" & strDefaultCurrency & _
                    "' WHERE OrganizationID = " & lOrganizationID, False)

            End If

            objLogin = Nothing
            datSaved = Nothing

            Exit Function



        Catch ex As Exception
            If DisplayErrorMessages = True Then
                MsgBox(ex.Source, MsgBoxStyle.Critical, _
                    "iManagement - Database or system error")
            End If

        End Try

    End Function

    Public Function SaveRoundingOff(ByVal DisplayErrorMessages As Boolean, _
            ByVal DisplaySuccessMessages As Boolean, _
                ByVal DisplayFailureMessages As Boolean)

        'Saves a new base organization
        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin


            If lOrganizationID = 0 Then
                MsgBox("You must provide an existing organization.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Invalid or incomplete information")

                objLogin = Nothing
                datSaved = Nothing
                Exit Function

            End If

            If Find("SELECT * FROM CompanyMaster WHERE OrganizationID = " _
                & lOrganizationID, _
                    False, False, False) = False Then

                MsgBox("Please select an existing company name (Rounding Off Value Not Updated)", _
                MsgBoxStyle.Critical, "iManagement - missing company details")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            Else

                Update("UPDATE CompanyMaster SET" & _
                    " RoundingOff = " & lRoundingOff & _
                    " WHERE OrganizationID = " & lOrganizationID, False)

            End If

            objLogin = Nothing
            datSaved = Nothing

            Exit Function

        Catch ex As Exception
            If DisplayErrorMessages = True Then
                MsgBox(ex.Source, MsgBoxStyle.Critical, _
                    "iManagement - Database or system error")
            End If

        End Try

    End Function

    'Save informaiton
    Public Sub Save(ByVal DisplayErrorMessages As Boolean, _
            ByVal DisplaySuccessMessages As Boolean, _
                ByVal DisplayFailureMessages As Boolean)

        'Saves a new base organization
        Try


            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String
            Dim dirLocation As File

            If Trim(strOrganizationName) = "" Then

                If Trim(strSysOrganizationName) = "" Then
                    MsgBox("You must provide an organization name", _
                        MsgBoxStyle.Critical, _
                            "iManagement - Invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing
                    Exit Sub

                End If

                If Trim(strTaxInformationNumber) = "" Then
                    MsgBox("You must provide the organization's or persons Tax Information Number", _
                        MsgBoxStyle.Critical, _
                            "iManagement - Invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing
                    Exit Sub

                End If

                If Trim(strCompanyRegistrationNumber) = "" Then
                    MsgBox("You must provide the organization's or persons Company Registration Number", _
                        MsgBoxStyle.Critical, _
                            "iManagement - Invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing
                    Exit Sub

                End If

                If Trim(strPhysicalAddress) = "" Then
                    MsgBox("You must provide the organization's or persons Physical Address", _
                        MsgBoxStyle.Critical, _
                            "iManagement - Invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing
                    Exit Sub

                End If


                If Trim(strSysOrganizationFileName) = "" Then
                    MsgBox("Please select the location to save the file.", _
                        MsgBoxStyle.Critical, _
                            "iManagement - Select an existing directory")

                    objLogin = Nothing
                    datSaved = Nothing
                    Exit Sub

                End If

                If dirLocation.Exists(strSysOrganizationFileName) = False Then
                    MsgBox("The location selected does not exist.", _
                        MsgBoxStyle.Critical, _
                            "iManagement - Select an existing directory")

                    objLogin = Nothing
                    datSaved = Nothing
                    Exit Sub

                End If
            End If


            If Find("SELECT * FROM CompanyMaster WHERE OrganizationName = '" _
                & Trim(strSysOrganizationName) & "'", _
                    False, False, False) = True Then

                If MsgBox("Do you want to update the details?" & _
                " Remember updating is amending details of an existing record," _
                 & Chr(10) & "that is, this organization name already exists.", _
                    MsgBoxStyle.YesNo, "iManagement - Update existing company details?") _
                        = MsgBoxResult.Yes Then

                    Update("UPDATE CompanyMaster SET" & _
                        "    CompRegNum = '" & Trim(strCompanyRegistrationNumber) & _
                        "' , TaxInformationNumber = '" & Trim(strTaxInformationNumber) & _
                        "' , CompanyRegDate = '" & dtCompanyRegistrationDate & _
                        "' , PhysicalAddress = '" & Trim(strPhysicalAddress) & _
                        "' , CountryCode = '" & Trim(strCountryCode) & _
                        "' , CityCode = '" & Trim(strCityCode) & _
                        "' , Town = '" & Trim(strTown) & _
                        "' , PostAddress = '" & Trim(strPostAddress) & _
                        "' , PostCode = '" & Trim(strPostCode) & _
                        "' , PostCountry = '" & Trim(strPostCountry) & _
                        "' , PostCity = '" & Trim(strPostCity) & _
                        "' , CompanyTypeID =  " & lCompanyTypeID & _
                        " ,  VATRegistered = " & bVATRegistered & _
                        " , Phone1 = '" & Trim(strPhone1) & _
                        "' , Phone2 = '" & Trim(strPhone2) & _
                        "' , FaxNumber = '" & Trim(strFaxNumber) & _
                        "' , SysCompanyStatus = '" & Trim(strSysCompanyStatus) & _
                        "' WHERE OrganizationName = '" & strSysOrganizationName & "'", True)

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub

            End If


            If Find("SELECT * FROM CompanyMaster WHERE OrganizationName LIKE '%" _
            & Trim(strSysOrganizationName) & "%'", False, False, False) = True Then

                If MsgBox("There is an existing organization that has a name which is" & _
                Chr(10) & " almost similar name to the one you provided." & _
                    Chr(10) & " You may continue but it is advised that you " & _
                        Chr(10) & "confirm first before continuing" & _
                            Chr(10) & "Do you want to confirm?", _
                                MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, _
                                    "iManagement - Cannot save this new Base organization") = _
                                        MsgBoxResult.Yes Then

                    MsgBox("Select File > Open to see the list of existing organizations", _
                        MsgBoxStyle.Information, _
                            "iManagement - Advice on verification procedure")

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Sub
                End If

            End If


            '
            If Find("SELECT * FROM CompanyMaster WHERE TaxInformationNumber = '" _
            & Trim(strTaxInformationNumber) & "'", False, False, False) = True Then

                MsgBox("The Tax Information Number (PIN or TIN) you have provided exists." & _
                    " Please provide Tax Information Number that does not exist", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Cannot save new Base organization")

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub
            End If


            If Find("SELECT * FROM CompanyMaster WHERE CompRegNum = '" _
                & Trim(strCompanyRegistrationNumber) & "'", _
                        False, False, False) = True Then

                MsgBox("The Company Registration Number you have provided exists." & _
                    " Please provide Company Registration Number that does not exist", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Cannot save new Base organization")

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub
            End If

            If strPostAddress <> "" Or _
                strPostCity <> "" Or _
                    strPostCountry <> "" Then

                If strPostAddress = "" Or strPostCity = "" _
                    Or strPostCountry = "" Then

                    MsgBox("You have entered invalid postal details." & _
                    Chr(10) & "You must enter the Post Address, Post Code, and" & _
                        Chr(10) & "PostCountry for a valid postal entry")

                End If
            End If

            strInsertInto = "INSERT INTO CompanyMaster (" & _
                "OrganizationName," & _
                "TaxInformationNumber," & _
                "CompRegNum," & _
                "CompanyRegDate," & _
                "SysCreationDate," & _
                "PhysicalAddress," & _
                "CountryCode," & _
                "CityCode," & _
                "Town," & _
                "PostAddress," & _
                "PostCode," & _
                "PostCountry," & _
                "PostCity," & _
                "CompanyTypeID," & _
                "VATRegistered," & _
                "Phone1," & _
                "Phone2," & _
                "FaxNumber," & _
                "DefaultCurrency," & _
                "SysCompanyStatus" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & strSysOrganizationName & _
            "','" & strTaxInformationNumber & _
            "','" & strCompanyRegistrationNumber & _
            "','" & dtCompanyRegistrationDate & _
            "','" & dtSysCreationDate & _
            "','" & strPhysicalAddress & _
            "','" & strCountryCode & _
            "','" & strCityCode & _
            "','" & strTown & _
            "','" & strPostAddress & _
            "','" & strPostCode & _
            "','" & strPostCountry & _
            "','" & strPostCity & _
            "', " & lCompanyTypeID & _
            " , " & bVATRegistered & _
            " ,'" & strPhone1 & _
            "','" & strPhone2 & _
            "','" & strFaxNumber & _
            "','" & strDefaultCurrency & _
            "','" & strSysCompanyStatus & _
            "')"

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Organization Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - Record Saved")

                End If

            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save New Organization' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save New Record Failed")
                End If

            End If

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception
            If DisplayErrorMessages = True Then
                MsgBox(ex.Source, MsgBoxStyle.Critical, _
                    "iManagement - Database or system error")
            End If

        End Try

    End Sub

    'Find Informaiton
    Public Function Find(ByVal strQuery As String, _
        ByVal bReturnValues As Boolean, _
            ByVal bReturnOrganizationFileName As Boolean, _
                ByVal bReturnOrganizationDetails As Boolean) As Boolean

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
                    objLogin = Nothing
                    datRetData = Nothing


                    Exit Function

                End If


                If bReturnValues = True Then

                    For Each myDataRows In myDataTables.Rows

                        If bReturnOrganizationFileName = True Then
                            strSysOrganizationFileName = _
                                    myDataRows("FileName").ToString
                            strSysOrganizationPath = _
                                    myDataRows("Path").ToString()

                        End If


                        If bReturnOrganizationDetails = True Then

                            If bReturnOrganizationFileName = True Then
                                lOrganizationID = _
                                    myDataRows("CompanyMaster.OrganizationID")
                            Else
                                lOrganizationID = _
                                myDataRows("OrganizationID")
                            End If


                            strSysOrganizationName = _
                                myDataRows("OrganizationName").ToString
                            strTaxInformationNumber = _
                                myDataRows("TaxInformationNumber").ToString
                            strCompanyRegistrationNumber = _
                                myDataRows("CompRegNum").ToString
                            dtCompanyRegistrationDate = _
                                Format(myDataRows _
                                    ("CompanyRegDate"), "dd/MMM/yyyy")
                            dtSysCreationDate = _
                                Format(myDataRows _
                                    ("SysCreationDate"), "dd/MMM/yyyy")
                            strPhysicalAddress = _
                                myDataRows("PhysicalAddress").ToString
                            strCountryCode = _
                                myDataRows("CountryCode").ToString
                            strCityCode = _
                                myDataRows("CityCode").ToString
                            strTown = _
                                myDataRows("Town").ToString
                            strPostAddress = _
                                myDataRows("PostAddress").ToString
                            strPostCode = _
                                myDataRows("PostCode").ToString
                            strPostCountry = _
                                myDataRows("PostCountry").ToString
                            strPostCity = _
                                myDataRows("PostCity").ToString
                            lCompanyTypeID = _
                                myDataRows("CompanyTypeID")
                            bVATRegistered = _
                                myDataRows("VATRegistered")
                            strPhone1 = _
                                myDataRows("Phone1").ToString
                            strPhone2 = _
                                myDataRows("Phone2").ToString
                            strFaxNumber = _
                                myDataRows("FaxNumber").ToString
                            strDefaultCurrency = _
                                myDataRows("DefaultCurrency").ToString
                            lRoundingOff = _
                                myDataRows("RoundingOff").ToString
                            strSysCompanyStatus = _
                                myDataRows("SysCompanyStatus").ToString

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

            If lOrganizationID <> 0 Then

                objLogin.ConnectString = strAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - Organization Details Deleted")
                Else
                    MsgBox("'Organization delete' action failed", _
                        MsgBoxStyle.Exclamation, " Organization Deletion failed")
                End If
            Else

                MsgBox("Cannot Delete. Please select an existing salary type", _
                        MsgBoxStyle.Exclamation, "iManagement -Missing Information")

            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Update(ByVal strUpQuery As String, _
        ByVal bDisplayConfirm As Boolean)

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
                If bDisplayConfirm = True Then
                    MsgBox("Overall Organization Setup Details Updated Successfully", MsgBoxStyle.Information, _
                        "iManagement - Record Updated")
                End If
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub

#End Region

End Class
