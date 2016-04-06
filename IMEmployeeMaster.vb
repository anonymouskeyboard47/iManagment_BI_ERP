Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMEmployee

#Region "PrivateCustomerVariables"
    Private dtRegistrationDate As Date
    Private lCustomerNo As String
    Private strTitle As String
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
    Private strCustType As String
    Private strCustPriorityID As String
    Private strCustPhoto As String
    Private strCustScanIDSide1 As String
    Private strCustScanIDSide2 As String


#End Region

#Region "Properties"

    Public Property RegistrationDate() As Date

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return dtRegistrationDate
        End Get

        Set(ByVal Value As Date)
            dtRegistrationDate = Value
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

    Public Property Title() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strTitle
        End Get

        Set(ByVal Value As String)
            strTitle = Value
        End Set

    End Property

    Public Property Surname() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strSurname
        End Get

        Set(ByVal Value As String)
            strSurname = Value
        End Set

    End Property

    Public Property MiddleName() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strMiddleName
        End Get

        Set(ByVal Value As String)
            strMiddleName = Value
        End Set

    End Property

    Public Property FirstName() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strFirstName
        End Get

        Set(ByVal Value As String)
            strFirstName = Value
        End Set

    End Property

    Public Property OtherName() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strOtherName
        End Get

        Set(ByVal Value As String)
            strOtherName = Value
        End Set

    End Property

    Public Property Sex() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strSex
        End Get

        Set(ByVal Value As String)
            strSex = Value
        End Set

    End Property

    Public Property DateOfBirth() As Date

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return dtDateOfBirth
        End Get

        Set(ByVal Value As Date)
            dtDateOfBirth = Value
        End Set

    End Property

    Public Property CountryOfBirth() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCountryOfBirth
        End Get

        Set(ByVal Value As String)
            strCountryOfBirth = Value
        End Set

    End Property

    Public Property CountryOfCitizenship() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCountryofCitizenship
        End Get

        Set(ByVal Value As String)
            strCountryofCitizenship = Value
        End Set

    End Property

    Public Property CountryOfResidence() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCountryOfResidence
        End Get

        Set(ByVal Value As String)
            strCountryOfResidence = Value
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

    Public Property PostalAddress() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostalAddress
        End Get

        Set(ByVal Value As String)
            strPostalAddress = Value
        End Set

    End Property

    Public Property PostalCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostCode
        End Get

        Set(ByVal Value As String)
            strPostCode = Value
        End Set

    End Property

    Public Property PostCountryCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostCountryCode
        End Get

        Set(ByVal Value As String)
            strPostCountryCode = Value
        End Set

    End Property

    Public Property PostCityCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostCityCode
        End Get

        Set(ByVal Value As String)
            strPostCityCode = Value
        End Set

    End Property

    Public Property PostTown() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostTown
        End Get

        Set(ByVal Value As String)
            strPostTown = Value
        End Set

    End Property

    Public Property Phone1() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhone1
        End Get

        Set(ByVal Value As String)
            strPhone1 = Value
        End Set

    End Property

    Public Property Phone2() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhone2
        End Get

        Set(ByVal Value As String)
            strPhone2 = Value
        End Set

    End Property

    Public Property Phone3() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhone3
        End Get

        Set(ByVal Value As String)
            strPhone3 = Value
        End Set

    End Property

    Public Property EmailAddress() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strEmailAddress
        End Get

        Set(ByVal Value As String)
            strEmailAddress = Value
        End Set

    End Property

    Public Property PINNo() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPINNo
        End Get

        Set(ByVal Value As String)
            strPINNo = Value
        End Set

    End Property

    Public Property CustomerType() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCustType
        End Get

        Set(ByVal Value As String)
            strCustType = Value
        End Set

    End Property

    Public Property PriorityID() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCustPriorityID
        End Get

        Set(ByVal Value As String)
            strCustPriorityID = Value
        End Set

    End Property

    Public Property CustScanIDSide1() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCustScanIDSide1
        End Get

        Set(ByVal Value As String)
            strCustScanIDSide1 = Value
        End Set

    End Property

    Public Property CustScanIDSide2() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCustScanIDSide2
        End Get

        Set(ByVal Value As String)
            strCustScanIDSide2 = Value
        End Set

    End Property

    Public Property CustomerPhoto() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCustPhoto
        End Get

        Set(ByVal Value As String)
            strCustPhoto = Value
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
        strTitle = ""
        strSurname = ""
        strFirstName = ""
        strMiddleName = ""
        strOtherName = ""
        strSex = ""
        dtDateOfBirth = Now()
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
        strCustType = ""
        strCustPriorityID = ""
        strCustPhoto = ""
        strCustScanIDSide1 = ""
        strCustScanIDSide2 = ""

    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Function Save() As Boolean

        'Saves a new country name
        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String

            If Trim(strFirstName) <> "" And _
                Trim(strSurname) <> "" And _
                    dtDateOfBirth < Now() And _
                        strCustType <> "" _
                                Then
                If Trim(strEmailAddress) <> "" Then
                    If InStr(strEmailAddress, "@") = 0 Then
                        MsgBox("Please insert a valid email address", _
                            MsgBoxStyle.Exclamation, _
                            "iManagement - Invalid or incomplete data provided")

                        objLogin = Nothing
                        datSaved = Nothing

                        Return False
                        Exit Function
                    End If
                End If

                strInsertInto = "INSERT INTO CustomerMaster (" & _
                    "Title," & _
                    "Surname, " & _
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
                    "Phone1," & _
                    "Phone2," & _
                    "Phone3," & _
                    "EmailAddress," & _
                    "PINNo," & _
                    "CustType," & _
                    "CustPriorityID," & _
                    "CustPhoto," & _
                    "CustScanIDSide1," & _
                    "CustScanIDSide2" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "('" & Trim(strTitle) & _
                        "', '" & Trim(strSurname) & _
                        "', '" & Trim(strFirstName) & _
                        "', '" & Trim(strMiddleName) & _
                        "', '" & Trim(strOtherName) & _
                        "', '" & Trim(strSex) & _
                        "', '" & dtDateOfBirth & _
                        "', '" & Trim(strCountryOfBirth) & _
                        "', '" & Trim(strCountryofCitizenship) & _
                        "', '" & Trim(strCountryOfResidence) & _
                        "', '" & Trim(strPhysicalAddress) & _
                        "', '" & Trim(strPostalAddress) & _
                        "', '" & Trim(strPostCode) & _
                        "', '" & Trim(strPostCountryCode) & _
                        "', '" & Trim(strPostCityCode) & _
                        "', '" & Trim(strPostTown) & _
                        "', '" & Trim(strPhone1) & _
                        "', '" & Trim(strPhone2) & _
                        "', '" & Trim(strPhone3) & _
                        "', '" & Trim(strEmailAddress) & _
                        "', '" & Trim(strPINNo) & _
                        "', '" & Trim(strCustType) & _
                        "', '" & Trim(strCustPriorityID) & _
                        "', '" & Trim(strCustPhoto) & _
                        "', '" & Trim(strCustScanIDSide1) & _
                        "', '" & Trim(strCustScanIDSide2) & _
                        "')"

                objLogin.connectString = strOrgAccessConnString

                If objLogin.ConnectToDatabase = False Then
                    Return False
                    Exit Function
                End If

                bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strSaveQuery, _
                datSaved)

                objLogin.CloseDb()

                If bSaveSuccess = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - Customer Saved")
                    Return True

                Else

                    MsgBox("'Save Customer' action failed." & _
                        " Make sure all mandatory details are entered", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Customer Addition Failed")

                End If

            End If
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function Find(ByVal strQuery As String) As Boolean
        Try


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
                        strTitle = myDataRows("Title").ToString()
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
                        strCustType = myDataRows("CustType").ToString()
                        strCustPriorityID = myDataRows("CustPriorityID").ToString()
                        strCustPhoto = myDataRows("CustPhoto")
                        strCustScanIDSide1 = myDataRows("CustScanIDSide1")
                        strCustScanIDSide2 = myDataRows("CustScanIDSide2")
                        dtRegistrationDate = myDataRows("RegistrationDate")

                    Next

                Next

                Return True
            Else
                Return False
            End If

        Catch ex As Exception

        End Try

    End Function

    Public Sub Delete(ByVal strDelQuery As String)
        Try
            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = strDelQuery

            If lCustomerNo <> 0 Or _
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

            If lCustomerNo <> 0 Or _
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

        Catch ex As Exception

        End Try

    End Sub

    Public Function FillControl _
            (ByVal strFillConnString As String, _
                ByVal strTSQL As String, _
                    ByVal strValueField As String, _
                        ByVal strTextField As String) _
                            As String()


        Dim datFillData As DataSet
        Dim bReturnedSuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim strTextFieldData() As String
        Dim i As Integer
        Dim objLogin As IMLogin = New IMLogin

        Try

            datFillData = New DataSet

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            'The db is okay now get the recordset
            bReturnedSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                strTSQL, datFillData)

            objLogin.CloseDb()

            If datFillData Is Nothing Then
                Exit Function
            End If

            For Each myDataTables In datFillData.Tables

                'Check if there is any data. If not exit
                If myDataTables.Rows.Count = 0 Then
                    'Return an empty array
                    ReDim strTextFieldData(1)
                    strTextFieldData(0) = ""
                    Return strTextFieldData

                    Exit Function
                Else
                    'Resize the array
                    ReDim strTextFieldData(myDataTables.Rows.Count)

                End If

                i = 0
                For Each myDataRows In myDataTables.Rows
                    strTextFieldData(i) = myDataRows(0).ToString()
                    i = i + 1
                Next

            Next

            Return strTextFieldData
            datFillData.Dispose()

        Catch ex As Exception

        End Try

    End Function


#End Region

End Class
