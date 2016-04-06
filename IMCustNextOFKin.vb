Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCustNextOFKin

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
    Private strCityOfResidence As String
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

    Public Shadows Property CityOfResidence() As String


        Get
            Return Trim(strCityOfResidence)
        End Get

        Set(ByVal Value As String)
            strCityOfResidence = Value
        End Set

    End Property

    Public Shadows Property NOKID() As Long


        Get
            Return lNOKID
        End Get

        Set(ByVal Value As Long)
            lNOKID = Value
        End Set

    End Property

    Public Shadows Property Surname() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strSurname)
        End Get

        Set(ByVal Value As String)
            strSurname = Value
        End Set

    End Property

    Public Shadows Property MiddleName() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strMiddleName)
        End Get

        Set(ByVal Value As String)
            strMiddleName = Value
        End Set

    End Property

    Public Shadows Property FirstName() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strFirstName)
        End Get

        Set(ByVal Value As String)
            strFirstName = Value
        End Set

    End Property

    Public Shadows Property OtherName() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strOtherName)
        End Get

        Set(ByVal Value As String)
            strOtherName = Value
        End Set

    End Property

    Public Shadows Property Sex() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strSex)
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
            Return Trim(strCountryOfBirth)
        End Get

        Set(ByVal Value As String)
            strCountryOfBirth = Value
        End Set

    End Property

    Public Shadows Property CountryOfCitizenship() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strCountryofCitizenship)
        End Get

        Set(ByVal Value As String)
            strCountryofCitizenship = Value
        End Set

    End Property

    Public Shadows Property CountryOfResidence() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strCountryOfResidence)
        End Get

        Set(ByVal Value As String)
            strCountryOfResidence = Value
        End Set

    End Property

    Public Shadows Property PhysicalAddress() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strPhysicalAddress)
        End Get

        Set(ByVal Value As String)
            strPhysicalAddress = Value
        End Set

    End Property

    Public Shadows Property PostalAddress() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strPostalAddress)
        End Get

        Set(ByVal Value As String)
            strPostalAddress = Value
        End Set

    End Property

    Public Shadows Property PostalCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strPostCode)
        End Get

        Set(ByVal Value As String)
            strPostCode = Value
        End Set

    End Property

    Public Shadows Property PostCountryCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strPostCountryCode)
        End Get

        Set(ByVal Value As String)
            strPostCountryCode = Value
        End Set

    End Property

    Public Shadows Property PostCityCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strPostCityCode)
        End Get

        Set(ByVal Value As String)
            strPostCityCode = Value
        End Set

    End Property

    Public Shadows Property PostTown() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strPostTown)
        End Get

        Set(ByVal Value As String)
            strPostTown = Value
        End Set

    End Property

    Public Shadows Property Phone1() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strPhone1)
        End Get

        Set(ByVal Value As String)
            strPhone1 = Value
        End Set

    End Property

    Public Shadows Property Phone2() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strPhone2)
        End Get

        Set(ByVal Value As String)
            strPhone2 = Value
        End Set

    End Property

    Public Shadows Property Phone3() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strPhone3)
        End Get

        Set(ByVal Value As String)
            strPhone3 = Value
        End Set

    End Property

    Public Shadows Property EmailAddress() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strEmailAddress)
        End Get

        Set(ByVal Value As String)
            strEmailAddress = Value
        End Set

    End Property

    Public Shadows Property PINNo() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return Trim(strPINNo)
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

        'Saves a new Next Of Kin
        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String


            If Trim(strFirstName) <> "" And _
                Trim(strSurname) <> "" And _
                    dtDateOfBirth > Now() _
                                Then

                MsgBox("To save the Next Of Kin you must provide:" & _
                    Chr(10) & "1 : The Next Of Kin's First Name" & _
                        Chr(10) & "2: The Next Of Kin's Surname" & _
                            Chr(10) & "3: The Next Of Kin's Date Of birth")

                datSaved = Nothing
                objLogin = Nothing

                Exit Sub

            End If


                'Check if there is an existing series with this name
                If Find("SELECT * FROM CustomerNextOfKin WHERE  CustomerNo = " _
                            & CustomerNo & " AND NOKID = " & lNOKID, False) = True Then

                If MsgBox("The NOK Details already exists." & _
                Chr(10) & "Do you want to update the details?", _
                        MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                MsgBoxResult.Yes Then


                    Update("UPDATE CustomerNextOfKin SET " & _
                                "Surname = '" & Trim(strSurname) & _
                                "' , FirstName = '" & Trim(strFirstName) & _
                                "' , MiddleName = '" & Trim(strMiddleName) & _
                                "' , OtherName = '" & Trim(strOtherName) & _
                                "' , Sex = '" & Trim(strSex) & _
                                "' , DateOfBirth = '" & dtDateOfBirth & _
                                "' , CountryOfBirth = '" & Trim(strCountryOfBirth) & _
                                "' , CountryofCitizenship = '" & Trim(strCountryofCitizenship) & _
                                "' , CountryOfResidence = '" & Trim(strCountryOfResidence) & _
                                "' , CityOfResidence = '" & Trim(strCityOfResidence) & _
                                "' , PhysicalAddress = '" & Trim(strPhysicalAddress) & _
                                "' , PostalAddress = '" & Trim(strPostalAddress) & _
                                "' , PostCode = '" & Trim(strPostCode) & _
                                "' , PostCountryCode = '" & Trim(strPostCountryCode) & _
                                "' , PostCityCode = '" & Trim(strPostCityCode) & _
                                "' , PostTown = '" & Trim(strPostTown) & _
                                "' , Phone1 = '" & Trim(strPhone1) & _
                                "' , Phone2 = '" & Trim(strPhone2) & _
                                "' , Phone3 = '" & Trim(strPhone3) & _
                                "' , PINNo = '" & Trim(strPINNo) & _
                                    "' WHERE  CustomerNo = " _
                                        & CustomerNo & " AND NOKID = " & lNOKID)

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub
            End If

                strInsertInto = "INSERT INTO CustomerNextOfKin (" & _
                    "CustomerNo," & _
                    "Surname, " & _
                    "FirstName," & _
                    "MiddleName," & _
                    "OtherName," & _
                    "Sex," & _
                    "DateOfBirth," & _
                    "CountryOfBirth," & _
                    "CountryOfCitizenship," & _
                    "CountryOfResidence," & _
                    "CityOfResidence," & _
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
                    "PINNo" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "(" & CustomerNo & _
                        ", '" & Trim(strSurname) & _
                        "', '" & Trim(strFirstName) & _
                        "', '" & Trim(strMiddleName) & _
                        "', '" & Trim(strOtherName) & _
                        "', '" & Trim(strSex) & _
                        "', '" & dtDateOfBirth & _
                        "', '" & Trim(strCountryOfBirth) & _
                        "', '" & Trim(strCountryofCitizenship) & _
                        "', '" & Trim(strCountryOfResidence) & _
                        "', '" & Trim(strCityOfResidence) & _
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
                        "')"

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strSaveQuery, _
                datSaved)

                objLogin.CloseDb()

                If bSaveSuccess = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - Customer Next Of Kin Saved")

                Else

                    MsgBox("'Save Customer Next Of Kin action failed." & _
                        " Make sure all mandatory details are entered", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Customer Next Of Kin Addition Failed")

                End If

        Catch ex As Exception

        End Try

    End Sub

    Public Shadows Function Find(ByVal strQuery As String, _
            ByVal ReturnValues As Boolean) As Boolean

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

                    If ReturnValues = True Then
                        CustomerNo = myDataRows("CustomerNo")
                        strSurname = myDataRows("Surname").ToString()
                        strFirstName = myDataRows("FirstName").ToString()
                        strMiddleName = myDataRows("MiddleName").ToString()
                        strOtherName = myDataRows("OtherName").ToString()
                        strSex = myDataRows("Sex").ToString()
                        dtDateOfBirth = myDataRows("DateOfBirth")
                        strCountryOfBirth = myDataRows("CountryOfBirth").ToString()
                        strCountryofCitizenship = myDataRows("CountryofCitizenship").ToString()
                        strCountryOfResidence = myDataRows("CountryOfResidence").ToString()
                        strCityOfResidence = myDataRows("CityOfResidence").ToString()
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


        If lNOKID = 0 Then

            MsgBox("Please select an existing NOK Identifier or record.", _
                MsgBoxStyle.Exclamation, _
                    "iManagement - invalid or incomplete informaiton")

            datDelete = Nothing
            objLogin = Nothing

            Exit Sub
        End If

        'Confirm deletion
        If MsgBox("Are you sure you want to delete this record?" _
        , MsgBoxStyle.YesNo, "iManagement - Delete this record?") _
        = MsgBoxResult.No Then

            datDelete = Nothing
            objLogin = Nothing

            Exit Sub
        End If

        strDeleteQuery = "DELETE * FROM CustomerNextOfKin WHERE  NOKID = " _
                       & lNOKID

        objLogin.connectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
        datDelete)

        objLogin.CloseDb()

        If bDelSuccess = True Then

            MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                "iManagement -  Customer Next Of Kin action Details Deleted")

        Else

            MsgBox("'Delete Customer Next Of Kin action failed", _
                MsgBoxStyle.Exclamation, "  Customer Next Of Kin action Deletion failed")

        End If


    End Sub

    Public Shadows Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If CustomerNo <> 0 _
                        Then

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -   Customer Next Of Kin action Details Updated")
            End If

        End If

    End Sub


#End Region

End Class
