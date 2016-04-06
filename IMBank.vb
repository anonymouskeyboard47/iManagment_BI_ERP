Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMBank

#Region "PrivateBankVariables"
    Private lBankID As Long
    Private strBankName As String
    Private strBranchName As String
    Private strBranchEFTCode As String
    Private strPhysicalAddress As String
    Private strCountryCode As String
    Private strCityCode As String
    Private strPostalAddress As String
    Private strPostCode As String
    Private strPostCountryCode As String
    Private strPostCityCode As String
    Private strPostTownCode As String
    Private strBankSummaryDetails As String
    Private strPhone1 As String

#End Region

#Region "Properties"

    Public Property BankID() As Long

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return lBankID
        End Get

        Set(ByVal Value As Long)
            lBankID = Value
        End Set

    End Property

    Public Property BankName() As String

        'USED TO SET AND RETRIEVE THE BANK NAME (STRING)
        Get
            Return strBankName
        End Get

        Set(ByVal Value As String)
            strBankName = Value
        End Set

    End Property

    Public Property BranchName() As String

        'USED TO SET AND RETRIEVE THE BRANCH NAME (STRING)
        Get
            Return strBranchName
        End Get

        Set(ByVal Value As String)
            strBranchName = Value
        End Set

    End Property

    Public Property BranchEFTCode() As String

        'USED TO SET AND RETRIEVE THE BRANCH EFT CODE (STRING)
        Get
            Return strBranchEFTCode
        End Get

        Set(ByVal Value As String)
            strBranchEFTCode = Value
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

    Public Property CountryCode() As String

        'USED TO SET AND RETRIEVE THE COUNTRY CODE (STRING)
        Get
            Return strCountryCode
        End Get

        Set(ByVal Value As String)
            strCountryCode = Value
        End Set

    End Property

    Public Property CityCode() As String

        'USED TO SET AND RETRIEVE THE CITY CODE
        Get
            Return strCityCode
        End Get

        Set(ByVal Value As String)
            strCityCode = Value
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
            Return strPostCode
        End Get

        Set(ByVal Value As String)
            strPostCode = Value
        End Set

    End Property

    Public Property PostCountryCode() As String

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return strPostCountryCode
        End Get

        Set(ByVal Value As String)
            strPostCountryCode = Value
        End Set

    End Property

    Public Property PostCityCode() As String

        'USED TO SET AND RETRIEVE THE POST CITY CODE (STRING)
        Get
            Return strPostCityCode
        End Get

        Set(ByVal Value As String)
            strPostCityCode = Value
        End Set

    End Property

    Public Property PostTownCode() As String

        'USED TO SET AND RETRIEVE THE POST TOWN CODE (STRING)
        Get
            Return strPostTownCode
        End Get

        Set(ByVal Value As String)
            strPostTownCode = Value
        End Set

    End Property

    Public Property BankSummaryDetails() As String

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return strBankSummaryDetails
        End Get

        Set(ByVal Value As String)
            strBankSummaryDetails = Value
        End Set

    End Property

    Public Property Phone1() As String

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return strPhone1
        End Get

        Set(ByVal Value As String)
            strPhone1 = Value
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

        lBankID = 0
        strBankName = ""
        strBranchName = ""
        strPhysicalAddress = ""
        strCountryCode = ""
        strCityCode = ""
        strPostalAddress = ""
        strPostCode = ""
        strPostCountryCode = ""
        strPostCityCode = ""
        strPostTownCode = ""
        strBankSummaryDetails = ""
        strBranchEFTCode = ""
        strPhone1 = ""

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


            If Trim(strBankName) <> "" Or _
                Trim(strBranchName) <> "" Or _
                    Trim(strPhysicalAddress) <> "" Then

                strInsertInto = "INSERT INTO DetailsOfBank (" & _
                    "BankName, BranchName, BranchEFTCode, PhysicalAddress," & _
                        "CountryCode,CityCode, PostalAddress, PostCode, " & _
                            "PostCountryCode, PostCityCode," & _
                                "PostTownCode, BankDetails, Phone1) VALUES "

                strSaveQuery = strInsertInto & _
                            "('" & strBankName & _
                            "', '" & strBranchName & _
                            "', '" & strBranchEFTCode & _
                            "', '" & strPhysicalAddress & _
                            "', '" & strCountryCode & _
                            "', '" & strCityCode & _
                            "', '" & strPostalAddress & _
                            "', '" & strPostCode & _
                            "', '" & strPostCountryCode & _
                            "', '" & strPostCityCode & _
                            "', '" & strPostTownCode & _
                            "', '" & strBankSummaryDetails & _
                            "', '" & strPhone1 & _
                            "')"

                objLogin.connectString = strAccessConnString
                objLogin.ConnectToDatabase()

                bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, strSaveQuery, _
                datSaved)

                objLogin.CloseDb()

                If bSaveSuccess = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - Bank Saved")

                Else

                    MsgBox("'Save Bank' action failed." & _
                        " Make sure all mandatory details are entered", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Bank Failed")

                End If

            Else
                MsgBox("Cannot save. Missing information", _
                        MsgBoxStyle.Exclamation, "iManagement -Missing Information")
            End If

        Catch ex As Exception

        End Try

    End Sub

    Public Function Find(ByVal strQuery As String) As Boolean

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

                    lBankID = myDataRows("BankID")
                    strBankName = myDataRows("BankName").ToString()
                    strBranchName = myDataRows("BranchName").ToString()
                    strBranchEFTCode = myDataRows("BranchEFTCode").ToString()
                    strPhysicalAddress = myDataRows("PhysicalAddress").ToString()
                    strCountryCode = myDataRows("CountryCode").ToString()
                    strCityCode = myDataRows("CityCode").ToString()
                    strPostalAddress = myDataRows("PostalAddress").ToString()
                    strPostCode = myDataRows("PostCode").ToString()
                    strPostCountryCode = myDataRows("PostCountryCode").ToString()
                    strPostCityCode = myDataRows("PostCityCode").ToString()
                    strPostTownCode = myDataRows("PostTownCode").ToString()
                    strBankSummaryDetails = myDataRows("BankDetails").ToString()
                    strPhone1 = myDataRows("Phone1").ToString()


                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Sub Delete(ByVal strDelQuery As String)
        'Deletes the country details of the country with the country code
        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strDeleteQuery = strDelQuery

        If lBankID <> 0 Or _
             Trim(strBankName) <> "" Or _
               Trim(strBranchName) <> "" Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Bank Lookup Details Deleted")
            Else
                MsgBox("'Bank delete' action failed", _
                    MsgBoxStyle.Exclamation, " Bank Deletion failed")
            End If
        Else
            MsgBox("Cannot Delete. Please select an existing bank", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

        End If

    End Sub

    Public Sub Update(ByVal strUpQuery As String)
        'Updates country details of the country with the country code

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If lBankID <> 0 Or _
                    Trim(strBankName) <> "" Or _
                      Trim(strBranchName) <> "" Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, strUpdateQuery, _
            datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement - Bank Lookup Details Updated")
            End If

        End If

    End Sub

   
#End Region



End Class
