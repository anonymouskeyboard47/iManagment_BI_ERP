Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCountries

#Region "PrivateVariables"

    Private strCountryName As String
    Private strCountryCode As String
    Private strCurrencyCode As String
    Private strCurrencyName As String

#End Region

#Region "Properties"
    Public Property CountryName() As String
        'USED TO SET AND RETRIEVE THE COUNTRY NAME STRING VALUE
        Get
            Return strCountryName
        End Get

        Set(ByVal Value As String)
            strCountryName = Value
        End Set
    End Property

    Public Property CountryCode() As String
        'USED TO SET AND RETRIEVE THE COUNTRY CODE STRING VALUE
        Get
            Return strCountryCode
        End Get

        Set(ByVal Value As String)
            strCountryCode = Value
        End Set
    End Property

    Public Property CurrencyName() As String
        'USED TO SET AND RETRIEVE THE CONNECTION STRING VALUE
        Get
            Return strCurrencyName
        End Get

        Set(ByVal Value As String)
            strCurrencyName = Value
        End Set
    End Property

    Public Property CurrencyCode() As String
        'USED TO SET AND RETRIEVE THE CURRENCY CODE STRING VALUE
        Get
            Return strCurrencyCode
        End Get

        Set(ByVal Value As String)
            strCurrencyCode = Value
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

        'Clear country variables
        strCountryName = ""
        strCountryCode = ""
        strCurrencyName = ""
        strCurrencyCode = ""

    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Sub Save()
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try

            If Trim(strCountryName) <> "" Or _
                 Trim(strCountryCode) <> "" Or _
                   Trim(strCurrencyName) <> "" Or _
                       Trim(strCurrencyCode) <> "" Then

                strSaveQuery = "INSERT INTO Country (CountryName, CountryCode," & _
                        "CurrencyCode, CurrencyName) VALUES " & _
                                "('" & strCountryName & "', '" & strCountryCode & _
                                "', '" & strCurrencyCode & _
                                "', '" & strCurrencyName & "')"

                objLogin.connectString = strAccessConnString
                objLogin.ConnectToDatabase()

                bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, strSaveQuery, _
                datSaved)

                objLogin.CloseDb()

                If bSaveSuccess = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - City Saved")
                End If

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

        Try

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
                    strCountryName = myDataRows("CountryName").ToString()
                    strCountryCode = myDataRows("CountryCode").ToString()
                    strCurrencyName = myDataRows("CurrencyName").ToString()
                    strCurrencyCode = myDataRows("CurrencyCode").ToString()
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
        'Deletes the country details of the country with the country code
        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try
        
        strDeleteQuery = strDelQuery

        If Trim(strCountryName) <> "" Or _
             Trim(strCountryCode) <> "" Or _
               Trim(strCurrencyName) <> "" Or _
                   Trim(strCurrencyCode) <> "" Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Country Lookup Details Deleted")
            End If

        End If

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Update(ByVal strUpQuery As String)
        'Updates country details of the country with the country code

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try

        strUpdateQuery = strUpQuery

        If Trim(strCountryName) <> "" Or _
             Trim(strCountryCode) <> "" Or _
               Trim(strCurrencyName) <> "" Or _
                   Trim(strCurrencyCode) <> "" Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, strUpdateQuery, _
            datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement - Country Lookup Details Updated")
            End If

        End If

        Catch ex As Exception

        End Try

    End Sub


#End Region

#Region "ErrorProcedures"



#End Region


End Class
