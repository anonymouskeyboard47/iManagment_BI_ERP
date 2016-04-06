Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections


Public Class IMCity
    Inherits IMCountries

#Region "PrivateVariables"

    Private strCityName As String
    Private lCityCode As Long
    Private bCapitalCity As Boolean

#End Region

#Region "Properties"
    Public Property CityName() As String
        'USED TO SET AND RETRIEVE THE CityName value (String)
        Get
            Return strCityName
        End Get

        Set(ByVal Value As String)
            strCityName = Value
        End Set
    End Property

    Public Property CityCode() As Long
        'USED TO SET AND RETRIEVE THE CityCode value (String)
        Get
            Return lCityCode
        End Get

        Set(ByVal Value As Long)
            lCityCode = Value
        End Set
    End Property

    Public Property CapitalCity() As Boolean
        'USED TO SET AND RETRIEVE THE Capital City Value (Boolean)
        Get
            Return bCapitalCity
        End Get

        Set(ByVal Value As Boolean)
            bCapitalCity = Value
        End Set
    End Property

#End Region

#Region "IntitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    'Clears the derived and owner class variables
    Public Shadows Sub NewRecord()

        'Clear city variables
        strCityName = ""
        lCityCode = 0
        bCapitalCity = False

        'Clear country variables
        CountryName = ""
        CountryCode = ""
        CurrencyName = ""
        CurrencyCode = ""

    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Shadows Sub Save()
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        If Trim(CountryName) <> "" Or _
             Trim(strCityName) <> "" Then

            strSaveQuery = "INSERT INTO City (CountryName, CityName," & _
                    "CapitalCity) VALUES " & _
                            "('" & CountryName & "', '" & CityName & _
                            "', " & bCapitalCity & ")"

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()
            bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - City's Lookup Saved")
            End If

        End If

    End Sub

    Public Shadows Function Find(ByVal strQuery As String) As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow

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
                    strCityName = myDataRows("CityName").ToString()
                    lCityCode = myDataRows("CityCode")
                    bCapitalCity = myDataRows("CapitalCity")
                    CountryName = myDataRows("CountryName").ToString()
                Next

            Next

            Return True
        Else
            Return False
        End If

    End Function

    Public Shadows Function Delete(ByVal strQuery As String) As Boolean
        'Deletes the country details of the country with the country code
        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strDeleteQuery = strQuery

        If Trim(CountryName) <> "" Or _
             Trim(strCityName) <> "" Or _
               lCityCode <> 0 Then

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - City's Lookup Details Deleted")
                Return True
            End If

        End If
    End Function

    Public Shadows Function Update(ByVal strQuery As String) As Boolean
        'Updates country details of the country with the country code

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strQuery

        If Trim(CountryName) <> "" Or _
             Trim(strCityName) <> "" Or _
               lCityCode <> 0 Then

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, strUpdateQuery, _
            datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement - City's Lookup Details Updated")
            End If

        End If
    End Function

#End Region

End Class
