Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMOrgTypes

#Region "PrivateOrgTypeVariables"
    Private lCompanyTypeID As Long
    Private strCompanyTypeTitle As String
#End Region

#Region "Properties"

    Public Property CompanyTypeID() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lCompanyTypeID
        End Get

        Set(ByVal Value As Long)
            lCompanyTypeID = Value
        End Set

    End Property

    Public Property CompanyTypeTitle() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strCompanyTypeTitle
        End Get

        Set(ByVal Value As String)
            strCompanyTypeTitle = Value
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
        lCompanyTypeID = 0
        strCompanyTypeTitle = ""

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

        If Trim(strCompanyTypeTitle) <> "" Then

            strInsertInto = "INSERT INTO CompanyTypes (" & _
                "CompanyTypeTitle) VALUES "

            strSaveQuery = strInsertInto & _
                        "('" & strCompanyTypeTitle & _
                        "')"

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Organization Type Saved")

            Else

                MsgBox("'Save Organization Type' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Save Organization Type Failed")

            End If

        End If

    End Sub

    Public Function Find(ByVal strQuery As String, _
            ByVal bReturnDetails As Boolean) As Boolean

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
                    datRetData = Nothing
                    objLogin = Nothing
                    Return False
                    Exit Function

                End If

                If bReturnDetails = True Then

                    For Each myDataRows In myDataTables.Rows

                        lCompanyTypeID = myDataRows("CompanyTypeID")
                        strCompanyTypeTitle = myDataRows("CompanyTypeTitle").ToString()

                    Next

                End If

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strDeleteQuery = strDelQuery

        If lCompanyTypeID <> 0 Or _
                      Trim(strCompanyTypeTitle) <> "" Then

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Organization Type Lookup Details Deleted")
            Else
                MsgBox("'Organization Type delete' action failed", _
                    MsgBoxStyle.Exclamation, " Activity Deletion failed")
            End If
        Else
            MsgBox("Cannot Delete. Please select an existing Organization Type", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

        End If

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If lCompanyTypeID <> 0 Or _
                       Trim(strCompanyTypeTitle) <> "" Then

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Lookup Details Updated")
            End If

        End If

    End Sub



#End Region


End Class
