Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMOrgActivities

#Region "PrivateBankVariables"
    Private lActivityID As Long
    Private strActivityTitle As String

#End Region

#Region "Properties"

    Public Property ActivityID() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lActivityID
        End Get

        Set(ByVal Value As Long)
            lActivityID = Value
        End Set

    End Property

    Public Property ActivityTitle() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strActivityTitle
        End Get

        Set(ByVal Value As String)
            strActivityTitle = Value
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
        lActivityID = 0
        strActivityTitle = ""

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

        If Trim(strActivityTitle) <> "" Then

            strInsertInto = "INSERT INTO CompanyActivity (" & _
                "ActivityTitle) VALUES "

            strSaveQuery = strInsertInto & _
                        "('" & strActivityTitle & _
                        "')"

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Organization activity Saved")

            Else

                MsgBox("'Save Organization Activity' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Save Activity Failed")

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

                        lActivityID = myDataRows("ActivityID")
                        strActivityTitle = myDataRows("ActivityTitle").ToString()

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

        If lActivityID <> 0 Or _
                      Trim(strActivityTitle) <> "" Then

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Activity Lookup Details Deleted")
            Else
                MsgBox("'Bank delete' action failed", _
                    MsgBoxStyle.Exclamation, " Activity Deletion failed")
            End If
        Else
            MsgBox("Cannot Delete. Please select an existing Activity", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

        End If

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If lActivityID <> 0 Or _
                 Trim(strActivityTitle) <> "" Then

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
