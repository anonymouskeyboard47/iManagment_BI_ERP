
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMOrganizationSID

#Region "PrivateVariables"

    Private lOrganizationID As Long
    Private strSystemSID As String

#End Region


#Region "Properties"

    Public Property OrganizationID() As Long

        Get
            Return lOrganizationID
        End Get

        Set(ByVal Value As Long)
            lOrganizationID = Value
        End Set

    End Property

    Public Property SystemSID() As String

        Get
            Return Trim(strSystemSID)
        End Get

        Set(ByVal Value As String)
            strSystemSID = Value
        End Set

    End Property

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "DatabaseProcedures"

    Public Shadows Sub Save(ByVal DisplayConfirm As Boolean, _
            ByVal DisplayError As Boolean, _
                ByVal DisplaySuccess As Boolean, _
                    ByVal DisplayFailure As Boolean)

        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String
            Dim strNextSID As String


            If lOrganizationID = 0 Or _
            Trim(strSystemSID) = "" _
                                Then
                If DisplayError = False Then
                    MsgBox("Please provide an existing:" & _
                    Chr(10) & "1. Organization" & _
                    Chr(10) & "2. User or Group" _
                                , MsgBoxStyle.Exclamation, _
                                "iManagement - invalid or incomplete information")
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub
            End If


            'Check if there is an existing series with this name
            If Find("SELECT * FROM OrganizationSID WHERE  " & _
            "OrganizationID = " & lOrganizationID & _
            " AND SystemSID = '" & strSystemSID & "'" _
                            , False) = True Then

                If DisplayFailure = False Then
                    If MsgBox("The Group or User has been linked to the organization already exists.", _
                            MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                    MsgBoxResult.Yes Then

                    End If
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub
            End If

            If DisplayConfirm = True Then
                If MsgBox("Are you sure you want to add this new User or Group to the organization?" _
                , MsgBoxStyle.YesNo, "iManagment - Add new record?") _
                = MsgBoxResult.No Then

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Sub
                End If
            End If


            strInsertInto = "INSERT INTO OrganizationSID (" & _
                "OrganizationID," & _
                "SystemSID" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lOrganizationID & _
                    ",'" & Trim(strSystemSID) & _
                            "')"

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccess = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - User\Group to Organization Details Saved")
                End If

            Else
                If DisplayFailure = True Then
                    MsgBox("'Save User-Group' action failed." & _
                        " Make sure all mandatory details are entered", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - User\Group to organization Addition Failed")
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

    Public Shadows Function Find(ByVal strQuery As String, _
            ByVal bReturnValues As Boolean) As Boolean

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
                    Exit Function

                End If


                For Each myDataRows In myDataTables.Rows
                    If bReturnValues = True Then

                        lOrganizationID = myDataRows("OrganizationID")
                        strSystemSID = myDataRows("SystemSID").ToString

                    End If

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Sub Delete()

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin


        If lOrganizationID = 0 Or Trim(strSystemSID) = "" Then
            MsgBox("Cannot Delete due to missing information. Please provide an existing" & _
            " User\Group and Organization details." _
                        , MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")
            objLogin = Nothing
            datDelete = Nothing

            Exit Sub
        End If


        If MsgBox("Are you sure you want to delete this User\Group to organization detaisls?" _
            , MsgBoxStyle.YesNo, _
                "iManagement - Delete the Group's details?") = MsgBoxResult.No Then

            objLogin = Nothing
            datDelete = Nothing

            Exit Sub
        End If


        strDeleteQuery = "DELETE * FROM OrganizationSID WHERE OrganizationID = " & _
                    lOrganizationID & " AND SystemSID = " & strSystemSID

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
        datDelete)

        objLogin.CloseDb()

        If bDelSuccess = True Then
            MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                "iManagement - User\Group to organization Details Deleted")
        Else
            MsgBox("'Delete Group' action failed", _
                MsgBoxStyle.Exclamation, " User\Group to organization Deletion failed")
        End If

    End Sub

#End Region

End Class

