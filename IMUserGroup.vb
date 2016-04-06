Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMUserGroup

#Region "PrivateVariables"

    Private lGroupID As Long
    Private lUserID As Long

#End Region


#Region "Properties"

    Public Property GroupID() As Long

        Get
            Return lGroupID
        End Get

        Set(ByVal Value As Long)
            lGroupID = Value
        End Set

    End Property

    Public Property UserID() As Long

        Get
            Return luserid
        End Get

        Set(ByVal Value As Long)
            lUserID = Value
        End Set

    End Property

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "DatabaseProcedures"

    Public Shadows Sub Save()
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String
        Dim strNextSID As String

        If lGroupID = 0 Or _
        lUserID = 0 _
                            Then

            MsgBox("Please provide an existing" & _
            Chr(10) & "1. Group" & _
            Chr(10) & "2. User" _
                        , MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub

        End If


        'Check if there is an existing series with this name
        If Find("SELECT * FROM UserGroup WHERE  GroupID = " _
                    & lGroupID & " AND UserID = " & lUserID, False) = True Then

            If MsgBox("The User-Group already exists.", _
                    MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                            MsgBoxResult.Yes Then

            End If

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If


        If MsgBox("Are you sure you want to this new User-Group?" _
        , MsgBoxStyle.YesNo, "iManagment - Add new record?") _
        = MsgBoxResult.No Then

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If


        strInsertInto = "INSERT INTO UserGroup (" & _
            "GroupID," & _
            "UserID" & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                "(" & lGroupID & _
                "," & lUserID & _
                        ")"

        objLogin.ConnectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
        strSaveQuery, _
        datSaved)

        objLogin.CloseDb()

        If bSaveSuccess = True Then
            MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
            "iManagement - User-Group Details Saved")

        Else

            MsgBox("'Save User-Group' action failed." & _
                " Make sure all mandatory details are entered", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - User-Group's Addition Failed")

        End If


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

                        lGroupID = myDataRows("GroupID")
                        lUserID = myDataRows("UserID")
                       
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


        If lGroupID = 0 Or lUserID = 0 Then
            MsgBox("Cannot Delete due to missing information. Please provide an existing" & _
            " Group's and User's Details." _
                        , MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")
            objLogin = Nothing
            datDelete = Nothing

            Exit Sub
        End If


        If MsgBox("Are you sure you want to delete this User-Group's details?" _
            , MsgBoxStyle.YesNo, _
                "iManagement - Delete the Group's details?") = MsgBoxResult.No Then

            objLogin = Nothing
            datDelete = Nothing

            Exit Sub
        End If


        strDeleteQuery = "DELETE * FROM UserGroup WHERE GroupID = " & _
                    lGroupID & " AND UserID = " & lUserID

        objLogin.ConnectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strDeleteQuery, _
                datDelete)

        objLogin.CloseDb()

        If bDelSuccess = True Then
            MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                "iManagement - User-Group's Details Deleted")
        Else
            MsgBox("'Delete Group' action failed", _
                MsgBoxStyle.Exclamation, " User-Group's Deletion failed")
        End If

    End Sub

    'Public Sub Update(ByVal strUpQuery As String)

    '    Dim strUpdateQuery As String
    '    Dim datUpdated As DataSet = New DataSet
    '    Dim bUpdateSuccess As Boolean
    '    Dim objLogin As IMLogin = New IMLogin

    '    strUpdateQuery = strUpQuery

    '    If lUserID <> 0 Or lGroupID <> 0 Then

    '        objLogin.ConnectString = strAccessConnString
    '        objLogin.ConnectToDatabase()

    '        If lUserID = 0 Or lGroupID = 0 Then
    '            MsgBox("Cannot Update Security Identity details. Cannot update the User-Group's Details." _
    '                    , MsgBoxStyle.Exclamation, "iManagement - Cannot update the Group's Details")

    '            objLogin = Nothing
    '            datUpdated = Nothing

    '            Exit Sub
    '        End If


    '        bUpdateSuccess = False

    '        bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
    '                            strUpdateQuery, _
    '                                    datUpdated)

    '        objLogin.CloseDb()

    '        If bUpdateSuccess = True Then
    '            MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
    '                "iManagement -  Customer's Bank Account Details Updated")
    '        End If

    '    End If

    'End Sub

    'Public Function FillControl(ByVal strFillConnString As String, _
    '            ByVal strTSQL As String, ByVal strValueField As String, _
    '                ByVal strTextField As String) As String()

    '    Dim datFillData As DataSet
    '    Dim bReturnedSuccess As Boolean
    '    Dim myDataTables As DataTable
    '    Dim myDataColumns As DataColumn
    '    Dim myDataRows As DataRow
    '    Dim strTextFieldData() As String
    '    Dim i As Integer
    '    Dim objLogin As IMLogin = New IMLogin

    '    Try

    '        datFillData = New DataSet

    '        objLogin.connectString = strAccessConnString
    '        objLogin.ConnectToDatabase()

    '        'The db is okay now get the recordset
    '        bReturnedSuccess = objLogin.ExecuteQuery(strAccessConnString, _
    '            strTSQL, datFillData)

    '        objLogin.CloseDb()

    '        If datFillData Is Nothing Then
    '            Exit Function
    '        End If

    '        For Each myDataTables In datFillData.Tables

    '            'Check if there is any data. If not exit
    '            If myDataTables.Rows.Count = 0 Then
    '                'Return an empty array
    '                ReDim strTextFieldData(1)
    '                strTextFieldData(0) = ""
    '                Return strTextFieldData

    '                Exit Function
    '            Else
    '                'Resize the array
    '                ReDim strTextFieldData(myDataTables.Rows.Count)

    '            End If

    '            i = 0
    '            For Each myDataRows In myDataTables.Rows
    '                strTextFieldData(i) = myDataRows(0).ToString()
    '                i = i + 1
    '            Next

    '        Next

    '        Return strTextFieldData
    '        datFillData.Dispose()

    '    Catch ex As Exception

    '    End Try

    'End Function


#End Region

End Class
