Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMGroups
    Inherits IMSIDMaster

#Region "PrivateVariables"

    Private lGroupID As Long
    Private strGroupTitle As String
    Private strGroupDescription As String
    Private bGroupStatus As Boolean
    Private dtCreationRegistered As Date

#End Region


#Region "Properties"

    Public Property CreationRegistered() As Date

        Get
            Return dtCreationRegistered
        End Get

        Set(ByVal Value As Date)
            dtCreationRegistered = Value
        End Set

    End Property

    Public Property GroupID() As Long

        Get
            Return lGroupID
        End Get

        Set(ByVal Value As Long)
            lGroupID = Value
        End Set

    End Property

    Public Property GroupTitle() As String

        Get
            Return Trim(strGroupTitle)
        End Get

        Set(ByVal Value As String)
            strGroupTitle = Value
        End Set

    End Property

    Public Property GroupDescription() As String

        Get
            Return Trim(strGroupDescription)
        End Get

        Set(ByVal Value As String)
            strGroupDescription = Value
        End Set

    End Property

    Public Property GroupStatus() As Boolean

        Get
            Return bGroupStatus
        End Get

        Set(ByVal Value As Boolean)
            bGroupStatus = Value
        End Set

    End Property

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "DatabaseProcedures"

    Public Shadows Sub GroupSave()
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim objOrgSID As IMOrganizationSID
        Dim strInsertInto As String
        Dim strNextSID As String

        If Trim(strGroupTitle) = "" Or _
            Trim(strGroupDescription) = "" _
                            Then

            ReturnError += "Please provide an existing" & _
            Chr(10) & "1. Group's Title" & _
            Chr(10) & "2. Appropriate Group Description"

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If

        Dim objOvSetup As IMOverallSetup = New IMOverallSetup

        If objOvSetup.Find("SELECT * FROM CompanyMaster WHERE OrganizationName = '" & _
        strOrganizationName & "'", True, False, True) = False Then

            ReturnError += "Please open an existing organization in " & _
             "order to register Business Processes"

            objLogin = Nothing
            datSaved = Nothing

            objOvSetup = Nothing
            Exit Sub

        End If


        'Check if there is an existing series with this name
        If Find("SELECT * FROM Groups " & _
        " INNER JOIN OrganizationSID ON " & _
        " OrganizationSID.SystemSID = Groups.SystemSID " & _
        "WHERE  GroupTitle = '" & strGroupTitle & _
        "' AND OrganizationID = " & objOvSetup.OrganizationID, _
        False) = True Then

            'If MsgBox("The Group Name already exists." & _
            'Chr(10) & "Do you want to update the details?", _
            '        MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
            '                MsgBoxResult.No Then

            '    datSaved = Nothing
            '    objLogin = Nothing
            '    Exit Sub

            'End If

            Update("Update SIDMaster SET " & _
                " SIDStatus = " & bGroupStatus & _
                " WHERE SystemSID  = " & SystemSID, False)

            GroupUpdate("UPDATE Groups SET " & _
                        "GroupDescription = '" & Trim(strGroupDescription) & _
                        "', SystemSID = '" & SystemSID & _
                            "' WHERE  GroupTitle = '" _
                                & Trim(strGroupTitle) & "'")

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If


        'If MsgBox("Are you sure you want to this new Group?" _
        ', MsgBoxStyle.YesNo, "iManagment - Add new user record?") _
        '= MsgBoxResult.No Then

        '    objLogin = Nothing
        '    datSaved = Nothing

        '    Exit Sub
        'End If


        Type = "Group"
        TypeUserOrGroupID = "Group"
        SIDStatus = bGroupStatus


        If Save(False, False, False, False) = False Then
            ReturnError += "Cannot save Security Identity details. " & _
            "Cannot Save the Group Details."

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If


        If SystemSID = "" Then
            ReturnError += "Cannot save Security Identity details. " & _
                "Cannot Save the Group Details."

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If


        strInsertInto = "INSERT INTO Groups (" & _
            "GroupTitle," & _
            "GroupDescription," & _
            "SystemSID" & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                "(" & Chr(34) & Trim(strGroupTitle) & _
                 Chr(34) & "," & Chr(34) & Trim(strGroupDescription) & _
                Chr(34) & "," & Chr(34) & SystemSID & _
                         Chr(34) & ")"

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
        strSaveQuery, datSaved)

        objLogin.CloseDb()

        If bSaveSuccess = True Then
            returnerror += "Record Saved Successfully"

            objOrgSID = New IMOrganizationSID

            With objOrgSID

                .OrganizationID = objOvSetup.OrganizationID
                .SystemSID = SystemSID

                .Save(False, False, False, False)

                objOvSetup = Nothing

            End With

            objLogin = Nothing
            datSaved = Nothing

            objOrgSID = Nothing

        Else

           returnerror += "'Save Group' action failed." & _
                " Make sure all mandatory details are entered"

        End If


    End Sub

    Public Shadows Function GroupFind(ByVal strQuery As String, _
            ByVal bReturnValues As Boolean, _
                ByVal bUseSystemSID As Boolean) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.ConnectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bQuerySuccess = objLogin.ExecuteQuery(strAccessConnString, _
        strQuery, datRetData)

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
                        strGroupTitle = myDataRows("GroupTitle").ToString()
                        strGroupDescription = myDataRows("GroupDescription").ToString()
                        dtCreationRegistered = myDataRows("CreationDate")

                        If bUseSystemSID = True Then
                            SystemSID = ("Groups.SystemSID")
                        Else
                            SystemSID = ("SystemSID")

                        End If

                        Find("SELECT * FROM SIDMaster WHERE SystemSID = '" _
                                        & SystemSID & "'", True)

                        bGroupStatus = SIDStatus

                    End If

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Sub GroupDelete()

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin


        If lGroupID = 0 Then
            ReturnError += "Cannot Delete due to missing information. " & _
            "Please provide an existing"
            
            objLogin = Nothing
            datDelete = Nothing

            Exit Sub
        End If


        'If MsgBox("Are you sure you want to delete this Group's detaisls?" _
        '    , MsgBoxStyle.YesNo, _
        '        "iManagement - Delete the Group's details?") = MsgBoxResult.No Then

        '    objLogin = Nothing
        '    datDelete = Nothing

        '    Exit Sub
        'End If


        If Delete(False, False, False, False) = False Then
            returnerror += "Security ID Deletion Failed. 'Delete Group' " & _
                "action failed"

            objLogin = Nothing
            datDelete = Nothing

            Exit Sub
        End If

        strDeleteQuery = "DELETE * FROM Groups WHERE GroupID = " & _
                    lGroupID

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

        objLogin.CloseDb()

        If bDelSuccess = True Then
            returnerror += "Group Details Record Deleted Successfully"
        Else
            ReturnError += "'Delete Group' action failed"

        End If

    End Sub

    Public Sub GroupUpdate(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If (lGroupID) <> 0 Then

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            If SystemSID = "" Then
                returnerror+="Cannot Update Security Identity details. " & _
                "Cannot update the Group Details."

                objLogin = Nothing
                datUpdated = Nothing

                Exit Sub
            End If

            'If bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            '    "UPDATE SIDMaster SET SIDStatus  = " & bGroupStatus & _
            '        " WHERE SystemSID = '" & SystemSID & "'", _
            '            datUpdated) = False Then

            '    MsgBox("Cannot Update Security Identity details. Cannot update the Group Details." _
            '                            , MsgBoxStyle.Exclamation, "iManagement - Cannot update the Group's Details")

            '    objLogin = Nothing
            '    datUpdated = Nothing

            '    Exit Sub
            'End If

            bUpdateSuccess = False

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
               returnerror += "Customer's Bank Account record updated " & _
               "successfully"
            End If

        End If

    End Sub

#End Region


End Class
