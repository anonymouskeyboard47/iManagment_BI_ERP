Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMSIDToWorkFlow

#Region "PrivateVariables"

    Private strSystemSID As String
    Private lWorkFlowID As Long
    Private lBusinessProcessID As Long
    Private bAllocationStatus As Boolean
    Private lOrganizationID As Long
    Private dtDateRegistered As Date

#End Region

#Region "Properties"

    Public Property SystemSID() As String

        Get
            Return strSystemSID
        End Get

        Set(ByVal Value As String)
            strSystemSID = Value
        End Set

    End Property

    Public Property WorkFlowID() As Long

        Get
            Return lWorkFlowID
        End Get

        Set(ByVal Value As Long)
            lWorkFlowID = Value
        End Set

    End Property

    Public Property BusinessProcessID() As Long

        Get
            Return lBusinessProcessID
        End Get

        Set(ByVal Value As Long)
            lBusinessProcessID = Value
        End Set

    End Property

    Public Property AllocationStatus() As Boolean

        Get
            Return bAllocationStatus
        End Get

        Set(ByVal Value As Boolean)
            bAllocationStatus = Value
        End Set

    End Property

    Public Property OrganizationID() As Long

        Get
            Return lOrganizationID
        End Get

        Set(ByVal Value As Long)
            lOrganizationID = Value
        End Set

    End Property

    Public Property DateRegistered() As Date

        Get
            Return dtDateRegistered
        End Get

        Set(ByVal Value As Date)
            dtDateRegistered = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Function Save _
        (ByVal DisplayConfirm As Boolean, _
            ByVal DisplayError As Boolean, _
                ByVal DisplaySuccess As Boolean, _
                    ByVal DisplayFailure As Boolean) As Boolean
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        Dim objOvSetup As IMOverallSetup = New IMOverallSetup

        If objOvSetup.Find("SELECT * FROM CompanyMaster WHERE OrganizationName = '" & _
        strOrganizationName & "'", True, False, True) = False Then

            MsgBox("Please open an existing organization in order to register Business Processes", _
            MsgBoxStyle.Exclamation, "iManagement - Please Open an existing company")

            objOvSetup = Nothing
            Exit Function

        End If


        If Trim(strSystemSID) = "" Or _
            lWorkFlowID = 0 Or _
            lBusinessProcessID = 0 _
                            Then

            MsgBox("Please provide an existing" & _
            Chr(10) & "1. Group" & _
            Chr(10) & "2. Business Process" & _
            Chr(10) & "3. Business Process Work Flow" _
                        , MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

            objLogin = Nothing
            datSaved = Nothing

            Exit Function

        End If


        'Check if there is an existing series with this name
        If Find("SELECT * FROM SIDToWorkFlow WHERE" & _
        " WorkFlowID = " & lWorkFlowID & _
        " AND BusinessProcessID = " & lBusinessProcessID & _
        " AND OrganizationID = " & objOvSetup.OrganizationID & _
        " AND SystemSID = '" & Trim(strSystemSID) & "'" _
            , False) = True Then


            If DisplayConfirm = True Then
                If MsgBox("The (Group to Function) details already exist." & _
                    MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                            MsgBoxResult.Yes Then


                End If

            End If


            objLogin = Nothing
            datSaved = Nothing

            Exit Function
        End If


        If DisplayConfirm = True Then
            If MsgBox("Are you sure you want to link the Group Function to the" & _
            Chr(10) & " Organization and Business Process Work Flow?" _
            , MsgBoxStyle.YesNo, "iManagment - Add new user record?") _
            = MsgBoxResult.No Then
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

        End If


        strInsertInto = "INSERT INTO SIDToWorkFlow (" & _
            "SystemSID," & _
            "WorkFlowID," & _
            "BusinessProcessID," & _
            "AllocationStatus," & _
            "OrganizationID" & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                "('" & Trim(strSystemSID) & _
                "'," & lWorkFlowID & _
                "," & lBusinessProcessID & _
                "," & bAllocationStatus & _
                "," & objOvSetup.OrganizationID & _
                        ")"

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
        strSaveQuery, _
        datSaved)

        objLogin.CloseDb()
        objOvSetup = Nothing


        If bSaveSuccess = True Then
            If DisplaySuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Group Function Details Saved")
            End If

            Return True

        Else

            If DisplayFailure = False Then
                MsgBox("'Save Group Function' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Group Function Addition Failed")
            End If

        End If


    End Function

    Public Function Find(ByVal strQuery As String, _
            ByVal bReturnValues As Boolean) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.ConnectString = strOrgAccessConnString
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
                    If bReturnValues = True Then

                        strSystemSID = myDataRows("SystemSID").ToString()
                        lWorkFlowID = myDataRows("WorkFlowID")
                        lBusinessProcessID = myDataRows("BusinessProcessID")
                        bAllocationStatus = myDataRows("AllocationStatus")
                        lOrganizationID = myDataRows("OrganizationID")
                        dtDateRegistered = myDataRows("DateRegistered")


                    End If

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Function Delete(ByVal DisplayError As Boolean, _
        ByVal DisplayConfirm As Boolean, _
            ByVal DisplayFailure As Boolean, _
                ByVal DisplaySuccess As Boolean) As Boolean

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Dim objOvSetup As IMOverallSetup = New IMOverallSetup

        If objOvSetup.Find("SELECT * FROM CompanyMaster WHERE OrganizationName = '" & _
        strOrganizationName & "'", True, False, True) = False Then

            MsgBox("Please open an existing organization in order to register Business Processes", _
            MsgBoxStyle.Exclamation, "iManagement - Please Open an existing company")

            objOvSetup = Nothing
            Exit Function

        End If

        If Trim(strSystemSID) = "" Or lWorkFlowID = 0 Or lBusinessProcessID = 0 Then

            If DisplayError = True Then

                MsgBox("Cannot Delete due to missing information. Please provide an existing" & _
                " Security ID Details, Work Flow Details, Business Process, and Organization with the details." _
                            , MsgBoxStyle.Exclamation, _
                            "iManagement - invalid or incomplete information")
            End If

            objLogin = Nothing
            datDelete = Nothing

            Exit Function

        End If


        If DisplayConfirm = True Then
            If MsgBox("Are you sure you want to delete this " & _
            "(Group to Function) details?" _
            , MsgBoxStyle.YesNo, _
            "iManagement - Delete the user's details?") = MsgBoxResult.No Then

                objLogin = Nothing
                datDelete = Nothing

                Exit Function

            End If
        End If


        strDeleteQuery = "DELETE * FROM SIDToWorkFlow WHERE " & _
                    " SystemSID = '" & strSystemSID & _
                    "' AND WorkFlowID = " & lWorkFlowID & _
                    " AND BusinessProcessID = " & BusinessProcessID & _
                    " AND OrganizationID = " & objOvSetup.OrganizationID

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
        datDelete)

        objLogin.CloseDb()
        objOvSetup = Nothing

        If bDelSuccess = True Then
            If DisplaySuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - (Group to Function) Details Deleted")
            End If

            Return True

        Else

            If DisplayFailure = True Then
                MsgBox("'Delete (Group to Function)' action failed", _
                    MsgBoxStyle.Exclamation, " (Group to Function) Deletion failed")
            End If
        End If

    End Function


#End Region

End Class
