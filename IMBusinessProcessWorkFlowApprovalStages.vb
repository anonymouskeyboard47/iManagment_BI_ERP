Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections


Public Class IMBusinessProcessWorkFlowApprovalStages


#Region "PrivateVariables"

    Private lBusinessProcessID As Long
    Private lWorkFlowID As Long
    Private lAllowedApprovalOfficerUserID As Long

#End Region

#Region "Properties"

    Public Property BusinessProcessID() As Long

        Get
            Return lBusinessProcessID
        End Get

        Set(ByVal Value As Long)
            lBusinessProcessID = Value
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

    Public Property AllowedApprovalOfficerUserID() As Long

        Get
            Return lAllowedApprovalOfficerUserID
        End Get

        Set(ByVal Value As Long)
            lAllowedApprovalOfficerUserID = Value
        End Set

    End Property


#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "DatabaseProcedures"

    'Save informaiton
    Public Function Save(ByVal DisplayErrorMessages As Boolean, _
            ByVal DisplaySuccessMessages As Boolean, _
                ByVal DisplayFailureMessages As Boolean) As Boolean

        'Saves a new base organization
        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String
            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String


            If Trim(strOrganizationName) = "" Then

                MsgBox("Please open an existing company.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Select an existing company")
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If lBusinessProcessID = 0 Or _
            lAllowedApprovalOfficerUserID = 0 Or _
            lWorkFlowID = 0 _
                Then

                MsgBox("You must provide an existing Business Process, an associated" & _
                Chr(10) & " Work Flow, and an existing User's  details." _
                                , MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if there is an existing series with this name
            If Find("SELECT * FROM BusinessProcessWorkFlowApprovalStages WHERE BusinessProcessID = " _
                & lBusinessProcessID & " AND WorkFlowID = " & lWorkFlowID, False) = True Then

                'confirm Update
                If MsgBox("This Business Process Name and Work Flow approval details already exists." & _
                    Chr(10) & "Do you want to update the  details?", _
                            MsgBoxStyle.YesNo, "iManagement - Record Exists. Update") = _
                                    MsgBoxResult.No Then

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If


                Update("UPDATE BusinessProcessWorkFlowApprovalStages SET " & _
                    "AllowedApprovalOfficerUserID = " & lAllowedApprovalOfficerUserID & _
                        " WHERE WorkFlowID = " _
                                & lWorkFlowID & " AND BusinessProcessID = " & _
                                        lBusinessProcessID)

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            'Confirm Addition
            If MsgBox("Do you want to add this new Work Flow Approval details?" _
                , MsgBoxStyle.YesNo, _
                    "iManagement - Add the Business Process Work Flow Approval?") _
                        = MsgBoxResult.No Then

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If



            strInsertInto = "INSERT INTO BusinessProcessWorkFlowApprovalStages (" & _
                    "BusinessProcessID," & _
                    "WorkFlowID," & _
                    "AllowedApprovalOfficerUserID" & _
                            ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lBusinessProcessID & _
                        "," & lWorkFlowID & _
                        "," & lAllowedApprovalOfficerUserID & _
                            ")"


            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - Work Flow Approval Details Saved")

                End If
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save Business Process details' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Work Flow Approval Details Failed")
                End If
            End If

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception
            If DisplayErrorMessages = True Then
                MsgBox(ex.Source, MsgBoxStyle.Critical, _
                    "iManagement - Database or system error")
            End If

        End Try

    End Function

    'Find Informaiton
    Public Function Find(ByVal strQuery As String, _
        ByVal bReturnValues As Boolean) As Boolean

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
                    objLogin = Nothing
                    datRetData = Nothing


                    Exit Function

                End If



                For Each myDataRows In myDataTables.Rows

                    If bReturnValues = True Then

                        lBusinessProcessID = _
                                myDataRows("BusinessProcessID")
                        lWorkFlowID = _
                                myDataRows("WorkFlowID").ToString
                        lAllowedApprovalOfficerUserID = _
                                myDataRows("AllowedApprovalOfficerUserID").ToString

                    End If
                Next

            Next

            objLogin = Nothing
            datRetData = Nothing

            Return True
        Else

            objLogin = Nothing
            datRetData = Nothing
            Return False

        End If

        objLogin = Nothing
        datRetData = Nothing

    End Function

    'Delete data
    Public Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try

            If lBusinessProcessID = 0 Or lWorkFlowID = 0 _
                Then

                MsgBox("Cannot Delete. Please select an existing Work Flow Approval Details.", _
                                    MsgBoxStyle.Exclamation, _
                                    "iManagement - invalid or incomplete information")

                objLogin = Nothing
                datDelete = Nothing

                Exit Sub

            End If


            strDeleteQuery = "DELETE * FROM BusinessProcessMaster WHERE BusinessProcessID = " & _
                lBusinessProcessID


            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully.", MsgBoxStyle.Information, _
                    "iManagement - Work Flow Approval Details Deleted")
            Else
                MsgBox("'Delete Work Flow Approval' action failed.", _
                    MsgBoxStyle.Exclamation, "Work Flow Approval Details Deletion failed")
            End If




            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully.", MsgBoxStyle.Information, _
                    "iManagement - Work Flow Approval Details Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub


#End Region

End Class
