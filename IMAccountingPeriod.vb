Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMAccountingPeriod


#Region "PrivateAccountingPeriodVariables"
    Private lAccountingPeriod As Long
    Private dtStartDate As Date
    Private dtEndDate As String
    Private strPeriodStatus As String
    

#End Region

#Region "Properties"

    Public Property AccountingPeriod() As Long

        Get
            Return lAccountingPeriod
        End Get

        Set(ByVal Value As Long)
            lAccountingPeriod = Value
        End Set

    End Property

    Public Property StartDate() As Date

        'USED TO SET AND RETRIEVE THE BANK NAME (STRING)
        Get
            Return dtStartDate
        End Get

        Set(ByVal Value As Date)
            dtStartDate = Value
        End Set

    End Property

    Public Property EndDate() As Date

        'USED TO SET AND RETRIEVE THE BRANCH NAME (STRING)
        Get
            Return dtEndDate
        End Get

        Set(ByVal Value As Date)
            dtEndDate = Value
        End Set

    End Property

    Public Property PeriodStatus() As String

        'USED TO SET AND RETRIEVE THE BRANCH NAME (STRING)
        Get
            Return strPeriodStatus
        End Get

        Set(ByVal Value As String)
            strPeriodStatus = Value
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

        lAccountingPeriod = 0
        dtStartDate = Now()
        dtEndDate = Now()

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

            'check if periodstatus is provided
            If Trim(strPeriodStatus) = "" Then

                MsgBox("Please provide a period status in order to save or update the record", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

                Exit Sub

            End If

            'Check if another enabled is provided
            If strPeriodStatus = "Enabled" Then
                If Find("SELECT * FROM AccountingPeriod" & _
                    " WHERE PeriodStatus = 'Enabled'") = True Then

                    MsgBox("You have indicated that this period is Enabled and there is another" & _
                    Chr(10) & " active period. You can either:" & _
                        Chr(10) & "1.Select 'Disabled' for this particular entry" & _
                        Chr(10) & "2.Deactivate the Enabled period by selecting 'Disabled'.", _
                        MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, _
                        "iManagement - Cannot add current period due to Period Status duplication")

                    Exit Sub
                End If
            End If

            'Range provided exists
            If Find("SELECT * FROM AccountingPeriod WHERE StarDate => '" & _
                dtStartDate & "' AND EndDate <= '" & _
                    dtEndDate & "'") = True Then

                MsgBox("Cannot save this value since there is a period falling in" & _
                Chr(10) & " between or on the specified period ranges." & _
                 Chr(10) & " Please provide different period ranges.", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Cannot add the period ranges")

                Exit Sub
            End If

            If lAccountingPeriod <> 0 Then
                If Find("SELECT * FROM AccountingPeriod WHERE" & _
                        " AccountingPeriod = " & lAccountingPeriod) = True Then
                    If MsgBox("Do you want to update this accounting period's details?", _
                        MsgBoxStyle.YesNo, "iManagement - Update accounting period's details?") _
                            = MsgBoxResult.Yes Then

                        Update("UPDATE AccountingPeriod SET " & _
                            "StartDate = '" & dtStartDate & "' AND" & _
                            "EndDate = '" & dtEndDate & "' AND" & _
                            "PeriodStatus = '" & strPeriodStatus & "'")

                    End If
                End If

            End If


            strInsertInto = "INSERT INTO AccountingPeriod (" & _
                    "StartDate," & _
                    "EndDate," & _
                    "PeriodStatus," & _
                        ") VALUES "

            strSaveQuery = strInsertInto & _
                        "'" & dtStartDate & _
                        "', '" & dtEndDate & _
                        "', '" & strPeriodStatus & _
                        "')"

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                            strSaveQuery, _
                            datSaved)

            objLogin.CloseDb()


            If bSaveSuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Accounting Period Saved")

            Else

                MsgBox("'Save Accounting Period' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Save Accounting Period Failed")

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
                    datRetData = Nothing
                    objLogin = Nothing

                    Exit Function

                End If


                For Each myDataRows In myDataTables.Rows

                    lAccountingPeriod = myDataRows("AccountingPeriod")
                    dtStartDate = myDataRows("StartDate").ToString()
                    dtEndDate = myDataRows("EndDate").ToString()
                    strPeriodStatus = myDataRows("PeriodStatus").ToString()

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Sub Delete(ByVal strDelQuery As String)

        Try

            'Deletes the country details of the country with the country code
            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            If strDeleteQuery = "" Then
                strDeleteQuery = "DELETE * FROM AccountingPeriod WHERE" & _
                    " AccountingPeriod = " & lAccountingPeriod
            End If

            If lAccountingPeriod <> 0 Then

                If Find("SELECT * FROM AccountingPeriod WHERE" & _
                        " AccountingPeriod = " & AccountingPeriod) = False Then

                    MsgBox("The accounting period provided for deletion is invalid." _
                        , MsgBoxStyle.Exclamation, _
                            "iManagament - invalid or incomplete informaiton")

                    Exit Sub
                End If

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - Accounting Period Details Deleted")
                Else
                    MsgBox("'Accounting Period Delete' action failed", _
                        MsgBoxStyle.Exclamation, " Accounting Period Deletion failed")
                End If

            Else
                MsgBox("Cannot Delete. Please select an existing Accounting Period ID", _
                        MsgBoxStyle.Exclamation, "iManagement -Missing Information")

            End If


        Catch ex As Exception
            MsgBox("Error during deletion. Please contact the Systems Administrator." & _
                Chr(10) & "(" & ex.Message & ")", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Deletion Error")
        End Try

    End Sub

    Public Sub Update(ByVal strUpQuery As String)
        'Updates country details of the country with the country code

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If lAccountingPeriod <> 0 Then

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strUpdateQuery, _
            datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement - Accounting Period Details Updated")
            End If

        End If

    End Sub

#End Region

End Class
