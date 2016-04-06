Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMMeetingAgendas

#Region "PrivateVariables"

    Private lMeetingID As Long
    Private lMeetingAgendaID As Long
    Private strAgendaTitle As String
    Private strAgendaDescription As String

#End Region

#Region "Properties"

    Public Property MeetingID() As Long

        Get
            Return lMeetingID
        End Get

        Set(ByVal Value As Long)
            lMeetingID = Value
        End Set

    End Property

    Public Property MeetingAgendaID() As Long

        Get
            Return lMeetingAgendaID
        End Get

        Set(ByVal Value As Long)
            lMeetingAgendaID = Value
        End Set

    End Property

    Public Property AgendaTitle() As String

        Get
            Return strAgendaTitle
        End Get

        Set(ByVal Value As String)
            strAgendaTitle = Value
        End Set

    End Property

    Public Property AgendaDescription() As String

        Get
            Return strAgendaDescription
        End Get

        Set(ByVal Value As String)
            strAgendaDescription = Value
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

            If lMeetingID = 0 Or Trim(strAgendaTitle) = "" Or _
                Trim(strAgendaDescription) = "" Then

                MsgBox("You must provide an available Meeting, an " & _
                    Chr(10) & "Agenda, and the Description for the Agenda." _
                                , MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if there is an existing series with this name
            If Find("SELECT * FROM MeetingAgendas WHERE strAgendaTitle = '" & _
                Trim(strAgendaTitle) & "'", False) = True Then

                If MsgBox("The Meeting Agendas Details already exists." & _
                               Chr(10) & "Do you want to update the details?", _
                                       MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                               MsgBoxResult.Yes Then

                    Update("UPDATE MeetingAgendas SET " & _
                        "AgendaTitle = '" & Trim(strAgendaTitle) & _
                                "' AND AgendaDescription = '" & Trim(strAgendaDescription) & _
                                    "' WHERE  AgendaTitle = '" _
                                        & Trim(strAgendaTitle) & "' AND lMeetingID = " & lMeetingID)

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

           
            strInsertInto = "INSERT INTO MeetingAgendas (" & _
                "MeetingID," & _
                "strAgendaTitle," & _
                "strAgendaDescription" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lMeetingID & _
                    ",'" & strAgendaTitle & _
                    "','" & strAgendaDescription & _
                    "')"


            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - New Meeting Agendas Details Saved")

                End If
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save New Meeting Agendas Details' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Meeting Agendas Details Failed")
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


                If bReturnValues = True Then
                    For Each myDataRows In myDataTables.Rows
                      
                        lMeetingID = _
                                myDataRows("MeetingID")
                        lMeetingAgendaID = _
                                myDataRows("lMeetingAgendaID")
                        strAgendaTitle = _
                                myDataRows("AgendaTitle").ToString
                        strAgendaDescription = _
                                myDataRows("AgendaDescription").ToString

                    Next
                End If
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


            strDeleteQuery = strDelQuery

            If lMeetingAgendaID = 0 _
                Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - Meeting Agendas Details Deleted")
                Else
                    MsgBox("'Meeting Agendas Details delete' action failed", _
                            MsgBoxStyle.Exclamation, "Meeting Agendas Details Deletion failed")
                End If
            Else

                MsgBox("Cannot Delete. Please select an existing Meeting Agendas.", _
                            MsgBoxStyle.Exclamation, "iManagement - invalid or incomplete information")

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
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement - Meeting Agendas Details Details Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub

    Public Function FillControl(ByVal strFillConnString As String, _
                ByVal strTSQL As String, ByVal strValueField As String, _
                    ByVal strTextField As String) As String()

        Dim datFillData As DataSet
        Dim bReturnedSuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim strTextFieldData() As String
        Dim i As Integer
        Dim objLogin As IMLogin = New IMLogin

        Try

            datFillData = New DataSet

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            'The db is okay now get the recordset
            bReturnedSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strTSQL, datFillData)

            objLogin.CloseDb()

            If datFillData Is Nothing Then
                Exit Function
            End If

            For Each myDataTables In datFillData.Tables

                'Check if there is any data. If not exit
                If myDataTables.Rows.Count = 0 Then
                    'Return an empty array
                    ReDim strTextFieldData(1)
                    strTextFieldData(0) = ""
                    Return strTextFieldData

                    Exit Function
                Else
                    'Resize the array
                    ReDim strTextFieldData(myDataTables.Rows.Count)

                End If

                i = 0
                For Each myDataRows In myDataTables.Rows
                    strTextFieldData(i) = myDataRows(0).ToString()
                    i = i + 1
                Next

            Next


            objLogin = Nothing
            datFillData = Nothing

            Return strTextFieldData
            datFillData.Dispose()

        Catch ex As Exception

        End Try

    End Function

#End Region

End Class
