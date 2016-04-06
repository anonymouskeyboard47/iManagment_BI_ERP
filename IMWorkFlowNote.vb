
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMWorkFlowNote


#Region "PrivateVariables"

    Private lWorkFlowRecordItem As Long
    Private lWorkFlowNoteID As Long
    Private strWorkFlowNote As String
    Private strWorkFlowNoteType As String
    Private dtDateCreated As String
    Private lCreatedBy As Long 'Person who was at the business process at that time
    Private lCheckedBy As Long 'Person who reviewed the notes for that time
    Private dtDateChecked As Date

#End Region


#Region "Properties"

    Public Property WorkFlowRecordItem() As Long

        Get
            Return lWorkFlowRecordItem
        End Get

        Set(ByVal Value As Long)
            lWorkFlowRecordItem = Value
        End Set

    End Property

    Public Property WorkFlowNoteID() As Long

        Get
            Return lWorkFlowNoteID
        End Get

        Set(ByVal Value As Long)
            lWorkFlowNoteID = Value
        End Set

    End Property

    Public Property WorkFlowNote() As String

        Get
            Return strWorkFlowNote
        End Get

        Set(ByVal Value As String)
            strWorkFlowNote = Value
        End Set

    End Property

    Public Property DateCreated() As Date

        Get
            Return dtDateCreated
        End Get

        Set(ByVal Value As Date)
            dtDateCreated = Value
        End Set

    End Property

    Public Property DateChecked() As Date

        Get
            Return dtDateChecked
        End Get

        Set(ByVal Value As Date)
            dtDateChecked = Value
        End Set

    End Property

    Public Property WorkFlowNoteType() As String

        Get
            Return strWorkFlowNoteType
        End Get

        Set(ByVal Value As String)
            strWorkFlowNoteType = Value
        End Set

    End Property

    Public Property CreatedID() As Long

        Get
            Return lCreatedBy
        End Get

        Set(ByVal Value As Long)
            lCreatedBy = Value
        End Set

    End Property

    Public Property CheckedBy() As Long

        Get
            Return lCheckedBy
        End Get

        Set(ByVal Value As Long)
            lCheckedBy = Value
        End Set

    End Property

    
#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "DatabaseProcedures"

    Public Function MarkWorkFlowNoteChecked _
        (ByVal lValWorkFlowItem As Long, _
        ByVal bDisplayErrorMessages As Boolean, _
        ByVal bDisplaySuccess As Boolean, _
        ByVal bDisplayFailure As Boolean) As Boolean

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String


        Try

            If lValWorkFlowItem = 0 Then

                If bDisplayErrorMessages = True Then

                    ReturnError += "Please select the actual work " & _
                        "flow note item you want to check"
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM WorkFlowNote WHERE " & _
            " WorkFlowNoteID = " & lWorkFlowNoteID, False) _
            = False Then

                ReturnError += "This work flow item does not exist"
                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            If Find("SELECT * FROM WorkFlowNote WHERE " & _
        " WorkFlowNoteID = " & lWorkFlowNoteID & _
        " AND CheckedBy <> 0 OR  CheckedBy IS NOT NULL", False) _
        = True Then

                ReturnError += "This work flow note item has been marked as " & _
                "checked. Therefore you are not able to change " & _
                "it's details"

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            strSaveQuery = "UPDATE WorkFlowNote SET " & _
            " CheckedBy = " & lCheckedBy & " AND DateChecked = Now()"


            'If MsgBox(" Are you sure you want to add this Cost Centre?", _
            '            MsgBoxStyle.YesNo, _
            '            "iManagement - Add New Record?") = _
            '            MsgBoxResult.No Then

            '    datSaved = Nothing
            '    objLogin = Nothing
            '    Exit Function
            'End If


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)


            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If bDisplaySuccess = True Then

                    ReturnError += "Work Flow Note marked as checked successfully."

                End If

                Return True

            Else

                If bDisplayFailure = True Then

                    ReturnError += "'Mark Work Flow Note Checked' action failed." & _
            " Make sure all mandatory details are entered."

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If

            objLogin = Nothing
            datSaved = Nothing


        Catch ex As Exception
            If bDisplayErrorMessages = True Then
                ReturnError += ex.Message.ToString

            End If
        End Try


    End Function


    Public Function Save(ByVal DisplayErrorMessages As Boolean, _
        ByVal DisplayConfirmation As Boolean, _
            ByVal DisplayFailure As Boolean, _
                ByVal DisplaySuccess As Boolean) As Boolean

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String


        Try

            If lWorkFlowRecordItem = 0 Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please select the business process " & _
                        "and the work flow item itself"
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM WorkFlowNote WHERE " & _
            " WorkFlowNoteID = " & lWorkFlowNoteID, False) _
            = True Then

                If Find("SELECT * FROM WorkFlowNote WHERE " & _
                 " WorkFlowNoteID = " & lWorkFlowNoteID & _
                 " AND CheckedBy <> 0 OR  CheckedBy IS NOT NULL", False) _
                 = True Then

                    ReturnError += "This work flow note item has been marked as " & _
                    "checked. Therefore you are not able to change " & _
                    "it's details"

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If

                Update("UPDATE WorkFlowNote SET " & _
                "   WorkFlowNote = '" & strWorkFlowNote & _
                "', WorkFlowNoteType = '" & strWorkFlowNoteType & _
                "'")


                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If




            'If MsgBox(" Are you sure you want to add this Cost Centre?", _
            '            MsgBoxStyle.YesNo, _
            '            "iManagement - Add New Record?") = _
            '            MsgBoxResult.No Then

            '    datSaved = Nothing
            '    objLogin = Nothing
            '    Exit Function
            'End If


            strInsertInto = "INSERT INTO WorkFlowNote (" & _
                "WorkFlowRecordItem," & _
                "WorkFlowNote," & _
                "WorkFlowNoteType," & _
                "CreatedBy" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lWorkFlowRecordItem & _
                    "," & strWorkFlowNote & _
                    "," & strWorkFlowNoteType & _
                    "," & lCreatedBy & _
                            ")"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)


            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccess = True Then

                    ReturnError += "Work Flow Note Saved Successfully."

                End If

                Return True

            Else

                If DisplayFailure = True Then

                    ReturnError += "'Save Work Flow Note' action failed." & _
            " Make sure all mandatory details are entered."

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If

            objLogin = Nothing
            datSaved = Nothing


        Catch ex As Exception
            If DisplayErrorMessages = True Then
                ReturnError += ex.Message.ToString

            End If
        End Try

    End Function

    Public Function Find(ByVal strQuery As String, _
                        ByVal ReturnStatus As Boolean) As Boolean
        'Query must contain at least rows from Sequence

        Try

            Dim datRetData As DataSet = New DataSet
            Dim bQuerySuccess As Boolean
            Dim myDataTables As DataTable
            Dim myDataColumns As DataColumn
            Dim myDataRows As DataRow
            Dim objLogin As IMLogin = New IMLogin

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bQuerySuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strQuery, datRetData)

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

                    'Whether to fill properties with values or not
                    If ReturnStatus = True Then

                        For Each myDataRows In myDataTables.Rows

                            lWorkFlowRecordItem = _
                                myDataRows("WorkFlowRecordItem")
                            lWorkFlowNoteID = _
                                myDataRows("WorkFlowNoteID")
                            strWorkFlowNote = _
                                myDataRows("WorkFlowNote")
                            strWorkFlowNoteType = _
                                myDataRows("WorkFlowNoteType")
                            lCreatedBy = _
                                myDataRows("CreatedBy")
                            lCheckedBy = _
                                myDataRows("CheckedBy")
                            dtDateChecked = _
                                myDataRows("DateChecked")
                            dtDateCreated = _
                               myDataRows("DateCreated")

                        Next
                    End If
                Next

                Return True

            End If



        Catch ex As Exception
            ReturnError += ex.Message.ToString

        End Try

    End Function

    Public Function Delete(Optional ByVal bDisplayError As Boolean = False, _
    Optional ByVal bDisplayConfirm As Boolean = False, _
    Optional ByVal bDisplaySuccess As Boolean = False) As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            If lWorkFlowRecordItem = 0 Then
                ReturnError += "Cannot Delete. Please select an " & _
                    "existing Work Flow item"

                datDelete = Nothing
                objLogin = Nothing

                Exit Function
            End If

            'If bDisplayConfirm = True Then
            '    If MsgBox(" Are you sure you want to Delete this Cost Centre?", _
            '                             MsgBoxStyle.YesNo, _
            '                             "iManagement - Delete Record?") = _
            '                              MsgBoxResult.No Then

            '        datDelete = Nothing
            '        objLogin = Nothing
            '        Exit Function
            '    End If
            'End If


            strDeleteQuery = "DELETE * FROM WorkFlowNote WHERE " & _
            "WorkFlowNoteID = " & lWorkFlowNoteID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            datDelete = Nothing
            objLogin = Nothing

            If bDelSuccess = True Then
                ReturnError += "Work Flow Note Details Deleted"
                Return True
            Else

                ReturnError += "'Delete Work Flow Note' action failed"

            End If

        Catch ex As Exception

        End Try

    End Function

    Public Sub Update(ByVal strUpQuery As String)

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If lWorkFlowNoteID <> 0 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery _
                                (strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    ReturnError += "Work Flow Note details updated Successfully"
                End If

            End If

            datUpdated = Nothing
            objLogin = Nothing


        Catch ex As Exception

        End Try

    End Sub

#End Region

End Class
