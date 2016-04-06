
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMWorkFlowObject

#Region "PrivateVariables"

    Private lWorkFlowRecordItem As Long
    Private lWorkFlowObjectID As Long
    Private lWorkFlowID As String
    Private lBusinessProcessID As String
    Private bWorkFlowStatus As Double
    Private dtDateCreated As Date
    Private lObjectID2 As Long
    Private lUserID As Long
    Private dtDateCompleted As Date


#End Region


#Region "Properties"

    

    Public Property UserID() As Long

        Get
            Return lUserID
        End Get

        Set(ByVal Value As Long)
            lUserID = Value
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

    Public Property DateCompleted() As Date

        Get
            Return dtDateCompleted
        End Get

        Set(ByVal Value As Date)
            dtDateCompleted = Value
        End Set

    End Property

    Public Property ObjectID2() As Long

        Get
            Return lObjectID2
        End Get

        Set(ByVal Value As Long)
            lObjectID2 = Value
        End Set

    End Property

    Public Property WorkFlowObjectID() As Long

        Get
            Return lWorkFlowObjectID
        End Get

        Set(ByVal Value As Long)
            lWorkFlowObjectID = Value
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

    Public Property WorkFlowStatus() As Boolean

        Get
            Return bWorkFlowStatus
        End Get

        Set(ByVal Value As Boolean)
            bWorkFlowStatus = Value
        End Set

    End Property


#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "DatabaseProcedures"

    Public Function MarkWorkFlowComplete _
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
                        "flow item you want to mark as competed"
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM WorkFlowObject WHERE " & _
            " WorkFlowRecordItem = " & lWorkFlowRecordItem, False) _
            = False Then

                ReturnError += "This work flow item does not exist"
                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            If Find("SELECT * FROM WorkFlowObject WHERE " & _
        " WorkFlowRecordItem = " & lWorkFlowRecordItem & _
        " AND WorkFlowStatus = FALSE", False) _
        = True Then

                ReturnError += "This work flow item has been marked as " & _
                "complete. Therefore you are not able to change " & _
                "it's details"

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            strSaveQuery = "UPDATE WorkFlowObject SET " & _
            " WorkFlowStatus = TRUE AND DateCompleted = Now()"


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

                    ReturnError += "Work Flow Marked Complete successfully."

                End If

                Return True

            Else

                If bDisplayFailure = True Then

                    ReturnError += "'Mark Work Flow Complete' action failed." & _
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

    Public Function ChangeWorkFlowUser _
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
                        "flow item you want to mark as competed"
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM WorkFlowObject WHERE " & _
            " WorkFlowRecordItem = " & lWorkFlowRecordItem, False) _
            = False Then

                ReturnError += "This work flow item does not exist"
                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            If Find("SELECT * FROM WorkFlowObject WHERE " & _
            " WorkFlowRecordItem = " & lWorkFlowRecordItem & _
            " AND WorkFlowStatus = FALSE", False) _
            = True Then

                ReturnError += "This work flow item has been marked as " & _
                "complete. Therefore you are not able to change " & _
                "it's details"

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            strSaveQuery = "UPDATE WorkFlowObject SET " & _
            " WorkFlowStatus = TRUE AND DateCompleted = Now()"




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

                    ReturnError += "Work Flow Marked Complete successfully."

                End If

                Return True

            Else

                If bDisplayFailure = True Then

                    ReturnError += "'Mark Work Flow Complete' action failed." & _
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

            If lWorkFlowID = 0 Or _
                lBusinessProcessID = 0 Or lWorkFlowObjectID = 0 Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please select the business process, " & _
                        "the work flow item itself, and the " & _
                            "purpose of the work flow"
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM WorkFlowObject WHERE " & _
            " WorkFlowRecordItem = " & lWorkFlowRecordItem, False) _
            = True Then

                ReturnError += "This work flow item exists"
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


            strInsertInto = "INSERT INTO WorkFlowObject (" & _
                "WorkFlowObjectID," & _
                "WorkFlowID," & _
                "BusinessProcessID," & _
                "WorkFlowStatus," & _
                "ObjectID2" & _
                "UserID" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lWorkFlowObjectID & _
                    "," & lWorkFlowID & _
                    "," & lBusinessProcessID & _
                    ",FALSE" & _
                    "," & lObjectID2 & _
                    "," & lUserID & _
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

                    ReturnError += "Work Flow Saved Successfully."

                End If

                Return True

            Else

                If DisplayFailure = True Then

                    ReturnError += "'Save Work Flow' action failed." & _
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
                            lWorkFlowObjectID = _
                                myDataRows("WorkFlowObjectID")
                            lWorkFlowID = _
                                myDataRows("WorkFlowID")
                            lBusinessProcessID = _
                                myDataRows("BusinessProcessID")
                            bWorkFlowStatus = _
                                myDataRows("WorkFlowStatus")
                            lObjectID2 = _
                                myDataRows("ObjectID2")
                            lUserID = _
                                myDataRows("UserID")
                            dtDateCompleted = _
                                myDataRows("DateCompleted")
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


            strDeleteQuery = "DELETE * FROM WorkFlowObject WHERE " & _
            "WorkFlowRecordItem = " & lWorkFlowRecordItem

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            datDelete = Nothing
            objLogin = Nothing

            If bDelSuccess = True Then
                ReturnError += "Work Flow Object Details Deleted"
                Return True
            Else

                ReturnError += "'Delete Work Flow Object' action failed"

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

            If lWorkFlowRecordItem <> 0 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery _
                                (strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    ReturnError += "Work Flow Object details updated Successfully"
                End If

            End If

            datUpdated = Nothing
            objLogin = Nothing


        Catch ex As Exception

        End Try

    End Sub

#End Region


End Class
