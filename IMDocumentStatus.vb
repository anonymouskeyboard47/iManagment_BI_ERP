Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMDocumentStatus

#Region "PrivateVariables"

    Private lDocumentStatusID As Long
    Private strDocStatusDescription As String
    Private lDocumentTypeID As Long

#End Region

#Region "Properties"

    Public Property DocumentStatusID() As Long

        Get
            Return lDocumentStatusID
        End Get

        Set(ByVal Value As Long)
            lDocumentStatusID = Value
        End Set

    End Property

    Public Property DocStatusDescription() As String

        Get
            Return strDocStatusDescription
        End Get

        Set(ByVal Value As String)
            strDocStatusDescription = Value
        End Set

    End Property

    Public Property DocumentTypeID() As Long

        Get
            Return lDocumentTypeID
        End Get

        Set(ByVal Value As Long)
            lDocumentTypeID = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "DatabaseProcedures"

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

            If Trim(strDocStatusDescription) = "" Or lDocumentTypeID = 0 Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide a  Document Type and an associated Document Status." _
                , MsgBoxStyle.Critical, _
            "iManagement - Save Action Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            Else

                If Find("SELECT * FROM DocumentStatus WHERE DocStatusDescription = '" & _
                Trim(strDocStatusDescription) & "'", False) = True Then

                    If MsgBox("The Document Status Details already exists." & _
                        Chr(10) & "Do you want to update the details?", _
                                MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                        MsgBoxResult.Yes Then

                        Update("UPDATE DocumentStatus SET " & _
                                    "DocStatusDescription = '" & Trim(strDocStatusDescription) & _
                                    "' AND DocumentTypeID = " & lDocumentTypeID & _
                                    " WHERE (DocStatusDescription = '" & _
                                    strDocStatusDescription & _
                                    "' AND DocumentTypeID = " & lDocumentTypeID & _
                                        ") OR DocumentStatusID = " & lDocumentStatusID)

                    End If

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If


                strInsertInto = "INSERT INTO DocumentStatus (" & _
                    "DocStatusDescription," & _
                    "DocumentTypeID" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "('" & Trim(strDocStatusDescription) & _
                        "'," & lDocumentTypeID & _
                                ")"

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bSaveSuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, _
                strSaveQuery, _
                datSaved)


                objLogin.CloseDb()

                If bSaveSuccess = True Then
                    If DisplaySuccess = True Then

                        MsgBox("Document Status Saved Successfully.", _
                            MsgBoxStyle.Information, _
                                "iManagement - Record Saved Successfully")

                    End If
                Else

                    If DisplayFailure = True Then

                        MsgBox("'Save Document Status' action failed." & _
                " Make sure all mandatory details are entered.", _
                MsgBoxStyle.Exclamation, _
                "iManagement -  Addition Failed")

                    End If

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            objLogin = Nothing
            datSaved = Nothing

            Return True

        Catch ex As Exception
            If DisplayErrorMessages = True Then
                MsgBox(ex.Message.ToString, _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Critical System Error")
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

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bQuerySuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strQuery, _
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

                    'Whether to fill properties with values or not
                    If ReturnStatus = True Then

                        For Each myDataRows In myDataTables.Rows

                            lDocumentStatusID = _
                                myDataRows("lDocumentStatusID")
                            strDocStatusDescription = _
                                myDataRows("DocStatusDescription")
                            lDocumentTypeID = _
                                myDataRows("lDocumentTypeID")

                        Next
                    End If
                Next

                Return True

            Else
                Return False

            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString, _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Critical System Error")

        End Try

    End Function

    Public Sub Delete(ByVal strDelQuery As String)

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = strDelQuery

            If Trim(strDocStatusDescription) <> "" _
                                                Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

               

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Document Status Details Deleted", MsgBoxStyle.Information, _
                        "iManagement - Record Deleted Successfully")

                Else

                    MsgBox("'Delete Document Status action failed", _
                        MsgBoxStyle.Exclamation, "Document Status Deletion failed")

                    objLogin.RollbackTheTrans()

                End If

            Else

                MsgBox("Cannot Delete. Please select an existing Document Status Detail", _
                        MsgBoxStyle.Exclamation, "iManagement - Missing Information")

                objLogin.RollbackTheTrans()

            End If
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

            If Trim(strDocStatusDescription) <> "" Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    MsgBox("Record Updated Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement - Document Status Updated")
                End If

            End If

        Catch ex As Exception

        End Try
    End Sub

  

#End Region


End Class
