Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMDocumentMaster

#Region "PrivateVariables"

    Private lDocumentID As Long
    Private lDocumentProducerUserID As Long
    Private lDocumentTypeID As Long
    Private dtDateProduced As Date
    Private lDocumentStatusID As Long
    Private strDocumentSerialNumber As String
    Private strForInputOutput As String

#End Region

#Region "Properties"

    Public Property DocumentID() As Long

        Get
            Return lDocumentID
        End Get

        Set(ByVal Value As Long)
            lDocumentID = Value
        End Set

    End Property

    Public Property DocumentProducerUserID() As Long

        Get
            Return lDocumentProducerUserID
        End Get

        Set(ByVal Value As Long)
            lDocumentProducerUserID = Value
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

    Public Property DateProduced() As Date

        Get
            Return dtDateProduced
        End Get

        Set(ByVal Value As Date)
            dtDateProduced = Value
        End Set

    End Property

    Public Property DocumentStatusID() As Long

        Get
            Return lDocumentStatusID
        End Get

        Set(ByVal Value As Long)
            lDocumentStatusID = Value
        End Set

    End Property

    Public Property DocumentSerialNumber() As String

        Get
            Return strDocumentSerialNumber
        End Get

        Set(ByVal Value As String)
            strDocumentSerialNumber = Value
        End Set

    End Property

    Public Property ForInputOutput() As String

        Get
            Return strForInputOutput
        End Get

        Set(ByVal Value As String)
            strForInputOutput = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function CalculateNextDocSerialNo() As String
        Try

            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String
            Dim strProposedDocNo As String

            Dim objLogin As IMLogin = New IMLogin

            With objLogin

                MyMaxValue = .FillArray(strOrgAccessConnString, _
                            "SELECT COUNT(*) AS TotalRecords FROM" & _
                                " DocumentID WHERE DateProduced = Now()", "", "")

            End With

            objLogin = Nothing


            If Not MyMaxValue Is Nothing Then
                For Each strItem In MyMaxValue
                    If Not strItem Is Nothing Then

                        MaxValue = CLng(Val(strItem))


                    End If
                Next
            End If

            MaxValue = MaxValue + 1

            strProposedDocNo = "Doc" & Now.Day.ToString _
                & Now.Month.ToString & _
                    Now.Year.ToString & _
                            MaxValue.ToString

            Return strProposedDocNo

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, _
                "iManagement - System Error")
        End Try

    End Function

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

           
            If lDocumentProducerUserID = 0 Or _
                lDocumentTypeID = 0 Or _
                    strForInputOutput = 0 Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide the following details in" & _
                " order to save a Document in iManagement Document Manager" & _
                Chr(10) & "1. User ID of the person producing the document." & _
                Chr(10) & "2. The Document Type." & _
                Chr(10) & "3. Indicate whether the document is meant for input (scanned)" & _
        Chr(10) & "     or if the document is meant for output (e.g. A voucher from the system)" & _
                MsgBoxStyle.Critical, _
            "iManagement - Save Action Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            Else

                If Find("SELECT * FROM DocumentMaster WHERE DocumentSerialNumber = '" & _
                        Trim(strDocumentSerialNumber) & "' OR DocumentID = " _
                            & lDocumentID, False) = True Then

                    If MsgBox("The Document Details already exists." & _
                        Chr(10) & "Do you want to update the details?", _
                                MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                        MsgBoxResult.Yes Then

                        Update("UPDATE DocumentMaster SET " & _
                                    " AND DocumentProducerUserID = '" _
                                            & lDocumentProducerUserID & _
                                    "' AND DocumentTypeID = " & lDocumentTypeID & _
                                    " AND DocumentStatusID = " & lDocumentStatusID & _
                                    " AND ForInputOutput = " & strForInputOutput & _
                                    " WHERE DocumentSerialNumber = '" & _
                                        Trim(strDocumentSerialNumber) & _
                                            "' OR DocumentID = " & lDocumentID)

                    End If

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If


                strInsertInto = "INSERT INTO DocumentMaster (" & _
                    "lDocumentProducerUserID," & _
                    "lDocumentTypeID," & _
                    "dtDateProduced," & _
                    "lDocumentStatusID," & _
                    "strDocumentSerialNumber," & _
                    "strForInputOutput" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "(" & lDocumentProducerUserID & _
                        "," & lDocumentTypeID & _
                        ",'" & Now() & _
                        "'," & lDocumentStatusID & _
                        ",'" & Trim(CalculateNextDocSerialNo()) & _
                        "','" & strForInputOutput & _
                                "')"

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bSaveSuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, _
                strSaveQuery, _
                datSaved)


                objLogin.CloseDb()

                If bSaveSuccess = True Then
                    If DisplaySuccess = True Then

                        MsgBox("Document Saved Successfully.", _
                            MsgBoxStyle.Information, _
                                "iManagement - Record Saved Successfully")

                    End If
                Else

                    If DisplayFailure = True Then

                        MsgBox("'Save Document' action failed." & _
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

                            lDocumentID = _
                                myDataRows("DocumentID")
                            lDocumentProducerUserID = _
                                myDataRows("DocumentProducerUserID").ToString
                            lDocumentTypeID = _
                                myDataRows("DocumentTypeID")
                            dtDateProduced = _
                                myDataRows("DateProduced")
                            lDocumentStatusID = _
                                myDataRows("DocumentStatusID")
                            strDocumentSerialNumber = _
                                myDataRows("DocumentSerialNumber").ToString
                            lDocumentID = _
                                myDataRows("DocumentID")
                            strForInputOutput = _
                                myDataRows("ForInputOutput")
                           
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

            If lDocumentID <> 0 Or Trim(strDocumentSerialNumber) <> "" _
                                                            Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

               

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Document Details Deleted", MsgBoxStyle.Information, _
                        "iManagement - Record Deleted Successfully")

                Else

                    MsgBox("'Delete Document' action failed", _
                        MsgBoxStyle.Exclamation, "Document Deletion failed")

                    objLogin.RollbackTheTrans()

                End If

            Else

                MsgBox("Cannot Delete. Please select an existing Document Detail", _
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

            If lDocumentID <> 0 Or Trim(strDocumentSerialNumber) <> "" _
                            Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    MsgBox("Record Updated Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement - Document Updated")
                End If

            End If

        Catch ex As Exception

        End Try
    End Sub

   

#End Region


End Class
