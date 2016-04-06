Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMMethodOfPayment

#Region "PrivateVariables"

    Private lMethodOfPaymentID As Long
    Private strPaymentMethodTitle As String

#End Region

#Region "Properties"

    Public Property MethodOfPaymentID() As Long

        Get
            Return lMethodOfPaymentID
        End Get

        Set(ByVal Value As Long)
            lMethodOfPaymentID = Value
        End Set

    End Property

    Public Property PaymentMethodTitle() As String

        Get
            Return strPaymentMethodTitle
        End Get

        Set(ByVal Value As String)
            strPaymentMethodTitle = Value
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



            If Trim(strPaymentMethodTitle) = "" Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide a Payment Method Title" _
                , MsgBoxStyle.Critical, _
            "iManagement - Save Action Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            Else

                If Find("SELECT * FROM MethodOfPayment WHERE PaymentMethodTitle = '" & _
                Trim(strPaymentMethodTitle) & "'", False) = True Then

                    If MsgBox("The Payment Method Title Details already exists." & _
                        Chr(10) & "Do you want to update the details?", _
                                MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                        MsgBoxResult.Yes Then

                        Update("UPDATE MethodOfPayment SET " & _
                                    "PaymentMethodTitle = " & Trim(strPaymentMethodTitle) & _
                                    " WHERE strPaymentMethodTitle = '" & _
                                    strPaymentMethodTitle & "' OR MethodOfPaymentID = " _
                                    & lMethodOfPaymentID)

                    End If

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If


                strInsertInto = "INSERT INTO MethodOfPayment (" & _
                    "PaymentMethodTitle" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "('" & strPaymentMethodTitle & _
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

                        MsgBox("Payment Method Title Saved Successfully.", _
                            MsgBoxStyle.Information, _
                                "iManagement - Record Saved Successfully")

                    End If
                Else

                    If DisplayFailure = True Then

                        MsgBox("'Save Payment Method Title' action failed." & _
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

                            lMethodOfPaymentID = _
                                myDataRows("MethodOfPaymentID")
                            strPaymentMethodTitle = _
                                myDataRows("PaymentMethodTitle")

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

            If Trim(strPaymentMethodTitle) <> "" _
                                                Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

               

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Payment Method Title Details Deleted", MsgBoxStyle.Information, _
                        "iManagement - Record Deleted Successfully")

                Else

                    MsgBox("'Delete Payment Method Title action failed", _
                        MsgBoxStyle.Exclamation, "Payment Method Deletion failed")

                    objLogin.RollbackTheTrans()

                End If

            Else

                MsgBox("Cannot Delete. Please select an existing Payment Method Title Detail", _
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

            If Trim(strPaymentMethodTitle) <> "" Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    MsgBox("Record Updated Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement - Payment Method Title Updated")
                End If

            End If

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

            Return strTextFieldData
            datFillData.Dispose()

        Catch ex As Exception

        End Try

    End Function


#End Region


End Class
