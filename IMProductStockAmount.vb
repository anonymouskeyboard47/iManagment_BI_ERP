Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMProductStockAmount

#Region "PrivateVariables"

    Private lProductID As Long
    Private dbAmountInStock As Double

#End Region


#Region "Properties"

    Public Property ProductID() As Long

        Get
            Return lProductID
        End Get

        Set(ByVal Value As Long)
            lProductID = Value
        End Set

    End Property

    Public Property AmountInStock() As Decimal

        Get
            Return dbAmountInStock
        End Get

        Set(ByVal Value As Decimal)
            dbAmountInStock = Value
        End Set

    End Property

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "DatabaseProcedures"

    Public Function SaveStockAmount(ByVal DisplayErrorMessages As Boolean, _
        ByVal DisplayConfirmation As Boolean, _
            ByVal DisplayFailure As Boolean, _
                ByVal DisplaySuccess As Boolean) As Boolean
        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        Try

            If lProductID = 0 Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide the following details in" & _
                " order to save a Stock Amount detail:" & _
                Chr(10) & "1. Existing Product", MsgBoxStyle.Critical, _
                "iManagement - Save Action Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If FindStockAmount("SELECT * FROM ProductMaster WHERE ProductID = " & _
            lProductID, False) = False Then

                If DisplayErrorMessages = True Then
                    MsgBox("Please select an existing product ID", _
                     MsgBoxStyle.Critical, _
                        "iManagement - invalid or incomplete information")
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If FindStockAmount("SELECT * FROM ProductStockAmount WHERE ProductID = " & _
                    lProductID, False) = True Then

                If DisplayErrorMessages = True Then
                    If MsgBox("Do you want to update the product stock amount?.", _
                        MsgBoxStyle.YesNo, _
                        "iManagement - Record Exists. Update Record?") _
                        = MsgBoxResult.No Then

                        datSaved = Nothing
                        objLogin = Nothing
                        Exit Function

                    End If
                End If

                UpdateStockAmount("UPDATE ProductStockAmount SET " & _
                " AmountInStock = " & dbAmountInStock & _
                " WHERE ProductID = " & lProductID, False, False, False, False)

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO ProductStockAmount(" & _
                "ProductID," & _
                "AmountInStock" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lProductID & _
                    "," & dbAmountInStock & _
                            ")"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, strSaveQuery, datSaved)


            objLogin.CloseDb()


            If bSaveSuccess = True Then
                If DisplaySuccess = True Then
                    MsgBox("Stock Management Saved Successfully.", _
                    MsgBoxStyle.Information, _
                    "iManagement - Record Saved Successfully")

                End If

            Else

                If DisplayFailure = True Then
                    MsgBox("'Save Stock Management' action failed." & _
                    " Make sure all mandatory details are entered", _
                    MsgBoxStyle.Exclamation, _
                    "iManagement -  Addition Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

            End If

            objLogin = Nothing
            datSaved = Nothing

            If bSaveSuccess = True Then
                Return True
            End If


        Catch ex As Exception
            If DisplayErrorMessages = True Then
                MsgBox(ex.Message.ToString, _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Critical System Error")
            End If

        End Try

    End Function

    Public Function FindStockAmount(ByVal strQuery As String, _
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

                            lProductID = _
                                myDataRows("ProductID")
                            dbAmountInStock = _
                                myDataRows("AmountInStock")

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

    Public Function DeleteStockAmount() As Boolean

        Try

            If lProductID = 0 Then
                MsgBox("Cannot Delete. Please select an existing " & _
                " Stock Amount Detail", _
                        MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete Information")
                Exit Function
            End If

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = "DELETE * FROM ProductStockAmount " & _
                " WHERE ProductID = " & lProductID


            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, strDeleteQuery, datDelete)


            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Stock Amount Details Deleted", _
                MsgBoxStyle.Information, _
                "iManagement - Record Deleted Successfully")

            Else

                MsgBox("'Delete Stock Amount' action failed", _
                    MsgBoxStyle.Exclamation, "Deletion failed")

            End If

            datDelete = Nothing
            objLogin = Nothing

            If bDelSuccess = True Then
                Return True
            End If


        Catch ex As Exception

        End Try
    End Function

    Public Function UpdateStockAmount(ByVal strUpQuery As String, _
    ByVal DisplayErrorMessages As Boolean, _
        ByVal DisplayConfirmation As Boolean, _
            ByVal DisplayFailure As Boolean, _
                ByVal DisplaySuccess As Boolean) As Boolean

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If lProductID <> 0 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    If DisplaySuccess = True Then
                        MsgBox("Record Updated Successfully.", _
                            MsgBoxStyle.Information, _
                                "iManagement - Stock Amount Updated")
                    End If

                End If

            End If

        Catch ex As Exception

        End Try
    End Function


#End Region


End Class
