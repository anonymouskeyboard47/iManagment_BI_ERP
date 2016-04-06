
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMStockManager

#Region "PrivateVariables"

    Private lProductID As Long
    Private lStockManagementID As Long
    Private lReorderLevel As Long
    Private lReorderQuantity As Long
    Private bStockManagementStatus As Boolean

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

    Public Property StockManagementID() As Long

        Get
            Return lStockManagementID
        End Get

        Set(ByVal Value As Long)
            lStockManagementID = Value
        End Set

    End Property

    Public Property ReorderLevel() As Long

        Get
            Return lReorderLevel
        End Get

        Set(ByVal Value As Long)
            lReorderLevel = Value
        End Set

    End Property

    Public Property ReorderQuantity() As Long

        Get
            Return lReorderQuantity
        End Get

        Set(ByVal Value As Long)
            lReorderQuantity = Value
        End Set

    End Property

    Public Property StockManagementStatus() As Boolean

        Get
            Return bStockManagementStatus
        End Get

        Set(ByVal Value As Boolean)
            bStockManagementStatus = Value
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

            If lProductID = 0 Or lReorderLevel = 0 Or lReorderQuantity = 0 Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide the following details in" & _
                " order to save a Stock Management detail:" & _
                Chr(10) & "1. Existing Product" & _
                Chr(10) & "2. Reorder Level" & _
                Chr(10) & "3. Reorder Quantity" _
                , MsgBoxStyle.Critical, _
                "iManagement - Save Action Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM ProductMaster WHERE ProductID = " & _
                lProductID, False) = False Then

                MsgBox("Please select an existing Product ID", _
                MsgBoxStyle.Critical, _
                "iManagement - invalid or incomplete information")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO StockDeductions (" & _
                "ProductID," & _
                "StockManagementID," & _
                "ReorderLevel," & _
                "ReorderQuantity," & _
                "StockManagementStatus" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lProductID & _
                    "," & lStockManagementID & _
                    "," & lReorderLevel & _
                    "," & lReorderQuantity & _
                    "," & bStockManagementStatus & _
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

                Exit Function

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

                            lProductID = _
                                myDataRows("ProductID")
                            lStockManagementID = _
                                myDataRows("StockManagementID")
                            lReorderLevel = _
                                myDataRows("ReorderLevel")
                            lReorderQuantity = _
                                myDataRows("ReorderQuantity")
                            bStockManagementStatus = _
                                myDataRows("StockManagementStatus")

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

            If lProductID <> 0 _
                                Then

                objLogin.connectString = strAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
                datDelete)

               

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Stock Management Details Deleted", MsgBoxStyle.Information, _
                        "iManagement - Record Deleted Successfully")

                Else

                    MsgBox("'Delete Stock Management' action failed", _
                        MsgBoxStyle.Exclamation, "Deletion failed")


                End If

            Else

                MsgBox("Cannot Delete. Please select an existing Stock Management Detail", _
                        MsgBoxStyle.Exclamation, "iManagement -Missing Information")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub Update(ByVal strUpQuery As String, _
    ByVal DisplayErrorMessages As Boolean, _
        ByVal DisplayConfirmation As Boolean, _
            ByVal DisplayFailure As Boolean, _
                ByVal DisplaySuccess As Boolean)

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If lProductID <> 0 _
                            Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    If DisplaySuccess = True Then
                        MsgBox("Record Updated Successfully", _
                            MsgBoxStyle.Information, _
                                "iManagement -  Stock Management Updated")
                    End If

                End If

            End If

        Catch ex As Exception

        End Try
    End Sub

#End Region

End Class
