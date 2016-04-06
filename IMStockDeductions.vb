Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMStockDeductions
    Inherits IMProductStockAmount

#Region "PrivateVariables"

    Private lProductID As Long
    Private lDeductionID As Long
    Private dtDeductionDate As Date
    Private lQuantityRemoved As Long
    Private strPurposeOfRemoval As String
    Private dbRemovalUnitPrice As Decimal
    
#End Region

#Region "Properties"


    Public Property DeductionID() As Long

        Get
            Return lDeductionID
        End Get

        Set(ByVal Value As Long)
            lDeductionID = Value
        End Set

    End Property

    Public Property DeductionDate() As Date

        Get
            Return dtDeductionDate
        End Get

        Set(ByVal Value As Date)
            dtDeductionDate = Value
        End Set

    End Property

    Public Property QuantityRemoved() As Long

        Get
            Return lQuantityRemoved
        End Get

        Set(ByVal Value As Long)
            lQuantityRemoved = Value
        End Set

    End Property

    Public Property PurposeOfRemoval() As String

        Get
            Return strPurposeOfRemoval
        End Get

        Set(ByVal Value As String)
            strPurposeOfRemoval = Value
        End Set

    End Property

    Public Property RemovalUnitPrice() As Decimal

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return dbRemovalUnitPrice
        End Get

        Set(ByVal Value As Decimal)
            dbRemovalUnitPrice = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Function SaveDeduction(ByVal DisplayErrorMessages As Boolean, _
        ByVal DisplayConfirmation As Boolean, _
            ByVal DisplayFailure As Boolean, _
                ByVal DisplaySuccess As Boolean) As Boolean

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        Try

            If ProductID = 0 Or lQuantityRemoved = 0 Or _
                    Trim(strPurposeOfRemoval) = "" Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide the following details in" & _
                " order to save a Stock Deduction:" & _
                Chr(10) & "1. Existing Product" & _
                Chr(10) & "2. Amount Removed" & _
                Chr(10) & "3. Purpose Of Removal" _
                , MsgBoxStyle.Critical, _
            "iManagement - Save Action Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            Else


                If FindDeduction("SELECT * FROM ProductMatser WHERE ProductID = " & _
                lProductID, False) = False Then

                    MsgBox("Please select an existing product id", _
                     MsgBoxStyle.Critical, _
                        "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If


                strInsertInto = "INSERT INTO StockDeductions (" & _
                    "ProductID," & _
                    "DeductionDate," & _
                    "QuantityRemoved," & _
                    "RelationshipTitle," & _
                    "PurposeOfRemoval," & _
                    "RemovalUnitPrice" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "(" & ProductID & _
                        ",#" & dtDeductionDate & _
                        "#," & lQuantityRemoved & _
                        ",'" & strPurposeOfRemoval & _
                        "'," & dbRemovalUnitPrice & _
                                ")"

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bSaveSuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strSaveQuery, datSaved)

                objLogin.CloseDb()

                If bSaveSuccess = True Then
                    If DisplaySuccess = True Then
                        MsgBox("Stock Deduction Saved Successfully", _
                        MsgBoxStyle.Information, _
                        "iManagement - Record Saved Successfully")

                    End If

                Else

                    If DisplayFailure = True Then
                        MsgBox("'Save Stock Deduction' action failed." & _
                        " Make sure all mandatory details are entered", _
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

    Public Function FindDeduction(ByVal strQuery As String, _
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

                            ProductID = _
                                myDataRows("ProductID")
                            lDeductionID = _
                                myDataRows("DeductionID")
                            dtDeductionDate = _
                                myDataRows("DeductionDate")
                            lQuantityRemoved = _
                                myDataRows("QuantityRemoved")
                            strPurposeOfRemoval = _
                                myDataRows("PurposeOfRemoval")
                            dbRemovalUnitPrice = _
                                myDataRows("RemovalUnitPrice")

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

    Public Function DeleteDeduction() As Boolean

        If ProductID <> 0 Then

            MsgBox("Cannot Delete. Please select an existing Stock Deduction Detail", _
                    MsgBoxStyle.Exclamation, _
                    "iManagement -Missing Information")

            Exit Function
        End If

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = "DELETE * FROM StockDeductions WHERE " & _
                " lDeductionID = " & lDeductionID

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Stock Deduction Details Deleted", _
                MsgBoxStyle.Information, _
                    "iManagement - Record Deleted Successfully")

            Else

                MsgBox("'Delete Stock Deduction' action failed", _
                    MsgBoxStyle.Exclamation, "Deletion failed")

            End If

            datDelete = Nothing
            objLogin = Nothing


        Catch ex As Exception

        End Try
    End Function

    Public Sub UpdateDeduction(ByVal strUpQuery As String)

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
                                    strUpdateQuery, datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    MsgBox("Record Updated Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement -  Stock Deduction Updated")
                End If

            End If

        Catch ex As Exception

        End Try
    End Sub



#End Region

End Class
