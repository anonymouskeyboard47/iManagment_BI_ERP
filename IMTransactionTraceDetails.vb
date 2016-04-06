Option Explicit On 
'Option Strict On
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMTransactionTraceDetails
    Inherits IMTransaction

#Region "PrivateVariables"

    Private decGrossTotal As Decimal
    Private decTotalDiscountAmount As Decimal
    Private decTotalVATAmount As Decimal
    Private lCostCentreID As Long
    Private decNetBalance As Decimal
    Private lCustomerNo As Long
    Private strPINNo As Long
    Private decTotalCustomerPrepayment As Decimal
    Private decCustomerPrepaymentBalance As Decimal
    Private strTransactionType As String
    Private dbTotalInvoiceDiscountPercentage As Double
    Private strVATRegistrationNumber As String
    Private decTotalOverallDiscountAmount As Decimal

#End Region


#Region "Properties"

    Public Property VATRegistrationNumber() As String

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return strVATRegistrationNumber
        End Get

        Set(ByVal Value As String)
            strVATRegistrationNumber = Value
        End Set

    End Property

    Public Property CustomerPrepaymentBalance() As Decimal

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return decCustomerPrepaymentBalance
        End Get

        Set(ByVal Value As Decimal)
            decCustomerPrepaymentBalance = Value
        End Set

    End Property

    Public Property TotalCustomerPrepayment() As Decimal

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return decTotalCustomerPrepayment
        End Get

        Set(ByVal Value As Decimal)
            decTotalCustomerPrepayment = Value
        End Set

    End Property

    Public Property CustomerNo() As Long

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return lCustomerNo
        End Get

        Set(ByVal Value As Long)
            lCustomerNo = Value
        End Set

    End Property

    Public Property PINNo() As String

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return strPINNo
        End Get

        Set(ByVal Value As String)
            strPINNo = Value
        End Set

    End Property

    'Gross total of all transactions
    Public ReadOnly Property GrossTotal() As Decimal

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return decGrossTotal
        End Get

    End Property

    'Discount for that partiular invoice
    Public ReadOnly Property TotalOverallDiscountPercentage() As Double

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return dbTotalInvoiceDiscountPercentage
        End Get

        'Set(ByVal Value As Double)
        '    dbTotalInvoiceDiscountPercentage = Value
        'End Set

    End Property

    'Total of TransactionDiscountTotalAmount and this particular invoice amount
    Public ReadOnly Property TotalOverallDiscountAmount() As Decimal

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return decTotalOverallDiscountAmount
        End Get

        'Set(ByVal Value As Decimal)
        '    decTotalOverallDiscountAmount = Value
        'End Set

    End Property

    Public Property TransactionType() As String

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return strTransactionType
        End Get

        Set(ByVal Value As String)
            strTransactionType = Value
        End Set

    End Property

    Public Property CostCentre() As Long

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return lCostCentreID
        End Get

        Set(ByVal Value As Long)
            lCostCentreID = Value
        End Set

    End Property

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region



#Region "DatabaseProcedures"

    Public Function SaveTraceDetails(ByVal bDisplayErrorMessages As Boolean, _
        ByVal bDisplayConfirmation As Boolean, _
            ByVal bDisplayFailure As Boolean, _
                ByVal bDisplaySuccess As Boolean) As Boolean

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String
        Dim i As Long


        Try

            If Save(bDisplayErrorMessages, bDisplayConfirmation, _
                bDisplayFailure, bDisplaySuccess) = False Then

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If



            'If bDisplayConfirmation = True Then
            '    If MsgBox("Do you want to add this new Transaction?", _
            '        MsgBoxStyle.YesNo, _
            '            "iManagement - Add Transaction Details?") _
            '                = MsgBoxResult.No Then

            '        objLogin = Nothing
            '        datSaved = Nothing

            '        Exit Function

            '    End If
            'End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            strInsertInto = "INSERT INTO TransactionTraceDetails (" & _
                "TransactionID," & _
                "GrossTotal," & _
                "TotalDiscountAmount," & _
                "TotalVATAmount," & _
                "CostCentreID," & _
                "NetBalance," & _
                "CustomerNo," & _
                "PINNo," & _
                "TotalCustomerPrepayment," & _
                "CustomerPrepaymentBalance," & _
                "TransactionType," & _
                "TotalInvoiceDiscountPercentage," & _
                "VATRegistrationNumber," & _
                "TotalOverallDiscountAmount," & _
                "TotalAmountPaid" & _
                    ") VALUES "


            strSaveQuery = strInsertInto & _
                    "(" & TransactionID & _
                    "," & decGrossTotal & _
                    "," & ReturnTotalTransactionDiscounts() & _
                    "," & ReturnTotalTransactionVATTotal() & _
                    "," & lCostCentreID & _
                    "," & decNetBalance & _
                    "," & lCustomerNo & _
                    ",'" & strPINNo & _
                    "'," & ReturnCustomerPrepayment() & _
                    "," & decCustomerPrepaymentBalance & _
                    ",'" & strTransactionType & _
                    "'," & dbTotalInvoiceDiscountPercentage & _
                    ",'" & strVATRegistrationNumber & _
                    "'," & decTotalOverallDiscountAmount & _
                    "," & ReturnTotalAmountPaid() & _
                            ")"

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, strSaveQuery, datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If bDisplaySuccess = True Then

                    ReturnError +=" Transaction Saved Successfully."

                End If
            Else

                If bDisplayFailure = True Then

                  ReturnError += "'Save Transaction' action failed." & _
            " Make sure all mandatory details are entered."

                End If
            End If


            objLogin = Nothing
            datSaved = Nothing

            If bSaveSuccess = True Then
                Return True
            End If


        Catch ex As Exception
            If bDisplayErrorMessages = True Then
                returnerror += ex.Message.ToString
            End If
        End Try

    End Function

    Public Function FindTraceDetails(ByVal strQuery As String, _
                        ByVal ReturnStatus As Boolean) As Boolean
        'Query must contain at least rows from Sequence

        Try

            Dim datRetData As DataSet = New DataSet
            Dim bQuerySuccess As Boolean
            Dim myDataTables As DataTable
            Dim myDataColumns As DataColumn
            Dim myDataRows As DataRow
            Dim objLogin As IMLogin = New IMLogin
            Dim i As Long

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

                            decGrossTotal = _
                                    myDataRows("GrossTotal")
                            decTotalDiscountAmount = _
                                    myDataRows("TotalDiscountAmount")
                            decTotalVATAmount = _
                                    myDataRows("TotalVATAmount")
                            lCostCentreID = _
                                    myDataRows("CostCentreID")
                            decNetBalance = _
                                    myDataRows("NetBalance")
                            lCustomerNo = _
                                    myDataRows("CustomerNo")
                            strPINNo = _
                                    myDataRows("PINNo").ToString
                            decTotalCustomerPrepayment = _
                                    myDataRows("TotalCustomerPrepayment")
                            decCustomerPrepaymentBalance = _
                                    myDataRows("CustomerPrepaymentBalance")
                            strTransactionType = _
                                    myDataRows("TransactionType").ToString
                            dbTotalInvoiceDiscountPercentage = _
                                   myDataRows("TotalInvoiceDiscountPercentage")
                            strVATRegistrationNumber = _
                                  myDataRows("VATRegistrationNumber")
                            decTotalOverallDiscountAmount = _
                                  myDataRows("TotalOverallDiscountAmount")

                        Next
                    End If
                Next

                datRetData = Nothing
                objLogin = Nothing

                Return True

            Else

                datRetData = Nothing
                objLogin = Nothing

                Return False

            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString, _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Critical System Error")

        End Try

    End Function

    Public Function DeleteTraceDetails(ByVal strDelQuery As String) As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = strDelQuery

            If TransactionID = 0 Then
                returnerror += "Cannot Delete. Please select an " & _
                    "existing Transaction Detail."

                objLogin = Nothing
                datDelete = Nothing
                Exit Function

            End If


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                ReturnError += "Transaction Details Deleted"

                datDelete = Nothing
                objLogin = Nothing
                Return True
            Else

                MsgBox("'Delete Transaction action failed", _
                    MsgBoxStyle.Exclamation, _
                        "Transaction Deletion failed")


            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Function


#End Region


End Class

