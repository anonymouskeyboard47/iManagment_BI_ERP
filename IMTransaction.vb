Option Explicit On 
'Option Strict On
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMTransaction

#Region "PrivateVariables"

    Private lTransactionRunningNo As Long
    Private lTransactionID As Long
    Private strTransactionSerialNo() As String
    Private dtTransactionDate() As Date
    Private dtTransactionBookDate() As Date
    Private lCOAAccountNr() As Long
    Private strPayee() As String
    Private lMethodOfPaymentID() As Long
    Private decCurrencyUsed() As Decimal
    Private strItemPaidFor() As String
    Private lDocumentID() As Long
    Private lUserID() As Long
    Private lSequenceGroupID() As Long
    Private strTransactionDescription() As String
    Private lApprovingOfficerUserID() As Long
    Private strTransactionInvoiceNo() As String
    Private dtDateCreated() As Date
    Private strTransactionVATCode() As String
    Private decCustomerPrepayment() As Decimal
    Private decTransactionGrossTotalAmount() As Decimal
    Private dbTransactionDiscountPercentage() As Double
    Private decTransactionDiscount() As Decimal
    Private dbTransactionVATPercentage() As Double
    Private decTransactionVATAmount() As Decimal
    Private decAmountPaid() As Decimal
    Private decTransactionNetAmount() As Decimal
    Private strTransactionProductName() As String
    Private strTransactionProductTextCode() As String
    Private strShippingAgentEmployerName() As String
    Private dtShippingDate() As Date
    Private bTransactionIsDebit() As Boolean
    Private bTransactionUseProduct() As Boolean

#End Region


#Region "Properties"

    'Total amount of the Transaction before deductions
    Public Property TransactionUseProduct() As Boolean()

        Get
            Return bTransactionUseProduct
        End Get

        Set(ByVal Value() As Boolean)
            bTransactionUseProduct = Value
        End Set

    End Property


    'Total amount of the Transaction before deductions
    Public Property TransactionIsDebitCredit() As Boolean()

        Get
            Return bTransactionIsDebit
        End Get

        Set(ByVal Value() As Boolean)
            bTransactionIsDebit = Value
        End Set

    End Property

    'Total amount of the Transaction before deductions
    Public Property TransactionShippingDate() As Date()

        Get
            Return dtShippingDate
        End Get

        Set(ByVal Value() As Date)
            dtShippingDate = Value
        End Set

    End Property

    'Total amount of the Transaction before deductions
    Public Property TransactionShippingAgentEmployerName() As String()

        Get
            Return strShippingAgentEmployerName
        End Get

        Set(ByVal Value() As String)
            strShippingAgentEmployerName = Value
        End Set

    End Property

    'Total amount of the Transaction before deductions
    Public Property TransactionProductName() As String()

        Get
            Return strTransactionProductName
        End Get

        Set(ByVal Value() As String)
            strTransactionProductName = Value
        End Set

    End Property

    'Total amount of the Transaction before deductions
    Public Property TransactionProductTextCode() As String()

        Get
            Return strTransactionProductTextCode
        End Get

        Set(ByVal Value() As String)
            strTransactionProductTextCode = Value
        End Set

    End Property

    'Total amount of the Transaction before deductions
    Public ReadOnly Property TransactionVATPercentage() As Double()

        Get
            Return dbTransactionVATPercentage
        End Get

        'Set(ByVal Value() As Double)
        '    dbTransactionVATPercentage = Value
        'End Set

    End Property

    'Total amount of the Transaction before deductions
    Public Property TransactionGrossTotalAmount() As Decimal()

        Get
            Return decTransactionGrossTotalAmount
        End Get

        Set(ByVal Value() As Decimal)
            decTransactionGrossTotalAmount = Value
        End Set

    End Property

    'Percentage required to be duducted from Gross
    Public Property TransactionDiscountPercentage() As Double()

        Get
            Return dbTransactionDiscountPercentage
        End Get

        Set(ByVal Value() As Double)
            dbTransactionDiscountPercentage = Value
        End Set

    End Property

    '[Discount for each transaction
    Public ReadOnly Property TransactionDiscount() As Decimal()

        Get
            Return decTransactionDiscount
        End Get

        'Set(ByVal Value() As Decimal)
        '    decTransactionDiscount = Value
        'End Set

    End Property

    'Transaction VAT calculated after discounts
    Public ReadOnly Property TransactionVATAmount() As Decimal()

        Get
            Return decTransactionVATAmount
        End Get

        'Set(ByVal Value() As Decimal)
        '    decTransactionVATAmount = Value
        'End Set

    End Property

    'Amount paid for a specific transaction
    Public Property AmountPaid() As Decimal()

        Get
            Return decAmountPaid
        End Get

        Set(ByVal Value() As Decimal)
            decAmountPaid = Value
        End Set

    End Property

    'Total amount of the Transaction after VAT, discount, and amount paid deductions
    Public ReadOnly Property TransactionNetAmount() As Decimal()

        Get
            Return decTransactionNetAmount
        End Get


    End Property

    Public Property CustomerPrepayment() As Decimal()

        Get
            Return decCustomerPrepayment
        End Get

        Set(ByVal Value() As Decimal)
            decCustomerPrepayment = Value
        End Set

    End Property

    Public Property TransactionVATCode() As String()

        Get
            Return strTransactionVATCode
        End Get

        Set(ByVal Value() As String)
            strTransactionVATCode = Value
        End Set

    End Property

    Public Property TransactionRunningNo() As Long

        Get
            Return lTransactionRunningNo
        End Get

        Set(ByVal Value As Long)
            lTransactionRunningNo = Value
        End Set

    End Property

    Public Property TransactionID() As Long

        Get
            Return lTransactionID
        End Get

        Set(ByVal Value As Long)
            lTransactionID = Value
        End Set

    End Property

    Public Property TransactionSerialNo() As String()

        Get
            Return strTransactionSerialNo
        End Get

        Set(ByVal Value() As String)
            strTransactionSerialNo = Value
        End Set

    End Property

    Public Property TransactionDate() As Date()

        Get
            Return dtTransactionDate
        End Get

        Set(ByVal Value() As Date)
            dtTransactionDate = Value
        End Set

    End Property

    Public Property TransactionBookDate() As Date()

        Get
            Return dtTransactionBookDate
        End Get

        Set(ByVal Value() As Date)
            dtTransactionBookDate = Value
        End Set

    End Property

    Public Property COAAccountNr() As Long()

        Get
            Return lCOAAccountNr
        End Get

        Set(ByVal Value() As Long)
            lCOAAccountNr = Value
        End Set

    End Property

    Public Property Payee() As String()

        Get
            Return strPayee
        End Get

        Set(ByVal Value() As String)
            strPayee = Value
        End Set

    End Property

    Public Property MethodOfPaymentID() As Long()

        Get
            Return lMethodOfPaymentID
        End Get

        Set(ByVal Value() As Long)
            lMethodOfPaymentID = Value
        End Set

    End Property


    Public Property CurrencyUsed() As Decimal()

        Get
            Return decCurrencyUsed
        End Get

        Set(ByVal Value() As Decimal)
            decCurrencyUsed = Value
        End Set

    End Property

    Public Property ItemPaidFor() As String()

        Get
            Return strItemPaidFor
        End Get

        Set(ByVal Value() As String)
            strItemPaidFor = Value
        End Set

    End Property

    Public Property DocumentID() As Long()

        Get
            Return lDocumentID
        End Get

        Set(ByVal Value() As Long)
            lDocumentID = Value
        End Set

    End Property

    Public Property UserID() As Long()

        Get
            Return lUserID
        End Get

        Set(ByVal Value() As Long)
            lUserID = Value
        End Set

    End Property

    Public Property SequenceGroupID() As Long()

        Get
            Return lSequenceGroupID
        End Get

        Set(ByVal Value() As Long)
            lSequenceGroupID = Value
        End Set

    End Property

    Public Property TransactionDescription() As String()

        Get
            Return strTransactionDescription
        End Get

        Set(ByVal Value() As String)
            strTransactionDescription = Value
        End Set

    End Property

    Public Property ApprovingOfficerUserID() As Long()

        Get
            Return lApprovingOfficerUserID
        End Get

        Set(ByVal Value() As Long)
            lApprovingOfficerUserID = Value
        End Set

    End Property

    Public Property TransactionInvoiceNo() As String()

        Get
            Return strTransactionInvoiceNo
        End Get

        Set(ByVal Value() As String)
            strTransactionInvoiceNo = Value
        End Set

    End Property

    Public ReadOnly Property DateCreated() As Date()

        Get
            Return dtDateCreated
        End Get

        'Set(ByVal Value() As Date)
        '    dtDateCreated = Value
        'End Set

    End Property

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "GeneralProcedures"

    Public Function RedimAllArrays(ByVal lRedimSize As Long)

        Try

            ReDim strTransactionSerialNo(lRedimSize)
            ReDim dtTransactionDate(lRedimSize)
            ReDim dtTransactionBookDate(lRedimSize)
            ReDim lCOAAccountNr(lRedimSize)
            ReDim strPayee(lRedimSize)
            ReDim lMethodOfPaymentID(lRedimSize)
            ReDim decCurrencyUsed(lRedimSize)
            ReDim strItemPaidFor(lRedimSize)
            ReDim lDocumentID(lRedimSize)
            ReDim lUserID(lRedimSize)
            ReDim lSequenceGroupID(lRedimSize)
            ReDim strTransactionDescription(lRedimSize)
            ReDim lApprovingOfficerUserID(lRedimSize)
            ReDim strTransactionInvoiceNo(lRedimSize)
            ReDim dtDateCreated(lRedimSize)
            ReDim strTransactionVATCode(lRedimSize)
            ReDim decCustomerPrepayment(lRedimSize)
            ReDim decTransactionGrossTotalAmount(lRedimSize)
            ReDim dbTransactionDiscountPercentage(lRedimSize)
            ReDim decTransactionDiscount(lRedimSize)
            ReDim dbTransactionVATPercentage(lRedimSize)
            ReDim decTransactionVATAmount(lRedimSize)
            ReDim decAmountPaid(lRedimSize)
            ReDim decTransactionNetAmount(lRedimSize)
            ReDim strTransactionProductName(lRedimSize)
            ReDim strTransactionProductTextCode(lRedimSize)
            ReDim strShippingAgentEmployerName(lRedimSize)
            ReDim dtShippingDate(lRedimSize)
            ReDim bTransactionIsDebit(lRedimSize)

        Catch ex As Exception

        End Try

    End Function

    Public Function NullifyAllArrays()

        Try

            strTransactionSerialNo = Nothing
            dtTransactionDate = Nothing
            dtTransactionBookDate = Nothing
            lCOAAccountNr = Nothing
            strPayee = Nothing
            lMethodOfPaymentID = Nothing
            decCurrencyUsed = Nothing
            strItemPaidFor = Nothing
            lDocumentID = Nothing
            lUserID = Nothing
            lSequenceGroupID = Nothing
            strTransactionDescription = Nothing
            lApprovingOfficerUserID = Nothing
            strTransactionInvoiceNo = Nothing
            dtDateCreated = Nothing
            strTransactionVATCode = Nothing
            decCustomerPrepayment = Nothing
            decTransactionGrossTotalAmount = Nothing
            dbTransactionDiscountPercentage = Nothing
            decTransactionDiscount = Nothing
            dbTransactionVATPercentage = Nothing
            decTransactionVATAmount = Nothing
            decAmountPaid = Nothing
            decTransactionNetAmount = Nothing
            strTransactionProductName = Nothing
            strTransactionProductTextCode = Nothing
            strShippingAgentEmployerName = Nothing
            dtShippingDate = Nothing
            bTransactionIsDebit = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnProductCOABalaneSheet()

        Try
            Dim objProd As IMProducts = New IMProducts
            Dim objDec As Object = New Decimal
            Dim decTotalDebit As Decimal
            Dim i As Long

            If Not decTransactionNetAmount Is Nothing Then
                For Each objDec In decTransactionNetAmount
                    If bTransactionUseProduct(i) = True Then
                        lCOAAccountNr(i) = objProd.ReturnProductCOABalanceSheet _
                            (objProd.ReturnProductID(strTransactionProductName(i), _
                                    strTransactionProductTextCode(i)))


                    End If

                    i = i + 1

                Next
            End If

            objDec = Nothing
            Return decTotalDebit

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnProductCOAProfitAndLoss()

        Try

            Dim objProd As IMProducts = New IMProducts
            Dim objDec As Object = New Decimal
            Dim decTotalDebit As Decimal
            Dim i As Long

            If Not decTransactionNetAmount Is Nothing Then
                For Each objDec In decTransactionNetAmount
                    If bTransactionUseProduct(i) = True Then
                        lCOAAccountNr(i) = objProd.ReturnProductCOABalanceSheet _
                            (objProd.ReturnProductID(strTransactionProductName(i), _
                                    strTransactionProductTextCode(i)))


                    End If

                    i = i + 1

                Next
            End If

            objDec = Nothing
            Return decTotalDebit

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnTransactionTotalDebits() As Decimal

        Try

            Dim objDec As Object = New Decimal
            Dim decTotalDebit As Decimal
            Dim i As Long

            If Not decTransactionNetAmount Is Nothing Then
                For Each objDec In decTransactionNetAmount
                    If bTransactionIsDebit(i) = True Then
                        decTotalDebit = decTotalDebit + objDec

                    End If

                Next
            End If

            objDec = Nothing
            Return decTotalDebit

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnTransactionTotalCredit() As Decimal

        Try

            Dim objDec As Object = New Decimal
            Dim decTotalCredit As Decimal
            Dim i As Long

            If Not decTransactionNetAmount Is Nothing Then
                For Each objDec In decTransactionNetAmount
                    If bTransactionIsDebit(i) = False Then
                        decTotalCredit = decTotalCredit + objDec

                    End If

                Next
            End If

            objDec = Nothing
            Return decTotalCredit

        Catch ex As Exception

        End Try

    End Function

    '[Gets the next TransactionID
    Private Function CalculateNextTransactionID() As Long

        Try

            Dim objLogin As IMLogin = New IMLogin

            With objLogin
                .ReturnMaxLongValue(strOrgAccessConnString, _
                    "SELECT Max(TransactionID) AS TrIDMax FROM " & _
                    " Transactions")

            End With

            objLogin = Nothing

        Catch ex As Exception
            MsgBox(ex.Message.ToString, _
                MsgBoxStyle.Critical, _
                    "iManagement - System Error")
        End Try

    End Function

    '[Gets the next invoice number
    Private Function CalculateNextTransactionRunningNo() As Long

        Try

            Dim objLogin As IMLogin = New IMLogin

            With objLogin
                .ReturnMaxLongValue(strOrgAccessConnString, _
                    "SELECT Max(TransactionRunningNo) AS TrNoMax FROM " & _
                "Transactions WHERE TransactionID = " & _
                    lTransactionID + 1)

            End With

            objLogin = Nothing

        Catch ex As Exception
            MsgBox(ex.Message.ToString, _
                MsgBoxStyle.Critical, _
                    "iManagement - System Error")
        End Try

    End Function

    '[Fill the decTransactionDiscount array with the Discount amounts
    Public Function CalculateTransactionDiscounts() As Decimal

        Try

            Dim decRetValue As Decimal
            Dim objDecItem As Object = New Decimal
            Dim objDBItem As Object = New Double
            Dim i As Integer

            If Not dbTransactionDiscountPercentage Is Nothing Then
                If Not decTransactionGrossTotalAmount Is Nothing Then

                    ReDim decTransactionDiscount _
                        (decTransactionGrossTotalAmount.GetLongLength(0))

                    For Each objDecItem In decTransactionGrossTotalAmount
                        If Not objDecItem Is Nothing Then

                            decTransactionDiscount(i) = _
                                decTransactionGrossTotalAmount(i) * _
                                    (dbTransactionDiscountPercentage(i) / 100)

                        End If
                        i = i + 1

                    Next

                End If
            End If

            objDBItem = Nothing
            objDecItem = Nothing

            Return decRetValue

        Catch ex As Exception

        End Try

    End Function

    '[Fill decTransactionVATAmount
    Public Function CalculateTransactionVATAmounts() As Decimal

        Try

            Dim decRetValue As Decimal
            Dim objDecItem As Object = New Decimal
            Dim objDBItem As Object = New Double
            Dim i As Integer

            If Not dbTransactionVATPercentage Is Nothing Then
                If Not decTransactionGrossTotalAmount Is Nothing Then

                    ReDim decTransactionVATAmount _
                        (decTransactionGrossTotalAmount.GetLongLength(0))

                    For Each objDecItem In decTransactionGrossTotalAmount
                        If Not objDecItem Is Nothing Then

                            decTransactionVATAmount(i) = _
                                (decTransactionGrossTotalAmount(i) - _
                                (decTransactionGrossTotalAmount(i) * _
                                dbTransactionDiscountPercentage(i) / 100)) * _
                                dbTransactionVATPercentage(i) / 100

                        End If
                        i = i + 1
                    Next

                End If
            End If

            objDBItem = Nothing
            objDecItem = Nothing

        Catch ex As Exception

        End Try

    End Function

    '[Total of all transaction discounts
    Public Function ReturnTotalTransactionDiscounts() As Decimal

        Try

            Dim decRetValue As Decimal
            Dim objItem As Object = New Decimal

            If Not decTransactionDiscount Is Nothing Then
                For Each objItem In decTransactionDiscount
                    If Not objItem Is Nothing Then
                        decRetValue = decRetValue + objItem

                    End If
                Next
            End If

            objItem = Nothing

            Return decRetValue

        Catch ex As Exception

        End Try

    End Function

    '[VAT Total for all transactions
    Public Function ReturnTotalTransactionVATTotal() As Decimal

        Try

            Dim decRetValue As Decimal
            Dim objItem As Object = New Decimal

            If Not decTransactionVATAmount Is Nothing Then
                For Each objItem In decTransactionVATAmount
                    If Not objItem Is Nothing Then
                        decRetValue = decRetValue + objItem

                    End If
                Next
            End If

            objItem = Nothing
            Return decRetValue

        Catch ex As Exception

        End Try

    End Function

    '[Total of all transaction discounts
    Public Function ReturnTotalAmountPaid() As Decimal

        Try

            Dim decRetValue As Decimal
            Dim objItem As Object = New Decimal

            If Not decAmountPaid Is Nothing Then
                For Each objItem In decAmountPaid
                    If Not objItem Is Nothing Then
                        decRetValue = decRetValue + objItem

                    End If
                Next
            End If

            objItem = Nothing

            Return decRetValue

        Catch ex As Exception

        End Try

    End Function

    '[Total of all transaction discounts
    Public Function ReturnCustomerPrepayment() As Decimal

        Try

            Dim decRetValue As Decimal
            Dim objItem As Object = New Decimal

            If Not decCustomerPrepayment Is Nothing Then
                For Each objItem In decCustomerPrepayment
                    If Not objItem Is Nothing Then
                        decRetValue = decRetValue + objItem

                    End If
                Next
            End If

            objItem = Nothing

            Return decRetValue

        Catch ex As Exception

        End Try

    End Function

    '[Fill the TrasactionNetAmount
    Public Function CalculateTransactionNetAmount() As Decimal

        Try

            Dim decRetValue As Decimal
            Dim objDecItem As Object = New Decimal
            Dim objDBItem As Object = New Double
            Dim i As Integer

            If Not decTransactionGrossTotalAmount Is Nothing Then

                ReDim decTransactionNetAmount _
                    (decTransactionGrossTotalAmount.GetLongLength(0))

                For Each objDecItem In decTransactionGrossTotalAmount
                    If Not objDecItem Is Nothing Then

                        If decTransactionDiscount Is Nothing And _
                            decTransactionVATAmount Is Nothing Then

                            decTransactionNetAmount(i) = _
                                decTransactionGrossTotalAmount(i)


                        ElseIf decTransactionDiscount Is Nothing And _
                            Not decTransactionVATAmount Is Nothing Then

                            decTransactionNetAmount(i) = _
                                decTransactionGrossTotalAmount(i) + _
                                    decTransactionVATAmount(i)


                        ElseIf Not decTransactionDiscount Is Nothing And _
                            decTransactionVATAmount Is Nothing Then

                            decTransactionNetAmount(i) = _
                                (decTransactionGrossTotalAmount(i) - _
                                    decTransactionDiscount(i))

                        ElseIf Not decTransactionDiscount Is Nothing And _
                           Not decTransactionVATAmount Is Nothing Then

                            decTransactionNetAmount(i) = _
                                (decTransactionGrossTotalAmount(i) + _
                                    decTransactionVATAmount(i)) - _
                                        decTransactionDiscount(i)

                        End If

                    End If

                    i = i + 1

                Next

            End If

            objDBItem = Nothing
            objDecItem = Nothing

            Return decRetValue

        Catch ex As Exception

        End Try

    End Function

    '[Fill the VAT Percentages
    Public Function CalculateVATPercentages() As Double

        Try

            Dim objTaxCode As IMTaxes = New IMTaxes
            Dim strItem As String
            Dim i As Integer

            If strTransactionVATCode Is Nothing Then
                objTaxCode = Nothing
                Exit Function
            End If

            ReDim dbTransactionVATPercentage _
                (strTransactionVATCode.GetLongLength(0))

            If Not strTransactionVATCode Is Nothing Then
                For Each strItem In strTransactionVATCode

                    With objTaxCode
                        .Find("SELECT * FROM TaxCodes WHERE TaxCodeID = '" & _
                            strTransactionVATCode(i) & "'", True)

                        dbTransactionVATPercentage(i) = .TaxPercentage
                        .NewDetails()

                    End With
                    i = i + 1
                Next
            End If

            strItem = Nothing
            objTaxCode = Nothing

        Catch ex As Exception

        End Try

    End Function

#End Region


#Region "DatabaseProcedures"

    Public Function Save(ByVal bDisplayErrorMessages As Boolean, _
        ByVal bDisplayConfirmation As Boolean, _
            ByVal bDisplayFailure As Boolean, _
                ByVal bDisplaySuccess As Boolean) As Boolean

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String
        Dim strItem As String
        Dim lItem As Long
        Dim m As Object = New Object
        Dim decItem As Decimal
        Dim i As Long
        Dim objProduct As IMProducts
        Dim objEmployers As IMEmployers

        m = lItem

        Try


            If lCOAAccountNr Is Nothing Then
                If bDisplayErrorMessages = True Then
                    ReturnError += "Please provide the following details in " & _
                                " order to save a Transaction. " & _
                                  Chr(10) & "1. A Transaction's Details. " & _
                                  Chr(10) & "2. A Chart Of Account for the transaction. " & _
                                  Chr(10) & "3. Description of the transaction. "
                End If

                m = Nothing
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If strTransactionDescription Is Nothing Then
                If bDisplayErrorMessages = True Then

                    returnerror += "Please provide the following details in" & _
                     " order to save a Transaction." & _
                       Chr(10) & "1. A Transaction's Details." & _
                       Chr(10) & "2. A Chart Of Account for the transaction." & _
                       Chr(10) & "3. Description of the transaction."

                End If

                m = Nothing
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If



            For Each m In lCOAAccountNr
                If Not m Is Nothing Then
                    If m = 0 Then

                        If bDisplayErrorMessages = True Then

                            ReturnError += "Please provide the following details in " & _
                                        " order to save a Transaction. " & _
                                          Chr(10) & "1. A Transaction's Details. " & _
                                          Chr(10) & "2. A Chart Of Account for the transaction. " & _
                                          Chr(10) & "3. Description of the transaction. "
                        End If

                        m = Nothing
                        objLogin = Nothing
                        datSaved = Nothing

                        Exit Function

                    End If
                End If
            Next


            For Each strItem In strTransactionDescription
                If Not strItem Is Nothing Then
                    If strItem = "" Then

                        If bDisplayErrorMessages = True Then

                            ReturnError += "Please provide the following details in " & _
                                       " order to save a Transaction. " & _
                                         Chr(10) & "1. A Transaction's Details. " & _
                                         Chr(10) & "2. A Chart Of Account for the transaction. " & _
                                         Chr(10) & "3. Description of the transaction. "
                        End If

                        m = Nothing
                        objLogin = Nothing
                        datSaved = Nothing

                        Exit Function

                    End If
                End If
            Next


            If Not strTransactionSerialNo Is Nothing Then
                If strTransactionSerialNo.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    ReturnError += "You must provide an equal number of elements " & _
                      "the Transactions."

                    m = Nothing

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If


            If Not dtTransactionDate Is Nothing Then
                If dtTransactionDate.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    ReturnError += "You must provide an equal number of elements " & _
                       "the Transactions."

                    m = Nothing

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If


            If Not dtTransactionBookDate Is Nothing Then
                If dtTransactionBookDate.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                     returnerror += "You must provide an equal number of elements " & _
                       "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not strPayee Is Nothing Then
                If strPayee.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    ReturnError += "You must provide an equal number of elements " & _
                        "the Transactions."

                    m = Nothing

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If


            If Not lMethodOfPaymentID Is Nothing Then
                If lMethodOfPaymentID.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    ReturnError += "You must provide an equal number of elements " & _
                      "the Transactions."

                    m = Nothing

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If


            If Not decAmountPaid Is Nothing Then
                If decAmountPaid.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    ReturnError += "You must provide an equal number of elements " & _
                      "the Transactions."

                    m = Nothing

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not decCurrencyUsed Is Nothing Then
                If decCurrencyUsed.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                     returnerror += "You must provide an equal number of elements " & _
                       "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not strItemPaidFor Is Nothing Then
                If strItemPaidFor.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    ReturnError += "You must provide an equal number of elements " & _
                      "the Transactions."

                    m = Nothing

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not lDocumentID Is Nothing Then
                If lDocumentID.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    ReturnError += "You must provide an equal number of elements " & _
                      "the Transactions."

                    m = Nothing

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not lUserID Is Nothing Then
                If lUserID.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    ReturnError += "You must provide an equal number of elements " & _
                       "the Transactions."

                    m = Nothing

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not lSequenceGroupID Is Nothing Then
                If lSequenceGroupID.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                     returnerror += "You must provide an equal number of elements " & _
                       "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not strTransactionDescription Is Nothing Then
                If strTransactionDescription.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                     returnerror += "You must provide an equal number of elements " & _
                       "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not lApprovingOfficerUserID Is Nothing Then
                If lApprovingOfficerUserID.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                     returnerror += "You must provide an equal number of elements " & _
                       "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If


            If Not strTransactionInvoiceNo Is Nothing Then
                If strTransactionInvoiceNo.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                     returnerror += "You must provide an equal number of elements " & _
                       "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If


            If Not dtDateCreated Is Nothing Then
                If dtDateCreated.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then
                     returnerror += "You must provide an equal number of elements " & _
                        "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If



            If Not decCustomerPrepayment Is Nothing Then
                If decCustomerPrepayment.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    ReturnError += "You must provide an equal number of elements " & _
                       "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If


            If Not decTransactionVATAmount Is Nothing Then
                If decTransactionVATAmount.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then
                   returnerror += "You must provide an equal number of elements " & _
                        "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If


            If Not strTransactionVATCode Is Nothing Then
                If strTransactionVATCode.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                     returnerror += "You must provide an equal number of elements " & _
                        "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If


            If Not decTransactionDiscount Is Nothing Then
                If decTransactionDiscount.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                     returnerror += "You must provide an equal number of elements " & _
                        "the Transactions."

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If



            'If bDisplayConfirmation = True Then
            '    If MsgBox("Do you want to add this new Transaction?", _
            '        MsgBoxStyle.YesNo, _
            '            "iManagement - Add Transaction Details?") _
            '                = MsgBoxResult.No Then

            '        m = Nothing
            '        objLogin = Nothing
            '        datSaved = Nothing

            '        Exit Function

            '    End If
            'End If


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            lTransactionID = CalculateNextTransactionID()


            CalculateVATPercentages()
            CalculateTransactionDiscounts()
            CalculateTransactionVATAmounts()
            CalculateTransactionNetAmount()


            objProduct = New IMProducts
            objEmployers = New IMEmployers

            For Each m In lCOAAccountNr

                strInsertInto = "INSERT INTO Transactions (" & _
                        "TransactionRunningNo," & _
                        "TransactionID," & _
                        "TransactionSerialNo," & _
                        "TransactionDate," & _
                        "TransactionBookDate," & _
                        "COAAccountNr," & _
                        "Payee," & _
                        "MethodOfPaymentID," & _
                        "AmountPaid," & _
                        "CurrencyUsed," & _
                        "ItemPaidFor," & _
                        "DocumentID," & _
                        "UserID," & _
                        "SequenceGroupID," & _
                        "TransactionDescription," & _
                        "ApprovingOfficerUserID," & _
                        "TransactionInvoiceNo," & _
                        "CustomerPrepayment," & _
                        "TransactionGrossTotalAmount," & _
                        "TransactionDiscountPercentage," & _
                        "TransactionDiscount," & _
                        "TransactionVATPercentage," & _
                        "TransactionVATAmount," & _
                        "TransactionVATCode," & _
                        "TransactionNetAmount," & _
                        "ProductID," & _
                        "ShippingDate," & _
                        "ShippingAgentEmployerID" & _
                         ") VALUES "

                strSaveQuery = strInsertInto & _
                        "(" & CalculateNextTransactionRunningNo() & _
                        "," & lTransactionID & _
                        ",'" & Trim(strTransactionSerialNo(i)) & _
                        "',#" & dtTransactionDate(i) & _
                        "#,#" & dtTransactionBookDate(i) & _
                        "#," & lCOAAccountNr(i) & _
                        ",'" & Trim(strPayee(i)) & _
                        "'," & lMethodOfPaymentID(i) & _
                        "," & decAmountPaid(i) & _
                        "," & decCurrencyUsed(i) & _
                        ",'" & strItemPaidFor(i) & _
                        "'," & lDocumentID(i) & _
                        "," & lUserID(i) & _
                        "," & lSequenceGroupID(i) & _
                        ",'" & Trim(strTransactionDescription(i)) & _
                        "'," & lApprovingOfficerUserID(i) & _
                        ",'" & Trim(strTransactionInvoiceNo(i)) & _
                        "'," & decCustomerPrepayment(i) & _
                        "," & decTransactionGrossTotalAmount(i) & _
                        "," & dbTransactionDiscountPercentage(i) & _
                        "," & decTransactionDiscount(i) & _
                        "," & dbTransactionVATPercentage(i) & _
                        "," & decTransactionVATAmount(i) & _
                        ",'" & strTransactionVATCode(i) & _
                        "'," & decTransactionNetAmount(i) & _
                        "," & objProduct.ReturnProductID _
                        (strTransactionProductName(i), _
                        strTransactionProductTextCode(i)) & _
                        ",#" & dtShippingDate(i) & _
                        "#," & objEmployers.ReturnEmployerIdFromEmployerName _
                        (strShippingAgentEmployerName(i)) & _
                        ")"

                bSaveSuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strSaveQuery, datSaved)

                i = i + 1

            Next

            objProduct = Nothing
            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If bDisplaySuccess = True Then

                    MsgBox("Voucher Transaction Saved Successfully.", _
                        MsgBoxStyle.Information, _
                            "iManagement - Record Saved Successfully")

                End If
            Else

                If bDisplayFailure = True Then

                    MsgBox("'Save Transaction' action failed." & _
            " Make sure all mandatory details are entered.", _
            MsgBoxStyle.Exclamation, _
            "iManagement - Transaction Addition Failed")

                End If
            End If

            m = Nothing

            objLogin = Nothing
            datSaved = Nothing

            If bSaveSuccess = True Then
                Return True
            End If


        Catch ex As Exception
            If bDisplayErrorMessages = True Then
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
            Dim i As Long
            Dim objProd As IMProducts

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

                        RedimAllArrays(myDataTables.Rows.Count)

                        For Each myDataRows In myDataTables.Rows

                            lTransactionRunningNo = _
                                myDataRows("TransactionRunningNo").ToString
                            lTransactionID = _
                                myDataRows("TransactionID").ToString
                            strTransactionSerialNo(i) = _
                                myDataRows("TransactionSerialNo").ToString
                            dtTransactionDate(i) = _
                                myDataRows("TransactionDate")
                            dtTransactionBookDate(i) = _
                                myDataRows("TransactionBookDate")
                            lCOAAccountNr(i) = _
                                myDataRows("COAAccountNr")
                            strPayee(i) = _
                                myDataRows("Payee").ToString
                            lMethodOfPaymentID(i) = _
                                myDataRows("MethodOfPaymentID")
                            decAmountPaid(i) = _
                                myDataRows("AmountPaid")
                            decCurrencyUsed(i) = _
                                myDataRows("CurrencyUsed")
                            strItemPaidFor(i) = _
                                myDataRows("ItemPaidFor").ToString
                            lDocumentID(i) = _
                                myDataRows("DocumentID")
                            lUserID(i) = _
                                myDataRows("UserID")
                            lSequenceGroupID(i) = _
                                myDataRows("SequenceGroupID")
                            strTransactionDescription(i) = _
                                myDataRows("TransactionDescription").ToString
                            lApprovingOfficerUserID(i) = _
                                myDataRows("ApprovingOfficerUserID")
                            strTransactionInvoiceNo(i) = _
                                myDataRows("TransactionInvoiceNo").ToString
                            decTransactionVATAmount(i) = _
                                myDataRows("TransactionVATAmount")
                            strTransactionVATCode(i) = _
                                myDataRows("TransactionVATCode")
                            decTransactionDiscount(i) = _
                                myDataRows("TransactionDiscount")
                            decTransactionNetAmount(i) = _
                                myDataRows("TransactionNetAmount")
                            dbTransactionDiscountPercentage(i) = _
                                myDataRows("TransactionDiscountPercentage")
                            dbTransactionVATPercentage(i) = _
                                myDataRows("TransactionVATPercentage")
                            decTransactionGrossTotalAmount(i) = _
                                myDataRows("TransactionGrossTotalAmount")
                            dtShippingDate = _
                                myDataRows("ShippingDate")

                            Dim objEmployers As IMEmployers = New IMEmployers

                            With objEmployers
                                strShippingAgentEmployerName(i) = _
                                  .ReturnEmployerNameFromEmployerID _
                                    (myDataRows("ShippingAgentEmployerID"))

                                objEmployers.NewRecord()

                            End With


                            i = i + 1

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

    Public Function Delete() As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin


            If lTransactionID = 0 Then
                MsgBox("Cannot Delete. Please select an existing " & _
                    "Transaction Detail.", MsgBoxStyle.Exclamation, _
                        "iManagement - Invalid or incomplete Information")

                objLogin = Nothing
                datDelete = Nothing
                Exit Function

            End If


            If MsgBox("Are you sure you want to delete this Transaction's" & _
                " details and its associated Trace details?", MsgBoxStyle.YesNo, _
                    "iManagement - Delete Transaction?") _
                        = MsgBoxResult.No Then

                objLogin = Nothing
                datDelete = Nothing
                Exit Function
            End If


            If Find("SELECT * FROM Transactions WHERE " & _
                       " TransactionID = " & lTransactionID & _
                    " AND ApprovingOfficerUserID IS NOT NULL OR " & _
                        " AND ApprovingOfficerUserID <> 0 ", False) = True Then

                MsgBox("Cannot delete approved Transactions", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Cannot Delee Record")

                objLogin = Nothing
                datDelete = Nothing
                Exit Function

            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()



            strDeleteQuery = "DELETE * FROM TransactionTraceDetails WHERE " & _
            " TransactionID = " & lTransactionID

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)


            strDeleteQuery = "DELETE * FROM Transaction WHERE " & _
                      " TransactionID = " & TransactionID

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Transaction Details Deleted", _
                    MsgBoxStyle.Information, _
                        "iManagement - Record Deleted Successfully")

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

    Public Function MakeContraEntry(ByVal bDisplayErrorMessages As Boolean, _
        ByVal bDisplayConfirmation As Boolean, _
            ByVal bDisplayFailure As Boolean, _
                ByVal bDisplaySuccess As Boolean)

        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String
            Dim strItem As String
            Dim lItem As Long
            Dim m As Object = lItem
            Dim decItem As Decimal
            Dim i As Long

            Dim lOldTransactionID As Long

            lOldTransactionID = lTransactionID

            If Find("SELECT * FROM Transactions WHERE " & _
            " TransactionID = " & lOldTransactionID, True) = False Then

                m = Nothing
                datSaved = Nothing
                objLogin = Nothing

                Exit Function

            End If



            If lCOAAccountNr Is Nothing Then
                If bDisplayErrorMessages = True Then
                    MsgBox("Please provide the following details in" & _
                                " order to save a Transaction." & _
                                  Chr(10) & "1. A Transaction's Details." & _
                                  Chr(10) & "2. A Chart Of Account for the transaction." & _
                                  Chr(10) & "3. Description of the transaction." & _
                                  MsgBoxStyle.Critical, _
                                      "iManagement - Save Action Failed")
                End If

                m = Nothing
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If strTransactionDescription Is Nothing Then
                If bDisplayErrorMessages = True Then

                    MsgBox("Please provide the following details in" & _
                     " order to save a Transaction." & _
                       Chr(10) & "1. A Transaction's Details." & _
                       Chr(10) & "2. A Chart Of Account for the transaction." & _
                       Chr(10) & "3. Description of the transaction." & _
                       MsgBoxStyle.Critical, _
                           "iManagement - Save Action Failed")

                End If

                m = Nothing
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If



            For Each m In lCOAAccountNr
                If Not m Is Nothing Then
                    If m = 0 Then

                        If bDisplayErrorMessages = True Then

                            MsgBox("Please provide the following details in" & _
                                        " order to save a Transaction." & _
                                          Chr(10) & "1. A Transaction's Details." & _
                                          Chr(10) & "2. A Chart Of Account for the transaction." & _
                                          Chr(10) & "3. Description of the transaction." & _
                                          MsgBoxStyle.Critical, _
                                              "iManagement - Save Action Failed")
                        End If

                        m = Nothing
                        objLogin = Nothing
                        datSaved = Nothing

                        Exit Function

                    End If
                End If
            Next


            For Each strItem In strTransactionDescription
                If Not strItem Is Nothing Then
                    If strItem = "" Then

                        If bDisplayErrorMessages = True Then

                            MsgBox("Please provide the following details in" & _
                                        " order to save a Transaction." & _
                                          Chr(10) & "1. A Transaction's Details." & _
                                          Chr(10) & "2. A Chart Of Account for the transaction." & _
                                          Chr(10) & "3. Description of the transaction." & _
                                          MsgBoxStyle.Critical, _
                                              "iManagement - Save Action Failed")
                        End If

                        m = Nothing
                        objLogin = Nothing
                        datSaved = Nothing

                        Exit Function

                    End If
                End If
            Next


            If Not strTransactionSerialNo Is Nothing Then
                If strTransactionSerialNo.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If


            If Not dtTransactionDate Is Nothing Then
                If dtTransactionDate.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If


            If Not dtTransactionBookDate Is Nothing Then
                If dtTransactionBookDate.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not strPayee Is Nothing Then
                If strPayee.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If


            If Not lMethodOfPaymentID Is Nothing Then
                If lMethodOfPaymentID.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If


            If Not decAmountPaid Is Nothing Then
                If decAmountPaid.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not decCurrencyUsed Is Nothing Then
                If decCurrencyUsed.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not strItemPaidFor Is Nothing Then
                If strItemPaidFor.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not lDocumentID Is Nothing Then
                If lDocumentID.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not lUserID Is Nothing Then
                If lUserID.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not lSequenceGroupID Is Nothing Then
                If lSequenceGroupID.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not strTransactionDescription Is Nothing Then
                If strTransactionDescription.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            If Not lApprovingOfficerUserID Is Nothing Then
                If lApprovingOfficerUserID.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If


            If Not strTransactionInvoiceNo Is Nothing Then
                If strTransactionInvoiceNo.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then

                    MsgBox("You must provide an equal number of elements " & _
                       "the Transactions.", _
                           MsgBoxStyle.Critical, _
                               "iManagement - invalid or incomplete information")

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If


            If Not dtDateCreated Is Nothing Then
                If dtDateCreated.GetLongLength(0) <> _
                                lCOAAccountNr.GetLongLength(0) Then
                    MsgBox("You must provide an equal number of elements " & _
                        "the Transactions.", _
                            MsgBoxStyle.Critical, _
                                "iManagement - invalid or incomplete information")

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If



            If bDisplayConfirmation = True Then
                If MsgBox("Do you want to add this new Transaction?", _
                    MsgBoxStyle.YesNo, _
                        "iManagement - Add Transaction Details?") _
                            = MsgBoxResult.No Then

                    m = Nothing
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            lTransactionID = CalculateNextTransactionID()

            For Each m In lCOAAccountNr

                strInsertInto = "INSERT INTO Transactions (" & _
                         "TransactionRunningNo," & _
                         "TransactionID," & _
                         "TransactionSerialNo," & _
                         "TransactionDate," & _
                         "TransactionBookDate," & _
                         "COAAccountNr," & _
                         "Payee," & _
                         "MethodOfPaymentID," & _
                         "AmountPaid," & _
                         "CurrencyUsed," & _
                         "ItemPaidFor," & _
                         "DocumentID," & _
                         "UserID," & _
                         "SequenceGroupID," & _
                         "TransactionDescription," & _
                         "ApprovingOfficerUserID," & _
                         "TransactionInvoiceNo," & _
                         "TransactionDiscount," & _
                         "TransactionVATAmount," & _
                         "TransactionVATCode," & _
                         "TransactionDiscount" & _
                             ") VALUES "

                strSaveQuery = strInsertInto & _
                        "(" & CalculateNextTransactionRunningNo() & _
                        "," & lTransactionID & _
                        ",'" & Trim(strTransactionSerialNo(i)) & _
                        "',#" & dtTransactionDate(i) & _
                        "#,#" & dtTransactionBookDate(i) & _
                        "#," & lCOAAccountNr(i) & _
                        ",'" & Trim(strPayee(i)) & _
                        "'," & lMethodOfPaymentID(i) & _
                        "," & decAmountPaid(i) & _
                        "," & decCurrencyUsed(i) & _
                        ",'" & strItemPaidFor(i) & _
                        "'," & lDocumentID(i) & _
                        "," & lUserID(i) & _
                        "," & lSequenceGroupID(i) & _
                        ",'" & Trim(strTransactionDescription(i)) & _
                        "'," & lApprovingOfficerUserID(i) & _
                        ",'" & Trim(strTransactionInvoiceNo(i)) & _
                        "'," & decTransactionVATAmount(i) & _
                        "," & TransactionVATCode(i) & _
                        "," & decTransactionDiscount(i) & _
                                ")"

                bSaveSuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strSaveQuery, datSaved)

                i = i + 1

            Next

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If bDisplaySuccess = True Then

                    MsgBox("Voucher Transaction Saved Successfully.", _
                        MsgBoxStyle.Information, _
                            "iManagement - Record Saved Successfully")

                End If
            Else

                If bDisplayFailure = True Then

                    MsgBox("'Save Transaction' action failed." & _
                    " Make sure all mandatory details are entered.", _
                    MsgBoxStyle.Exclamation, _
                    "iManagement - Transaction Addition Failed")

                End If
            End If


            objLogin = Nothing
            datSaved = Nothing

            If bSaveSuccess = True Then
                Return True
            End If


        Catch ex As Exception
            If bDisplayErrorMessages = True Then
                MsgBox(ex.Message.ToString, MsgBoxStyle.Exclamation, _
                        "iManagement - Critical System Error")
            End If
        End Try

    End Function


#End Region

End Class
