Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMInvoice

#Region "PrivateVariables"

    Private strInvoiceNo As String 'DDMMMYYYYNo where No is next number in series
    Private dtInvoiceDate As Date
    Private dtInvoiceMaturityDate As Date
    Private decInvoicedAmount As Decimal
    Private decVATAmount As Decimal
    Private decGrandTotal As Decimal
    Private lDocumentID As Long
    Private strClerkID As String
    Private lVATCodeID As Long
    Private lActivityID As Long
    Private strPINNo As String

    Private bInvoiceStatus As Boolean
    Private decAmountPaid As Decimal
    Private decBalance As Decimal
    Private lCustomerNo As Long
    Private lCostCentreID As Long
    Private strInvoiceType As String 'Determines whether the invoice is a debit or credit
    Private strNextInvNo As String 'Yes,No from interface. Certified or NULL from system. Certified is stronger than the rest
    Private strSequenceGroupID As String
    Private strTransactionType As String
    Private MaxValue As Long
    Private decTotalDiscountAmount As Decimal 'Total amount to be discounted
    Private dbInvDiscountPercentage As Double 'discount set on the entire invoice. Percentage on total
    Private decTotalTransactionDiscounts As Decimal 'Total of Discount amounts from individual tranactions.
    Private strTerms As Decimal
    Private strCurrencyID As String
    Private decChange As Decimal
    Private bIsDebit As Decimal
    Private strShippingAgentName As String
    Private dtShippingDate As Date
    Private dtInvoiceDueDate As Date
#End Region

#Region "Properties"

    Public Property TotalTransactionDiscounts() As Decimal

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return decTotalTransactionDiscounts
        End Get

        Set(ByVal Value As Decimal)
            decTotalTransactionDiscounts = Value
        End Set

    End Property

    Public Property InvDiscountPercentage() As Double

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return dbInvDiscountPercentage
        End Get

        Set(ByVal Value As Double)
            dbInvDiscountPercentage = Value
        End Set

    End Property

    Public Property TotalDiscountAmount() As Decimal

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return decTotalDiscountAmount
        End Get

        Set(ByVal Value As Decimal)
            decTotalDiscountAmount = Value
        End Set

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

    Public Property SequenceGroupID() As String

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return strSequenceGroupID
        End Get

        Set(ByVal Value As String)
            strSequenceGroupID = Value
        End Set

    End Property

    Public Property CostCentre() As String

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return lCostCentreID
        End Get

        Set(ByVal Value As String)
            lCostCentreID = Value
        End Set

    End Property

    Public Property InvoiceType() As String

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return strInvoiceType
        End Get

        Set(ByVal Value As String)
            strInvoiceType = Value
        End Set

    End Property

    Public Property NextInvNo() As String

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return strNextInvNo
        End Get

        Set(ByVal Value As String)
            strNextInvNo = Value
        End Set

    End Property

    Public Property InvoiceNo() As String

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return strInvoiceNo
        End Get

        Set(ByVal Value As String)
            strInvoiceNo = Value
        End Set

    End Property

    Public Property InvoiceDate() As Date

        'USED TO SET AND RETRIEVE THE BANK NAME (STRING)
        Get
            Return dtInvoiceDate
        End Get

        Set(ByVal Value As Date)
            dtInvoiceDate = Value
        End Set

    End Property

    Public Property InvoiceMaturityDate() As Date

        'USED TO SET AND RETRIEVE THE PHYSICAL ADDRESS (STRING)
        Get
            Return dtInvoiceMaturityDate
        End Get

        Set(ByVal Value As Date)
            dtInvoiceMaturityDate = Value
        End Set

    End Property

    Public Property InvoicedAmount() As Decimal

        'USED TO SET AND RETRIEVE THE POSTAL ADDRESS (STRING)
        Get
            Return decInvoicedAmount
        End Get

        Set(ByVal Value As Decimal)
            decInvoicedAmount = Value
        End Set

    End Property

    Public Property VATAmount() As Decimal

        'USED TO SET AND RETRIEVE THE POSTCODE (STRING)
        Get
            Return decVATAmount
        End Get

        Set(ByVal Value As Decimal)
            decVATAmount = Value
        End Set

    End Property

    Public Property GrandTotal() As Decimal

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return decGrandTotal
        End Get

        Set(ByVal Value As Decimal)
            decGrandTotal = Value
        End Set

    End Property

    Public Property DocumentID() As Long

        'USED TO SET AND RETRIEVE THE POST CITY CODE (STRING)
        Get
            Return lDocumentID
        End Get

        Set(ByVal Value As Long)
            lDocumentID = Value
        End Set

    End Property

    Public Property ClerkID() As String

        'USED TO SET AND RETRIEVE THE POST TOWN CODE (STRING)
        Get
            Return strClerkID
        End Get

        Set(ByVal Value As String)
            strClerkID = Value
        End Set

    End Property

    Public Property VATCodeID() As Long

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return lVATCodeID
        End Get

        Set(ByVal Value As Long)
            lVATCodeID = Value
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

    Public Property InvoiceExpiryDate() As Date

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return dtInvoiceDueDate
        End Get

        Set(ByVal Value As Date)
            dtInvoiceDueDate = Value
        End Set

    End Property

    Public Property InvoiceStatus() As Boolean

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return bInvoiceStatus
        End Get

        Set(ByVal Value As Boolean)
            bInvoiceStatus = Value
        End Set

    End Property

    Public Property AmountPaid() As Decimal

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return decAmountPaid
        End Get

        Set(ByVal Value As Decimal)
            decAmountPaid = Value
        End Set

    End Property

    Public Property Balance() As Decimal

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return decBalance
        End Get

        Set(ByVal Value As Decimal)
            decBalance = Value
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

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Sub NewRecord()

        strInvoiceNo = ""
        dtInvoiceDate = Now
        dtInvoiceMaturityDate = Now
        decInvoicedAmount = 0
        decVATAmount = 0
        decGrandTotal = 0
        lDocumentID = 0
        strClerkID = ""
        lVATCodeID = 0
        strPINNo = ""
        dtInvoiceDueDate = Now
        bInvoiceStatus = False
        decAmountPaid = 0
        decBalance = 0
        lCustomerNo = 0
        strSequenceGroupID = ""
        strTransactionType = ""

    End Sub

    Public Function RetrieveInvoiceSettings() As Boolean
        'Get invoice default series if series is not set

        'Get invoice next series number

        'Set the invoice type
        If Trim(strInvoiceType) = "" Then

        End If

    End Function

    Public Function SaveInvoiceSettings() As Boolean

        'Set invoice default series number

    End Function

    '[Gets the next invoice number
    Private Function CalculateNextInvNo() As String
        Try


            Dim objSequence As IMSequence = New IMSequence

            With objSequence

                If .ReturnNextSequenceNo(False, False, False) = 0 Then
                    objSequence = Nothing
                    Exit Function
                End If

                Return Microsoft.VisualBasic.DateAndTime.Day(Now()) _
                    & Month(Now()) & _
                        Year(Now()) & strSequenceGroupID & _
                                        .ReturnNextSequenceNo _
                                            (False, False, False) + 1

            End With

            objSequence = Nothing

        Catch ex As Exception
            MsgBox(ex.Message.ToString, _
                MsgBoxStyle.Critical, _
                    "iManagement - System Error")
        End Try

    End Function

#End Region

#Region "DatabaseProcedures"

    Public Function Save(ByVal DisplayMessages As Boolean) As Boolean
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        With objLogin

            Try

                objLogin.BeginTheTrans()

                If Trim(strInvoiceNo) = "" _
                                    Then


                    MsgBox("Cannot save Invoice Details. Missing information", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - invalid or incomplete Information")

                    datSaved = Nothing
                    objLogin = Nothing
                    Exit Function

                End If


                If Find("SELECT * FROM InvoiceMaster WHERE InvoiceNo = '" & _
                        strInvoiceNo & "'") = True Then

                    If MsgBox("Cannot save Invoice Details. Missing information", _
                             MsgBoxStyle.Exclamation, _
                                 "iManagement - invalid or incomplete " & _
                                    "Information") _
                                    = MsgBoxResult.Yes Then

                        Update("UPDATE InvoiceMaster SET " & _
                        "InvoiceNo = '" & Trim(strInvoiceNo) & _
                        "' , InvoiceMaturityDate = '" & dtInvoiceMaturityDate & _
                        "' , InvoicedAmount = " & decInvoicedAmount & _
                        " , VATAmount = " & decVATAmount & _
                        " , GrandTotal = " & decGrandTotal & _
                        " , DocumentID = " & lDocumentID & _
                        " , ClerkUserName = '" & Trim(strClerkID) & _
                        "' , VATCodeID = " & lVATCodeID & _
                        " , PINNo  = '" & strPINNo & _
                        "' , 'InvoiceExpiryDate  = '" & dtInvoiceDueDate & _
                        "' , InvoiceStatus = " & bInvoiceStatus & _
                        " , CurrencyID = '" & strCurrencyID & _
                        "' , SequenceGroupID = '" & Trim(strSequenceGroupID) & _
                        "' , InvoiceType = '" & Trim(strInvoiceType) & _
                        "' , Balance = " & decBalance & _
                        " , TransactionType = '" & Trim(strTransactionType) & _
                        "' , Change = " & decChange & _
                        " , IsDebit = " & bIsDebit & _
                        " , Terms = " & strTerms & _
                        " , AmountPaid = " & decAmountPaid & _
                        " , CostCentreID = " & lCostCentreID & _
                        " , DiscountAmount = " & decTotalDiscountAmount & _
                        " , CostCentreID = " & lCostCentreID & _
                        " , ShippingAgentName = '" & Trim(strShippingAgentName) & _
                        "' , ShippingDate = '" & dtShippingDate & _
                        "' , TotalDiscountAmount = " & decTotalDiscountAmount)

                    End If

                    datSaved = Nothing
                    objLogin = Nothing
                    Exit Function

                End If


                strInsertInto = "INSERT INTO InvoiceMaster (" & _
                        "InvoiceNo," & _
                        "InvoiceMaturityDate," & _
                        "InvoicedAmount," & _
                        "VATAmount," & _
                        "GrandTotal," & _
                        "DocumentID," & _
                        "ClerkUserName," & _
                        "VATCodeID," & _
                        "PINNo," & _
                        "InvoiceExpiryDate," & _
                        "InvoiceStatus," & _
                        "CurrencyID," & _
                        "SequenceGroupID," & _
                        "InvoiceType," & _
                        "Balance," & _
                        "TransactionType," & _
                        "Change," & _
                        "IsDebit," & _
                        "Terms," & _
                        "AmountPaid," & _
                        "CostCentreID," & _
                        "DiscuntAmount," & _
                        "CostCentreID," & _
                        "ShippingAgentName," & _
                        "ShippingDate," & _
                        "TotalDiscountAmount" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                            "(" & _
                        "'" & strInvoiceNo & _
                        "', '" & dtInvoiceMaturityDate & _
                        "', " & decInvoicedAmount & _
                        ", " & decVATAmount & _
                        ", " & decGrandTotal & _
                        ", " & lDocumentID & _
                        ", '" & Trim(strClerkID) & _
                        "', " & lVATCodeID & _
                        ", '" & Trim(strPINNo) & _
                        "', '" & dtInvoiceDueDate & _
                        "'," & bInvoiceStatus & _
                        ",'" & Trim(strCurrencyID) & _
                        "','" & Trim(strSequenceGroupID) & _
                        "','" & Trim(strInvoiceType) & _
                        "'," & decBalance & _
                        ",'" & Trim(strTransactionType) & _
                        "'," & decChange & _
                        "," & bIsDebit & _
                        ",'" & Trim(strTerms) & _
                        "'," & decAmountPaid & _
                        "," & lCostCentreID & _
                        "," & decTotalDiscountAmount & _
                        ",'" & strShippingAgentName & _
                        "','" & dtShippingDate & _
                        "'," & decTotalDiscountAmount & _
                            ")"

                .ConnectString = strOrgAccessConnString
                .ConnectToDatabase()

                bSaveSuccess = .ExecuteQuery(strOrgAccessConnString, _
                                    strSaveQuery, _
                                            datSaved)

                'CustomerInvoice
                strInsertInto = "INSERT INTO InvoiceCustomer (" & _
                                   "InvoiceNo," & _
                                           "CustomerNo" & _
                                                   ") VALUES "

                strSaveQuery = strInsertInto & _
                            "(" & _
                                "'" & strInvoiceNo & _
                                    "'," & lCustomerNo & _
                                                ")"

                bSaveSuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, _
                                  strSaveQuery, _
                                          datSaved)

                If bSaveSuccess = True Then
                    If DisplayMessages = True Then
                        MsgBox("Record Saved Successfully", _
                            MsgBoxStyle.Information, _
                                "iManagement - Invoice Details Saved")

                        .CommitTheTrans()
                    End If

                ElseIf bSaveSuccess = False Then

                    If DisplayMessages = True Then
                        MsgBox("'Save Invoice' action failed." & _
                            " Make sure all mandatory details are entered", _
                                MsgBoxStyle.Exclamation, _
                                    "iManagement - Save Invoice Details Failed")


                    End If
                    .RollbackTheTrans()

                End If

                .CloseDb()

                datSaved = Nothing
                objLogin = Nothing

            Catch ex As Exception

                'If Not objLogin Is Nothing Then
                '    .RollbackTheTrans()
                'End If

            End Try

        End With

    End Function

    Public Function Find(ByVal strQuery As String) As Boolean

        Try


            Dim datRetData As DataSet = New DataSet
            Dim bQuerySuccess As Boolean
            Dim myDataTables As DataTable
            Dim myDataColumns As DataColumn
            Dim myDataRows As DataRow
            Dim objLogin As IMLogin = New IMLogin

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bQuerySuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strQuery, _
                                                    datRetData)

            objLogin.CloseDb()

            If datRetData Is Nothing Then
                datRetData = Nothing
                objLogin = Nothing
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


                    For Each myDataRows In myDataTables.Rows

                        strInvoiceNo = myDataRows("InvoiceNo").ToString()
                        dtInvoiceDate = myDataRows("InvoiceDate")
                        dtInvoiceMaturityDate = myDataRows("InvoiceMaturityDate")
                        decInvoicedAmount = myDataRows("InvoicedAmount")
                        decVATAmount = myDataRows("VATAmount")
                        decGrandTotal = myDataRows("GrandTotal")
                        lDocumentID = myDataRows("DocumentID")
                        strClerkID = myDataRows("ClerkUserName").ToString()
                        lVATCodeID = myDataRows("VATCodeID")
                        strPINNo = myDataRows("PINNo").ToString()
                        dtInvoiceDueDate = myDataRows("InvoiceExpiryDate")
                        bInvoiceStatus = myDataRows("InvoiceStatus")

                        strCurrencyID = myDataRows("CurrencyID")
                        strSequenceGroupID = myDataRows("SequenceGroupID")
                        strInvoiceType = myDataRows("InvoiceType").ToString
                        decBalance = myDataRows("Balance")
                        strTransactionType = myDataRows("TransactionType")
                        decChange = myDataRows("Change")
                        bIsDebit = myDataRows("IsDebit")
                        decAmountPaid = myDataRows("AmountPaid")
                        decTotalDiscountAmount = myDataRows("TotalDiscountedAmount")
                        lCostCentreID = myDataRows("CostCentreID")
                        strShippingAgentName = myDataRows("ShippingAgentName").ToString
                        dtShippingDate = myDataRows("ShippingDate")


                    Next

                Next

                Return True
            Else
                Return False
            End If

            datRetData = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Sub Delete()
        Try

            'Deletes the country details of the country with the country code
            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin


            If Trim(strInvoiceNo) = "" _
                        Then
                MsgBox("Cannot Delete. Please select an existing Invoice.", _
                       MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete Information")

                datDelete = Nothing
                objLogin = Nothing
                Exit Sub

            End If

            strDeleteQuery = "DELETE * FROM InvoiceMaster WHERE " & _
            " InvoiceNo = '" & strInvoiceNo & "'"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strDeleteQuery, _
                            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Invoice Details Deleted")
            Else
                MsgBox("'Invoice delete' action failed", _
                    MsgBoxStyle.Exclamation, " Invoice Deletion failed")
            End If

            datDelete = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Update(ByVal strUpQuery As String)
        'Updates country details of the country with the country code

        Try


            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If Trim(strInvoiceNo) = "" _
                        Then

                MsgBox("Cannot Delete. Please select an existing Invoice.", _
                  MsgBoxStyle.Exclamation, _
                   "iManagement - invalid or incomplete Information")

                datUpdated = Nothing
                objLogin = Nothing
                Exit Sub

            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strUpdateQuery, _
                    datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", _
                    MsgBoxStyle.Information, _
                        "iManagement - Invoice Details Updated")

            Else
                MsgBox("Update of Invoice details failed", _
                    MsgBoxStyle.Information, _
                        "iManagement - Data update failed")

            End If

            datUpdated = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Sub

#End Region


End Class
