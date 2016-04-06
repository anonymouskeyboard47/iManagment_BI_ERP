Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMInvoiceTransactions

    Inherits IMInvoice

#Region "PrivateVariables"

    Private lTransactionID As Long
    Private lCostCentreID As Long
    Private lSequenceGroupID As Long
    Private lCOANr As Long
    Private lControlCOANr As Long
    Private dbCOAAmount As Decimal
    Private dbTransVATAmount As Decimal
    Private dbTransDiscount As Decimal

    Private lTransactionIDS() As Long
    Private lCostCentreIDS() As Long
    Private lSequenceGroupIDs() As Long
    Private lCOANrs() As Long
    Private lControlCOANrs() As Long
    Private dbCOAAmounts() As Decimal
    Private dbTransVATAmounts() As Decimal
    Private dbTransDiscounts() As Decimal 'Percentage per transaction

#End Region

#Region "Properties"

    Public Property TransactionID() As Long

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return lTransactionID
        End Get

        Set(ByVal Value As Long)
            lTransactionID = Value
        End Set

    End Property

    Public Property CostCentreID() As Long

        'USED TO SET AND RETRIEVE THE BANK NAME (STRING)
        Get
            Return lCostCentreID
        End Get

        Set(ByVal Value As Long)
            lCostCentreID = Value
        End Set

    End Property

    Public Shadows Property SequenceGroupID() As Long

        'USED TO SET AND RETRIEVE THE PHYSICAL ADDRESS (STRING)
        Get
            Return lSequenceGroupID
        End Get

        Set(ByVal Value As Long)
            lSequenceGroupID = Value
        End Set

    End Property

    Public Property COANr() As Long


        Get
            Return lCOANr
        End Get

        Set(ByVal Value As Long)
            lCOANr = Value
        End Set

    End Property

    Public Property COAAmount() As Decimal


        Get
            Return dbCOAAmount
        End Get

        Set(ByVal Value As Decimal)
            dbCOAAmount = Value
        End Set

    End Property

    Public Property ControlCOA() As Long
        Get
            Return lControlCOANr
        End Get

        Set(ByVal Value As Long)
            lControlCOANr = Value
        End Set

    End Property

    Public Property TransactionVATAmount() As Decimal
        Get
            Return dbTransVATAmount
        End Get

        Set(ByVal Value As Decimal)
            dbTransVATAmount = Value
        End Set

    End Property

    Public Property TransDiscount() As Decimal

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return dbTransDiscount
        End Get

        Set(ByVal Value As Decimal)
            dbTransDiscount = Value
        End Set

    End Property


#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Shadows Sub NewRecord()

        InvoiceNo = ""
        lTransactionID = 0
        lCostCentreID = 0
        lSequenceGroupID = 0
        lCOANr = 0
        dbCOAAmount = 0
        lControlCOANr = 0
        dbTransVATAmount = 0

    End Sub

    'Public Function ProcessInvoice() As Boolean
    '    Dim strRetProcArray As String() 'Array string of the delimited values
    '    Dim arLength As Long
    '    Dim lItem As Long
    '    Dim strItem As String
    '    Dim arPos As Long

    '    Try


    '    '''--------------------------------------------------------
    '    '''[Prepares the lTransactionIDS
    '    'strRetProcArray = strTransactionID.Split(",")

    '    'If strRetProcArray Is Nothing Then
    '    '    Return False
    '    '    Exit Function
    '    'Else

    '    '    arPos = 0

    '    '    If arLength <> strRetProcArray.GetLongLength(0) Then
    '    '        Return False
    '    '        Exit Function
    '    '    End If


    '    'ReDim lTransactionIDS(strRetProcArray.GetLongLength(0))

    '    '    For Each strItem In strRetProcArray
    '    '        lTransactionIDS(arPos) = CLng(strRetProcArray(strItem))
    '    '        arPos = arPos + 1
    '    '    Next
    '    'End If



    '    '--------------------------------------------------------
    '    '[
    '    strRetProcArray = Nothing
    '        strRetProcArray = strCostCentreID.Split(",")

    '    '[Used to preparea variable used for comparing variable sizes
    '    arLength = strRetProcArray.GetLongLength(0)

    '    If strRetProcArray Is Nothing Then
    '        Return False
    '        Exit Function
    '    Else

    '        arPos = 0

    '        If arLength <> strRetProcArray.GetLongLength(0) Then
    '            Return False
    '            Exit Function
    '        End If

    '        ReDim lCostCentreIDS(strRetProcArray.GetLongLength(0))

    '        For Each strItem In strRetProcArray
    '            lCostCentreIDS(arPos) = CLng(strRetProcArray(strItem))
    '        Next

    '    End If



    '    '--------------------------------------------------------
    '    '[
    '    strRetProcArray = Nothing
    '    strRetProcArray = strSequenceGroupID.Split(",")

    '    If strRetProcArray Is Nothing Then
    '        Return False
    '        Exit Function
    '    Else

    '        arPos = 0

    '        If arLength <> strRetProcArray.GetLongLength(0) Then
    '            Return False
    '            Exit Function
    '        End If

    '        ReDim strSequenceGroupIDs(strRetProcArray.GetLongLength(0))

    '        For Each strItem In strRetProcArray
    '            strSequenceGroupIDs(arPos) = strRetProcArray(strItem)
    '        Next
    '    End If



    '    '--------------------------------------------------------
    '    '[Chart of account number fo each transaction
    '    strRetProcArray = Nothing
    '    strRetProcArray = strCOANr.Split(",")

    '    If strRetProcArray Is Nothing Then
    '        Return False
    '        Exit Function
    '    Else

    '        arPos = 0

    '        If arLength <> strRetProcArray.GetLongLength(0) Then
    '            Return False
    '            Exit Function
    '        End If

    '        ReDim strCOANrs(strRetProcArray.GetLongLength(0))

    '        For Each strItem In strRetProcArray
    '            strCOANrs(arPos) = CLng(strRetProcArray(strItem))
    '        Next

    '    End If



    '    '--------------------------------------------------------
    '    '[Chart Of Account Amount for each transaction
    '    strRetProcArray = Nothing
    '    strRetProcArray = strCOAAmount.Split(",")

    '    If strRetProcArray Is Nothing Then
    '        Return False
    '        Exit Function
    '    Else

    '        arPos = 0

    '        If arLength <> strRetProcArray.GetLongLength(0) Then
    '            Return False
    '            Exit Function
    '        End If

    '        ReDim dbCOAAmounts(strRetProcArray.GetLongLength(0))


    '        For Each strItem In strRetProcArray
    '            dbCOAAmounts(arPos) = CDbl(strRetProcArray(strItem))
    '        Next
    '    End If



    '    '--------------------------------------------------------
    '    '[Chart of account number fo each transaction
    '    strRetProcArray = Nothing
    '    strRetProcArray = strControlCOANr.Split(",")

    '    If strRetProcArray Is Nothing Then
    '        Return False
    '        Exit Function
    '    Else

    '        arPos = 0

    '        If arLength <> strRetProcArray.GetLongLength(0) Then
    '            Return False
    '            Exit Function
    '        End If

    '        ReDim strControlCOANrs(strRetProcArray.GetLongLength(0))

    '        For Each strItem In strRetProcArray
    '            strControlCOANrs(arPos) = CLng(strRetProcArray(strItem))
    '        Next
    '    End If


    '    '--------------------------------------------------------
    '    '[VAT amount fo each transaction
    '    strRetProcArray = Nothing
    '    strRetProcArray = strTransVATAmount.Split(",")

    '    If strRetProcArray Is Nothing Then
    '        Return False
    '        Exit Function
    '    Else

    '        arPos = 0

    '        If arLength <> strRetProcArray.GetLongLength(0) Then
    '            Return False
    '            Exit Function
    '        End If

    '        ReDim dbTransVATAmounts(strRetProcArray.GetLongLength(0))

    '        For Each strItem In strRetProcArray
    '            dbTransVATAmounts(arPos) = CLng(strRetProcArray(strItem))
    '        Next
    '    End If


    '    '--------------------------------------------------------
    '    '[Discount amount fo each transaction
    '    strRetProcArray = Nothing
    '    strRetProcArray = strTransDiscount.Split(",")

    '    If strRetProcArray Is Nothing Then
    '        Return False
    '        Exit Function
    '    Else

    '        arPos = 0

    '        If arLength <> strRetProcArray.GetLongLength(0) Then
    '            Return False
    '            Exit Function
    '        End If

    '        ReDim dbTransDiscounts(strRetProcArray.GetLongLength(0))

    '        For Each strItem In strRetProcArray
    '            dbTransDiscounts(arPos) = CLng(strRetProcArray(strItem))
    '        Next
    '    End If



    '    Return True
    '        strRetProcArray = Nothing

    '    Catch ex As Exception
    '        MsgBox(ex.Message.ToString, MsgBoxStyle.Critical _
    '                         , "iManagement - Critical System Failure")
    '    End Try

    'End Function

    Public Function CalculateGrandTotal() As Boolean
        Try

        
        Dim dbItem As Double
        Dim dbInvAddition As Double

        CalculateTotalVAT()
        CalcTotalTransDiscounts()

        If dbCOAAmounts Is Nothing Then

            'No values in the array
            Return False

        Else

            For Each dbItem In dbCOAAmounts
                dbInvAddition = dbInvAddition + dbCOAAmounts(dbItem)
            Next

            'Deduct cumulative transaction discounts
            dbInvAddition = dbInvAddition - TotalTransactionDiscounts

            'Add VAT Amount
            GrandTotal = VATAmount + dbInvAddition

            'Deduct the Main discount amount from the grand total
            GrandTotal = GrandTotal - (GrandTotal * InvDiscountPercentage)

            'Sum of all discounts (Transaction Discounts + Main Invoice Discount)
            TotalDiscountAmount = (GrandTotal * InvDiscountPercentage) + _
                    TotalTransactionDiscounts

            Return True

            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical _
                 , "iManagement - Critical System Failure")
        End Try

    End Function

    Public Function CalculateTotalVAT() As Boolean
        Try
            Dim dbItem As Double
            Dim dbVATAddition As Double

            If dbTransVATAmounts Is Nothing Then

                'No values in the array
                Return False

            Else
                For Each dbItem In dbTransVATAmounts
                    dbVATAddition = dbVATAddition + dbTransVATAmounts(dbItem)
                Next

                VATAmount = dbVATAddition

                Return True

        End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical _
                             , "iManagement - Critical System Failure")
        End Try

    End Function

    Public Function CalculateBalance() As Boolean

        Try
            If AmountPaid < GrandTotal Then

                Balance = AmountPaid - GrandTotal
                Return False

            Else
                Balance = AmountPaid - GrandTotal
                Return True

            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical _
                             , "iManagement - Critical System Failure")
        End Try

    End Function

    Public Function CalcTotalTransDiscounts() As Boolean

        Try

            Dim dbItem As Double
            Dim dbTransAdditions As Double
            Dim arPos As Long

            If dbTransDiscounts Is Nothing Then

                'No values in the array
                Return False

            Else

                For Each dbItem In dbCOAAmounts

                    dbTransAdditions = dbTransAdditions + (dbCOAAmounts(dbItem) * dbTransDiscounts(arPos))
                    arPos = arPos + 1

                Next

                TotalTransactionDiscounts = dbTransAdditions

                Return True

            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical _
                             , "iManagement - Critical System Failure")
        End Try

    End Function

#End Region

#Region "DatabaseProcedures"

    Public Overloads Sub Save(ByVal DisplayMessages As Boolean, _
    ByVal strInvNo As String)

        Dim myTrans As OleDbTransaction
        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String
        Dim dbItem As Double
        Dim arPos As Long


        Try

            'Begin transaction (Save Invoice, Save Invoice Customer,
            'Save Each Transaction)
            myTrans.Begin(IsolationLevel.ReadCommitted)

            'Save the invoice details (Which also saves custmer invoice)
            Save(True)

            If Trim(InvoiceNo) <> "" _
                                Then

                For Each dbItem In dbCOAAmounts


                    strInsertInto = "INSERT INTO InvoiceTransactions (" & _
                            "InvoiceNo," & _
                            "TransactionID," & _
                            "CostCentreID," & _
                            "COANr," & _
                            "SequenceGroupID," & _
                            "COAAmount," & _
                            "TransDiscount," & _
                            "TransVATAmount" & _
                            ") VALUES "


                    strSaveQuery = strInsertInto & _
                                "(" & _
                            "'" & InvoiceNo & _
                            "'," & lTransactionIDS(arPos) & _
                            "," & lCostCentreIDS(arPos) & _
                            "," & lCOANrs(arPos) & _
                            "," & lSequenceGroupIDs(arPos) & _
                            "," & dbCOAAmounts(arPos) & _
                            "," & dbTransDiscounts(arPos) & _
                            "," & dbTransVATAmounts(arPos) & _
                                ")"

                    objLogin.connectString = strAccessConnString
                    objLogin.ConnectToDatabase()

                    bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                        strSaveQuery, _
                                                datSaved)

                Next


                objLogin.CommitTheTrans()

                myTrans.Commit()

                objLogin.CloseDb()


                If bSaveSuccess = True Then
                    MsgBox("Record Saved Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement - Invoice Details Saved")

                Else

                    MsgBox("'Save Invoice' action failed." & _
                        " Make sure all mandatory details are entered", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Invoice Details Failed")
                    myTrans.Rollback()
                    objLogin.RollbackTheTrans()
                End If

            Else
                MsgBox("Cannot save. Missing information", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement -Missing Information")

                myTrans.Rollback()
                objLogin.RollbackTheTrans()
            End If


        Catch ex As Exception
            MsgBox("Error during save operation. Must rollback to " & _
                                    "original state", _
                MsgBoxStyle.Critical, "iManagement - System Failure")

            myTrans.Rollback()
            objLogin.RollbackTheTrans()

        End Try

    End Sub

    Public Overloads Function Find(ByVal strQuery As String, _
                            ByVal strInvNo As String) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.connectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bQuerySuccess = objLogin.ExecuteQuery(strAccessConnString, strQuery, _
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


                For Each myDataRows In myDataTables.Rows

                    InvoiceNo = myDataRows("InvoiceNo").ToString()
                    InvoiceDate = myDataRows("InvoiceDate")
                    InvoiceMaturityDate = myDataRows("InvoiceMaturityDate")
                    InvoicedAmount = myDataRows("InvoicedAmount")
                    VATAmount = myDataRows("VATAmount")
                    GrandTotal = myDataRows("GrandTotal")
                    DocumentID = myDataRows("DocumentID")
                    ClerkID = myDataRows("ClerkID")
                    VATCodeID = myDataRows("VATCodeID")
                    PINNo = myDataRows("PINNo").ToString()
                    InvoiceExpiryDate = myDataRows("InvoiceExpiryDate")
                    InvoiceStatus = myDataRows("InvoiceStatus")

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Overloads Sub Delete(ByVal strDelQuery As String, _
                            ByVal strInvNo As String)
        'Deletes the country details of the country with the country code
        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strDeleteQuery = strDelQuery

        If Trim(InvoiceNo) <> "" _
                    Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, _
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
        Else
            MsgBox("Cannot Delete. Please select an existing Invoice", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

        End If

    End Sub

    Public Overloads Sub Update(ByVal strUpQuery As String, _
                            ByVal strInvNo As String)
        'Updates country details of the country with the country code

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If Trim(InvoiceNo) <> "" _
                    Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                strUpdateQuery, _
                    datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then

                MsgBox("Record Updated Successfully", _
                    MsgBoxStyle.Information, _
                        "iManagement - Invoice Details Updated")

            Else

                MsgBox("Update of employer details failed", MsgBoxStyle.Information, _
                    "iManagement - Data update failed")

            End If

        End If

    End Sub

#End Region

End Class
