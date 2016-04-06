Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections


Public Class IMProductOrders
    Inherits IMSupplierOrder

#Region "PrivateVariables"

    Private lProductID As Long
    Private lOrderID As Long
    Private lStockManagementID As Long
    Private strOrderType As String
    Private dbQuantityOrdered As Double
    Private dtDateOfOrder As Date
    Private lRequestUserID As Long
    Private lOrderUserID As Long
    Private lApprovalOfficerUserID As Long
    Private dtDateRegistered As Date
    Private strOrderSerialNo As String
    Private strOrderSummaryDetails As String
    Private bOrderSentToSupplier As Boolean
    Private dtDateOrderSentToSupplier As Date
    Private bOrderStatus As Boolean
    Private strOrderCumulativeSummary As String
    Private lCostCentreID As Long

#End Region

#Region "Properties"

    Public Property ApprovalOfficerUserID() As Long

        Get
            Return lApprovalOfficerUserID
        End Get

        Set(ByVal Value As Long)
            lApprovalOfficerUserID = Value
        End Set

    End Property

    Public Property CostCentreID() As Long

        Get
            Return lCostCentreID
        End Get

        Set(ByVal Value As Long)
            lCostCentreID = Value
        End Set

    End Property

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

    Public Property QuantityOrdered() As Double

        Get
            Return dbQuantityOrdered
        End Get

        Set(ByVal Value As Double)
            dbQuantityOrdered = Value
        End Set

    End Property

    Public Property OrderStatus() As Boolean

        Get
            Return bOrderStatus
        End Get

        Set(ByVal Value As Boolean)
            bOrderStatus = Value
        End Set

    End Property

    Public Property DateOfOrder() As Date

        Get
            Return dtDateOfOrder
        End Get

        Set(ByVal Value As Date)
            dtDateOfOrder = Value
        End Set

    End Property

    Public Property RequestUserID() As Long

        Get
            Return lRequestUserID
        End Get

        Set(ByVal Value As Long)
            lRequestUserID = Value
        End Set

    End Property

    Public Property OrderUserID() As Long

        Get
            Return lOrderUserID
        End Get

        Set(ByVal Value As Long)
            lOrderUserID = Value
        End Set

    End Property

    Public Property DateRegistered() As Date

        Get
            Return dtDateRegistered
        End Get

        Set(ByVal Value As Date)
            dtDateRegistered = Value
        End Set

    End Property

    Public Property DateOrderSentToSupplier() As Date

        Get
            Return dtDateOrderSentToSupplier
        End Get

        Set(ByVal Value As Date)
            dtDateOrderSentToSupplier = Value
        End Set

    End Property

    Public Property OrderSerialNo() As String

        Get
            Return strOrderSerialNo
        End Get

        Set(ByVal Value As String)
            strOrderSerialNo = Value
        End Set

    End Property

    Public Property OrderSummaryDetails() As String

        Get
            Return strOrderSummaryDetails
        End Get

        Set(ByVal Value As String)
            strOrderSummaryDetails = Value
        End Set

    End Property

    Public Property OrderType() As String

        Get
            Return strOrderType
        End Get

        Set(ByVal Value As String)
            strOrderType = Value
        End Set

    End Property

    Public Property OrderSentToSupplier() As Boolean

        Get
            Return bOrderSentToSupplier
        End Get

        Set(ByVal Value As Boolean)
            bOrderSentToSupplier = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function ReturnOrderObjectFromSerialNo _
        (ByVal strValOrderSerialNo As String) As Object

        Try

            FindOrder("SELECT * FROM ProductOrders WHERE " & _
                "OrderSerialNo = '" & strValOrderSerialNo & "'", True)

            Return Me

        Catch ex As Exception

        End Try

    End Function

    Public Function LoadReport(ByVal strOrderSerialNo As String) As Object

        Try

            Dim objData As DataSet = New DataSet
            Dim objLogin As IMLogin = New IMLogin


        Catch ex As Exception

        End Try

    End Function

    Public Function CalculateNextOrderSerialNo() As String

        Try

            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String
            Dim strProposedSerialNo As String

            Dim objLogin As IMLogin = New IMLogin

            With objLogin
                MyMaxValue = .FillArray(strOrgAccessConnString, _
                                "SELECT COUNT(*) AS TotalRecords FROM" & _
                                    " ProductOrders WHERE " & _
                                    " Day(ProductOrders.DateRegistered) " & _
                                    " = Day(Now())  AND " & _
                                    " Month(ProductOrders.DateRegistered) " & _
                                    " = Month(Now()) AND " & _
                                    " Year(ProductOrders.DateRegistered) " & _
                                    " = Year(Now())", "", "")
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

            strProposedSerialNo = "Order" & Now.Day.ToString _
                & Now.Month.ToString & _
                    Now.Year.ToString & _
                            MaxValue.ToString

            Return strProposedSerialNo

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, _
                "iManagement - System Error")
        End Try

    End Function

    Public Function ReturnAllSerialNumbers _
       (Optional ByVal strValProductName As String = "", _
            Optional ByVal lValProductID As Long = 0, _
                Optional ByVal bIncludeOrdersReceived As Boolean = True, _
                    Optional ByVal bIncludeOrdersNotApproved As Boolean = True, _
                        Optional ByVal bIncludeOrdersNotSentToSupplier _
                            As Boolean = True, Optional ByVal _
                                    bIncludeOrderNotConfirmedBySupplier _
                                        As Boolean = True) As String()

        Dim strQueryToUse As String
        Dim objLogin As IMLogin = New IMLogin
        Dim arItems() As String

        Try


            '==========================================
            'Indefinite
            If lValProductID <> 0 And bIncludeOrdersReceived = True And _
                bIncludeOrdersNotApproved = True And _
                    bIncludeOrdersNotSentToSupplier = True And _
                        bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders" & _
                    " WHERE ProductID = " & lValProductID

            ElseIf Trim(strValProductName) <> "" And _
                bIncludeOrdersReceived = True And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                    " INNER JOIN ProductMaster ON " & _
                        " ProductMaster.ProductID = ProductOrders.ProductID" & _
                            " WHERE ProductName = '" & strValProductName & "'"

            ElseIf lValProductID = 0 And strValProductName = "" And _
                bIncludeOrdersReceived = True And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders "

            End If


            '==========================================
            'Dont include Orders Received, regardless of approval, _
            'regardless of sending to supplier, regardless of being confirmed
            If lValProductID <> 0 And bIncludeOrdersReceived = False And _
                bIncludeOrdersNotApproved = True And _
                    bIncludeOrdersNotSentToSupplier = True And _
                        bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                    " INNER JOIN ProductDeliveries ON " & _
                    " ProductDeliveries.OrderID = " & _
                    " ProductOrders.OrderID " & _
                    " WHERE ProductID = " & lValProductID

            ElseIf Trim(strValProductName) <> "" And _
                bIncludeOrdersReceived = False And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM (ProductOrders " & _
                    " INNER JOIN ProductMaster ON " & _
                        " ProductMaster.ProductID = ProductOrders.ProductID)" & _
                        " INNER JOIN ProductDeliveries ON " & _
                        " ProductDeliveries.OrderID = ProductOrders.OrderID" & _
                            " WHERE ProductName = '" & strValProductName & "'"

            ElseIf lValProductID = 0 And strValProductName = "" And _
                bIncludeOrdersReceived = False And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                        " INNER JOIN ProductDeliveries ON " & _
                        " ProductOrders.OrderID = ProductDeliveries.OrderID "
            End If



            '==========================================
            'Dont include Orders Received, Approved and Sent
            If lValProductID <> 0 And bIncludeOrdersReceived = False And _
                bIncludeOrdersNotApproved = False And _
                    bIncludeOrdersNotSentToSupplier = False And _
                        bIncludeOrderNotConfirmedBySupplier = False Then

                strQueryToUse = "SELECT OrderSerialNo FROM (ProductOrders " & _
                    " INNER JOIN ProductMaster ON " & _
                    " ProductMaster.ProductID = ProductOrders.ProductID)" & _
                    " INNER JOIN ProductDeliveries ON " & _
                    " ProductDeliveries.OrderID = ProductOrders.OrderID" & _
                    " WHERE ProductID = " & lValProductID & _
                    " AND OrderSentToSupplier = TRUE " & _
                    " AND ProductOrders.ApprovalOfficerUserID <> 0 "

            ElseIf Trim(strValProductName) <> "" And _
                bIncludeOrdersReceived = False And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM (ProductOrders " & _
                    " INNER JOIN ProductMaster ON " & _
                        " ProductMaster.ProductID = ProductOrders.ProductID)" & _
                        " INNER JOIN ProductDeliveries ON " & _
                        " ProductDeliveries.OrderID = ProductOrders.OrderID" & _
                            " WHERE ProductName = '" & strValProductName & "'" & _
                    " AND OrderSentToSupplier = TRUE" & _
                    " AND ProductOrders.ApprovalOfficerUserID <> 0 "

            ElseIf lValProductID = 0 And strValProductName = "" And _
                bIncludeOrdersReceived = False And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM (ProductOrders " & _
                    " INNER JOIN ProductMaster ON " & _
                        " ProductMaster.ProductID = ProductOrders.ProductID)" & _
                        " INNER JOIN ProductDeliveries ON " & _
                        " ProductDeliveries.OrderID = ProductOrders.OrderID"

            End If


            '+++++++++++++++++++++++++++++++++++++++++++
            'DONT bIncludeOrdersNotApproved, regardless of reception, 
            'regardless of  sent, regardless of confirmed 
            If lValProductID <> 0 And bIncludeOrdersReceived = False And _
                bIncludeOrdersNotApproved = False And _
                    bIncludeOrdersNotSentToSupplier = True And _
                        bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                    " WHERE ProductID = " & lValProductID & _
                    " AND ProductOrders.ApprovalOfficerUserID <> 0 "

            ElseIf Trim(strValProductName) <> "" And _
                bIncludeOrdersReceived = True And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                    " INNER JOIN ProductMaster ON " & _
                        " ProductMaster.ProductID = ProductOrders.ProductID" & _
                            " WHERE ProductName = '" & strValProductName & _
                                "' AND ProductOrders.ApprovalOfficerUserID <> 0"

            ElseIf lValProductID = 0 And strValProductName = "" And _
                bIncludeOrdersReceived = True And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                " WHERE ProductOrders.ApprovalOfficerUserID <> 0"

            End If



            '+++++++++++++++++++++++++++++++++++++++++++
            ' Dont bIncludeOrdersNotApproved AND IncludeNotSentToSupplier regardless _
            ' of the rest
            If lValProductID <> 0 And bIncludeOrdersReceived = True And _
                bIncludeOrdersNotApproved = True And _
                    bIncludeOrdersNotSentToSupplier = True And _
                        bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                    " WHERE ProductID = " & lValProductID & _
                    " AND ProductOrders.ApprovalOfficerUserID <> 0 " & _
                    " AND OrderSentToSupplier = TRUE"

            ElseIf Trim(strValProductName) <> "" And _
                bIncludeOrdersReceived = True And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                    " INNER JOIN ProductMaster ON " & _
                        " ProductMaster.ProductID = ProductOrders.ProductID" & _
                            " WHERE ProductName = '" & strValProductName & _
                                "' AND ProductOrders.ApprovalOfficerUserID <> 0" & _
                    " AND OrderSentToSupplier = TRUE"

            ElseIf lValProductID = 0 And strValProductName = "" And _
                bIncludeOrdersReceived = True And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                " WHERE ProductOrders.ApprovalOfficerUserID <> 0" & _
                    " AND OrderSentToSupplier = TRUE"
            End If



            '+++++++++++++++++++++++++++++++++++++++++++
            ' bIncludeOrdersApproved AND Dont IncludeNotSentToSupplier,
            ' Dont IncludeOrderNotConfirmed regardless of the rest
            If lValProductID <> 0 And bIncludeOrdersReceived = True And _
                bIncludeOrdersNotApproved = True And _
                    bIncludeOrdersNotSentToSupplier = True And _
                        bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                    " WHERE ProductID = " & lValProductID & _
                    " AND ProductOrders.ApprovalOfficerUserID <> 0 " & _
                    " AND OrderSentToSupplier = TRUE"

            ElseIf Trim(strValProductName) <> "" And _
                bIncludeOrdersReceived = True And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                    " INNER JOIN ProductMaster ON " & _
                        " ProductMaster.ProductID = ProductOrders.ProductID" & _
                            " WHERE ProductName = '" & strValProductName & _
                                "' AND ProductOrders.ApprovalOfficerUserID <> 0" & _
                    " AND OrderSentToSupplier = TRUE"

            ElseIf lValProductID = 0 And strValProductName = "" And _
                bIncludeOrdersReceived = True And _
                    bIncludeOrdersNotApproved = True And _
                        bIncludeOrdersNotSentToSupplier = True And _
                            bIncludeOrderNotConfirmedBySupplier = True Then

                strQueryToUse = "SELECT OrderSerialNo FROM ProductOrders " & _
                " WHERE ProductOrders.ApprovalOfficerUserID <> 0" & _
                    " AND OrderSentToSupplier = TRUE"
            End If



            With objLogin
                arItems = .FillArray _
                    (strOrgAccessConnString, strQueryToUse, "", "")

            End With

            objLogin = Nothing

            Return arItems

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnDefaultOrderTypes() As String()

        Try

            Dim arItems() As String

            arItems = Split("Normal Order,Recorded Transaction Schedule,Recorded Delivery Schedule", ",")

            Return arItems


        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnDefaultOrderSentStatus() As String()

        Try

            Dim arItems() As String

            arItems = Split("Enabled,Disabled", ",")

            Return arItems

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnOrderIDFromSerialNo _
      (ByVal strValOrderSerialNo As String) As Object

        Try
            Dim arItem() As String
            Dim objLogin As IMLogin = New IMLogin

            arItem = objLogin.FillArray(strOrgAccessConnString, _
                "SELECT OrderID FROM ProductOrders WHERE " & _
                "OrderSerialNo = '" & strValOrderSerialNo & "'", _
                "", "")

            If Not arItem Is Nothing Then

            End If

            Return arItem(0)

        Catch ex As Exception

        End Try

    End Function

    Public Function ApproveOrder() As Boolean

        Try


        Catch ex As Exception

        End Try

    End Function

    Public Function MarkOrderAsConfirmed() As Boolean

        Try


        Catch ex As Exception

        End Try

    End Function

    Public Function SendOrderToSupplier(ByVal bEncryptOrder As Boolean) As Boolean

    End Function

    Public Function DisableOrder() As Boolean


    End Function

#End Region

#Region "DatabaseProcedures"

    'Save informaiton
    Public Function SaveOrder(ByVal DisplayErrorMessages As Boolean, _
            ByVal DisplaySuccessMessages As Boolean, _
                ByVal DisplayFailureMessages As Boolean) As Boolean

        'Saves a new base organization
        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String
            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String

            If Trim(strOrganizationName) = "" Then

                MsgBox("Please open an existing company.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Select an existing company")
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            If lProductID = 0 Or lStockManagementID = 0 Or _
                dbQuantityOrdered = 0 Or SupplierOrgID = 0 Then

                MsgBox("You must provide an available Product, Qauntity to " & _
                    Chr(10) & "Order, the Supplier, and the Supplier's default sale price." _
                                , MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            'Check if there is an existing Order with this series
            If FindOrder("SELECT * FROM ProductOrders WHERE OrderID = " & _
                OrderID & " OR OrderSerialNo = '" & strOrderSerialNo & "'", _
                False) = True Then

                If MsgBox("The Product Orders Details already exists." & _
                Chr(10) & "Do you want to update the details?", _
                    MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                            MsgBoxResult.Yes Then

                    'If the order has been sent and confirmed
                    If FindOrder("SELECT * FROM ProductOrders " & _
                    " INNER JOIN SupplierOrder ON " & _
                    " ProductOrders.OrderID = " & _
                    " SupplierOrder.OrderID " & _
                    " WHERE OrderID = " & OrderID & _
                    " OR OrderSerialNo = '" & strOrderSerialNo & _
                    "' WHERE OrderConfirmed = TRUE", _
                    False) = True Then

                        ReturnError += "This order has already been confirmed." & _
                        " This means that it cannot be altered any further."

                        Exit Function

                        objLogin = Nothing
                        datSaved = Nothing

                        Exit Function
                    End If


                    'If the order has been created and approved and sent 
                    'but not confirmed
                    If FindOrder("SELECT * FROM ProductOrders " & _
                    " INNER JOIN SupplierOrder ON " & _
                    " ProductOrders.OrderID = " & _
                    " SupplierOrder.OrderID " & _
                    " WHERE OrderID = " & OrderID & _
                    " OR OrderSerialNo = '" & strOrderSerialNo & _
                    "' WHERE OrderConfirmed = False AND " & _
                    " OrderSentToSupplier = TRUE", _
                    False) = True Then

                        Save(False, False, False, False)

                        objLogin = Nothing
                        datSaved = Nothing

                        Exit Function
                    End If


                    'If the order has been created and approved 
                    'but not sent
                    If FindOrder("SELECT * FROM ProductOrders " & _
                    " INNER JOIN SupplierOrder ON " & _
                    " ProductOrders.OrderID = " & _
                    " SupplierOrder.OrderID " & _
                    " WHERE OrderID = " & OrderID & _
                    " OR OrderSerialNo = '" & strOrderSerialNo & _
                    "' AND lApprovalOfficerUserID <> 0 AND " & _
                    " OrderSentToSupplier = FALSE", _
                     False) = True Then

                        UpdateOrder("UPDATE ProductOrders SET " & _
                   "   ProductID = " & lProductID & _
                   " , StockManagementID = " & lStockManagementID & _
                   " , QuantityOrdered = " & dbQuantityOrdered & _
                   " , OrderStatus = " & bOrderStatus & _
                   " , DateOrderSentToSupplier = #" & dtDateOrderSentToSupplier & _
                   "# , bOrderSentToSupplier = " & bOrderSentToSupplier & _
                   ", OrderSummaryDetails = '" & strOrderSummaryDetails & _
                   "', OrderType = '" & strOrderType & _
                   "'  WHERE  OrderID = " & OrderID & _
                   "   OR OrderSerialNo = '" & strOrderSerialNo & "'")

                        Save(False, False, False, False)

                        objLogin = Nothing
                        datSaved = Nothing

                        Exit Function
                    End If


                    'If the Order has been created but not approved
                    UpdateOrder("UPDATE ProductOrders SET " & _
                    " ProductID = " & lProductID & _
                    " , StockManagementID = " & lStockManagementID & _
                    " , QuantityOrdered = " & dbQuantityOrdered & _
                    " , OrderStatus = " & bOrderStatus & _
                    " , DateOfOrder = #" & dtDateOfOrder & _
                    "#, ApprovalOfficerUserID = " & lApprovalOfficerUserID & _
                    " , DateOrderSentToSupplier = #" & dtDateOrderSentToSupplier & _
                    "# , bOrderSentToSupplier = " & bOrderSentToSupplier & _
                    ", OrderSummaryDetails = '" & strOrderSummaryDetails & _
                    "', OrderType = '" & strOrderType & _
                    "', CostCentreID = " & lCostCentreID & _
                    " WHERE  OrderID = " & OrderID & _
                    " OR OrderSerialNo = '" & strOrderSerialNo & "'")

                    Save(False, False, False, False)

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            strOrderSerialNo = CalculateNextOrderSerialNo()

            strInsertInto = "INSERT INTO ProductOrders (" & _
                "ProductID," & _
                "StockManagementID," & _
                "QuantityOrdered," & _
                "OrderStatus," & _
                "DateOfOrder," & _
                "RequestUserID," & _
                "ApprovalOfficerUserID," & _
                "DateOrderSentToSupplier," & _
                "OrderSerialNo," & _
                "OrderSummaryDetails," & _
                "OrderType," & _
                "OrderSentToSupplier," & _
                "CostCentreID" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lProductID & _
                    "," & lStockManagementID & _
                    "," & dbQuantityOrdered & _
                    "," & bOrderStatus & _
                    ",#" & dtDateOfOrder & _
                    "#," & lRequestUserID & _
                    "," & lApprovalOfficerUserID & _
                    ",#" & dtDateOrderSentToSupplier & _
                    "#,'" & strOrderSerialNo & _
                    "','" & Trim(strOrderSummaryDetails) & _
                    "','" & Trim(strOrderType) & _
                    "'," & bOrderSentToSupplier & _
                    "," & lCostCentreID & _
                    ")"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strSaveQuery, datSaved)

            objLogin.CloseDb()

            OrderID = ReturnOrderIDFromSerialNo(strOrderSerialNo)
            OrderID = ReturnOrderIDFromSerialNo(strOrderSerialNo)

            Save(False, False, False, False)


            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    ReturnError += "New Product Orders record saved " & _
                        "successfully"

                End If
            Else

                If DisplayFailureMessages = True Then
                    returnerror += "'Save New Product Orders Details' action failed." & _
                        " Make sure all mandatory details are entered."
                End If
            End If

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception
            If DisplayErrorMessages = True Then
                ReturnError += "Management - Database or system error " & _
                    ex.Message
            End If

        End Try

    End Function

    'Find Informaiton
    Public Function FindOrder(ByVal strQuery As String, _
        ByVal bReturnValues As Boolean) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bQuerySuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strQuery, datRetData)

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
                    objLogin = Nothing
                    datRetData = Nothing

                    Exit Function
                End If


                If bReturnValues = True Then
                    For Each myDataRows In myDataTables.Rows

                        If IsDBNull(myDataRows("ProductID")) = False Then
                            lProductID = _
                                    myDataRows("ProductID")
                        End If

                        If IsDBNull(myDataRows("StockManagementID")) = False Then
                            lStockManagementID = _
                                    myDataRows("StockManagementID")
                        End If

                        If IsDBNull(myDataRows("QuantityOrdered")) = False Then
                            dbQuantityOrdered = _
                                    myDataRows("QuantityOrdered")
                        End If

                        If IsDBNull(myDataRows("OrderStatus")) = False Then
                            bOrderStatus = _
                                myDataRows("OrderStatus")
                        End If

                        If IsDBNull(myDataRows("DateOfOrder")) = False Then
                            dtDateOfOrder = _
                                    myDataRows("DateOfOrder")
                        End If

                        If IsDBNull(myDataRows("RequestUserID")) = False Then
                            lRequestUserID = _
                                    myDataRows("RequestUserID")
                        End If

                        If IsDBNull(myDataRows("ApprovalOfficerUserID")) = False Then
                            lApprovalOfficerUserID = _
                                    myDataRows("ApprovalOfficerUserID")
                        End If


                        If IsDBNull(myDataRows("DateRegistered")) = False Then
                            dtDateRegistered = _
                                    myDataRows("DateRegistered")
                        End If

                        If IsDBNull(myDataRows("DateOrderSentToSupplier")) = False Then
                            dtDateOrderSentToSupplier = _
                                    myDataRows("DateOrderSentToSupplier")
                        End If

                        If IsDBNull(myDataRows("OrderSerialNo")) = False Then
                            strOrderSerialNo = _
                                    myDataRows("OrderSerialNo")
                        End If

                        If IsDBNull(myDataRows("OrderSummaryDetails")) = False Then
                            strOrderSummaryDetails = _
                                    myDataRows("OrderSummaryDetails")
                        End If


                        If IsDBNull(myDataRows("OrderType")) = False Then
                            strOrderType = _
                                    myDataRows("OrderType")
                        End If

                        If IsDBNull(myDataRows("OrderSentToSupplier")) = False Then
                            bOrderSentToSupplier = _
                                    myDataRows("OrderSentToSupplier")
                        End If

                        If IsDBNull(myDataRows("OrderID")) = False Then
                            lOrderID = _
                                    myDataRows("OrderID")
                        End If

                        If IsDBNull(myDataRows("CostCentreID")) = False Then
                            lCostCentreID = _
                                    myDataRows("CostCentreID")
                        End If

                        Find("SELECT * FROM SupplierOrder WHERE OrderID = " & _
                            OrderID, True)

                    Next
                End If
            Next

            objLogin = Nothing
            datRetData = Nothing

            Return True
        Else

            objLogin = Nothing
            datRetData = Nothing
            Return False

        End If

        objLogin = Nothing
        datRetData = Nothing

    End Function

    'Delete data
    Public Sub DeleteOrder()

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try

            If lOrderID = 0 And Trim(strOrderSerialNo) = "" Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                strDeleteQuery = "DELETE * FROM ProductOrders " & _
                    "WHERE OrderID = " & lOrderID

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strDeleteQuery, datDelete)

                strDeleteQuery = "DELETE * FROM SupplierOrder " & _
                   "WHERE OrderID = " & lOrderID

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strDeleteQuery, datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement - Product Orders Details Deleted")
                Else
                    MsgBox("'Product Orders Details delete' action failed", _
                            MsgBoxStyle.Exclamation, _
                                "Product Orders Details Deletion failed")
                End If
            Else

                MsgBox("Cannot Delete. Please select an existing Product Orders.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - invalid or incomplete information")

            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Sub

    Public Sub UpdateOrder(ByVal strUpQuery As String)

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)
            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                returnerror+= "Product Orders record updated successfully"
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub

#End Region

End Class



