Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMProducts
    Inherits IMProductStockAmount

#Region "Connection Properties"

    Public Property TheReturnError() As String

        Get
            Return ReturnError
        End Get

        Set(ByVal Value As String)
            ReturnError = Value
        End Set


    End Property

    Public Property OrganizationName() As String

        Get
            Return strOrganizationName
        End Get

        Set(ByVal Value As String)
            strOrganizationName = Value
        End Set

    End Property


    Public Property AccessConnectionString() As String

        Get
            Return strAccessConnString
        End Get

        Set(ByVal Value As String)
            strAccessConnString = Value
        End Set

    End Property

    Public Property OrgConnectionString() As String

        Get
            Return strOrgAccessConnString
        End Get

        Set(ByVal Value As String)
            strOrgAccessConnString = Value
        End Set

    End Property

    Public Property OrgConnectionSQLServer() As String

        Get
            Return bOrgConnectionSQLServer
        End Get

        Set(ByVal Value As String)
            bOrgConnectionSQLServer = Value
        End Set

    End Property

    Public Property AccessConnStringADOX() As String

        Get
            Return strAccessConnStringADOX
        End Get

        Set(ByVal Value As String)
            strAccessConnStringADOX = Value
        End Set

    End Property

    Public Property OrgAccessConnStringADOX() As Boolean

        Get
            Return strOrgAccessConnStringADOX
        End Get

        Set(ByVal Value As Boolean)
            strOrgAccessConnStringADOX = Value
        End Set

    End Property

    Public Property SQLConnString() As String

        Get
            Return strSQLConnString
        End Get

        Set(ByVal Value As String)
            strSQLConnString = Value
        End Set

    End Property

    Public Property DBUserName() As String

        Get
            Return strDBUserName
        End Get

        Set(ByVal Value As String)
            strDBUserName = Value
        End Set

    End Property

    Public Property DBPassword() As String

        Get
            Return strDBPassword
        End Get

        Set(ByVal Value As String)
            strDBPassword = Value
        End Set

    End Property

    Public Property DBDatabase() As String

        Get
            Return strDBDatabase
        End Get

        Set(ByVal Value As String)
            strDBDatabase = Value
        End Set

    End Property

    Public Property DBDBPath() As String

        Get
            Return strDBDBPath
        End Get

        Set(ByVal Value As String)
            strDBDBPath = Value
        End Set

    End Property

#End Region

#Region "PrivateVariables"

    Private lProductID As Long
    Private strProductName As String
    Private strProductTextCode As String
    Private strProductType As String
    Private bProductStatus As Boolean
    Private lDefaultSupplierID As Long
    Private lDefaultPriceID As Long
    Private lCOABalanceSheet As Long
    Private lCOAExportID As Long
    Private lCOAProfitLoss As Long
    Private strUnitName As String
    Private lChargedPerUnit As Long

    Private lProductSize As Double
    Private strProductColor As String
    Private strProductModelNoSerialNo As String
    Private strProductModelName As String
    Private strProductManufacturer As String

#End Region

#Region "Properties"

    Public Property ProductSize() As Long

        Get
            Return lProductSize
        End Get

        Set(ByVal Value As Long)
            lProductSize = Value
        End Set

    End Property

    Public Property ProductColor() As String

        Get
            Return strProductColor
        End Get

        Set(ByVal Value As String)
            strProductColor = Value
        End Set

    End Property

    Public Property ProductModelNoSerialNo() As String

        Get
            Return strProductModelNoSerialNo
        End Get

        Set(ByVal Value As String)
            strProductModelNoSerialNo = Value
        End Set

    End Property

    Public Property ProductModelName() As String

        Get
            Return strProductModelName
        End Get

        Set(ByVal Value As String)
            strProductModelName = Value
        End Set

    End Property

    Public Property ProductManufacturer() As String

        Get
            Return strProductManufacturer
        End Get

        Set(ByVal Value As String)
            strProductManufacturer = Value
        End Set

    End Property

    Public Property UnitName() As String

        Get
            Return strUnitName
        End Get

        Set(ByVal Value As String)
            strUnitName = Value
        End Set

    End Property

    Public Property ChargedPerUnit() As String

        Get
            Return lChargedPerUnit
        End Get

        Set(ByVal Value As String)
            lChargedPerUnit = Value
        End Set

    End Property

    Public Property ProductType() As String

        Get
            Return strProductType
        End Get

        Set(ByVal Value As String)
            strProductType = Value
        End Set

    End Property

    Public Property ProductName() As String

        Get
            Return strProductName
        End Get

        Set(ByVal Value As String)
            strProductName = Value
        End Set

    End Property

    Public Property ProductTextCode() As String

        Get
            Return strProductTextCode
        End Get

        Set(ByVal Value As String)
            strProductTextCode = Value
        End Set

    End Property

    Public Property ProductStatus() As Boolean

        Get
            Return bProductStatus
        End Get

        Set(ByVal Value As Boolean)
            bProductStatus = Value
        End Set

    End Property

    Public Property DefaultSupplierID() As Long

        Get
            Return lDefaultSupplierID
        End Get

        Set(ByVal Value As Long)
            lDefaultSupplierID = Value
        End Set

    End Property

    Public Property DefaultPriceID() As Long

        Get
            Return lDefaultPriceID
        End Get

        Set(ByVal Value As Long)
            lDefaultPriceID = Value
        End Set

    End Property

    Public Property COABalanceSheet() As Long

        Get
            Return lCOABalanceSheet
        End Get

        Set(ByVal Value As Long)
            lCOABalanceSheet = Value
        End Set

    End Property

    Public Property COAExportID() As Long

        Get
            Return lCOAExportID
        End Get

        Set(ByVal Value As Long)
            lCOAExportID = Value
        End Set

    End Property

    Public Property COAProfitLoss() As Long

        Get
            Return lCOAProfitLoss
        End Get

        Set(ByVal Value As Long)
            lCOAProfitLoss = Value
        End Set

    End Property



#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "ProductsGeneralProcedures"

    Public Function ReturnProductID(ByVal strValProductName As String, _
      ByVal strValProductTextCode As String) As Long

        Try

            Dim objProd As IMProducts = New IMProducts

            With objProd

                If .Find("SELECT * FROM ProductMaster WHERE " & _
                " ProductName = '" & strValProductName & _
                "' AND ProductTextCode = '" & strValProductTextCode & "'", _
                    True) = False Then

                    objProd = Nothing
                    Return 0

                Else
                    Return .ProductID

                End If

            End With

            objProd = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnProductTextCode _
        (ByVal lValProductID As Long) As String

        Try

            Dim objProd As IMProducts = New IMProducts

            With objProd

                If .Find("SELECT * FROM ProductMaster WHERE " & _
                " ProductID = " & lValProductID, _
                    True) = False Then

                    objProd = Nothing
                    Return 0

                Else
                    Return .ProductTextCode

                End If

            End With

            objProd = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnProductName _
        (ByVal lValProductID As Long) As String

        Try

            Dim objProd As IMProducts = New IMProducts

            With objProd

                If .Find("SELECT * FROM ProductMaster WHERE " & _
                " ProductID = " & lValProductID, _
                    True) = False Then

                    objProd = Nothing
                    Return 0

                Else
                    objProd = Nothing
                    Return .ProductName

                End If

            End With

            objProd = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnProductCOABalanceSheet _
        (ByVal lValProductID As Long) As Long

        Try

            Dim objProd As IMProducts = New IMProducts

            With objProd

                If .Find("SELECT * FROM ProductMaster WHERE " & _
                " ProductID = " & lValProductID, _
                    True) = False Then

                    objProd = Nothing
                    Return 0

                Else
                    objProd = Nothing
                    Return .COABalanceSheet

                End If

            End With

            objProd = Nothing

        Catch ex As Exception

        End Try


    End Function

    Public Function ReturnProductCOAProfitAndLoss _
        (ByVal lValProductID As Long) As Long

        Try

            Dim objProd As IMProducts = New IMProducts

            With objProd

                If .Find("SELECT * FROM ProductMaster WHERE " & _
                " ProductID = " & lValProductID, _
                    True) = False Then

                    objProd = Nothing
                    Return 0

                Else
                    objProd = Nothing
                    Return .lCOAProfitLoss

                End If

            End With

            objProd = Nothing

        Catch ex As Exception

        End Try

    End Function


#End Region

#Region "DatabaseProcedures"

    'Save informaiton
    Public Function Save(ByVal DisplayErrorMessages As Boolean, _
            ByVal DisplaySuccessMessages As Boolean, _
                ByVal DisplayFailureMessages As Boolean, _
                    Optional ByVal DisplayConfirmMessages _
                        As Boolean = False) As Boolean

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


            'If Trim(strOrganizationName) = "" Then

            '    ReturnError = "iManagement Null Access AuthorizationPlease open an existing company. > " & strOrganizationName
            '    objLogin = Nothing
            '    datSaved = Nothing

            '    Exit Function
            'End If


            If Trim(strProductName) = "" Or _
                Trim(strProductTextCode) = "" Or _
                    Trim(strProductType) = "" _
                Then

                ReturnError += " -- You must provide an appropriate Product Name" & _
                    " and accompanying Product Text Code e.g. " & _
                        Chr(13) & " Product Name = 'Mumias Sugar', 'Product" & _
                            " Text Code = '1/2 Kg Package', and the Product's Type"

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Trim(strProductType) <> "Service" _
                And Trim(strProductType) <> _
                    "Purchased for Reselling (Stock)" _
                        And Trim(strProductType) <> _
                            "Purchased for internal use or Raw Material" _
                Then

                ReturnError += " -- Please provide the appropriate product type:" & _
                    Chr(10) & "The Product Type can either be a:" & _
                        Chr(10) & "1.Service." & _
                            Chr(10) & "2.Purchased for Reselling (Stock)."

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if the default price exists
            If lDefaultPriceID <> 0 Then

                If Find("SELECT * FROM ProductPrice WHERE ProductID = " _
                    & lProductID & " AND PricePerUnit = " & _
                        lDefaultPriceID, False) = False Then

                    ReturnError += " -- The Default Price does not exist. Please select an existing price identifier."

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            Else

                'Check if the price provided is enabled
                If Find("SELECT * FROM ProductPrice WHERE ProductID = " & _
                     lProductID & " AND PricePerUnit = " & _
                         lDefaultPriceID & " And PriceCodeStatus = True", False) Then

                    ReturnError += " -- The Default Price identifier exists but has been disabled." & _
                        Chr(13) & "Only enabled price identifiers can be used."

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If


            'Check if the default Supplier exists
            If lDefaultSupplierID <> 0 Then

                If Find("SELECT * FROM Suppliers WHERE SupplierOrganizationID = " _
                    & lDefaultSupplierID, False) = False Then

                    ReturnError += " -- The Default Supplier does not exist. Please select an existing Supplier."

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            Else

                'Check if the Supplier provided is enabled
                If Find("SELECT * FROM Suppliers WHERE SupplierOrganizationID = " _
                    & lDefaultSupplierID & _
                        " AND SupplierStatus = FALSE", _
                            False) = True Then

                    ReturnError += " -- The Supplier identifier been disabled." & _
                                        Chr(13) & "Only enabled Suppliers can be used."
                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If

            End If


            'Check if there is an existing series with this name
            If Find("SELECT * FROM ProductMaster WHERE ProductName = '" _
                & Trim(strProductName) & "' AND ProductTextCode = '" & _
                Trim(strProductTextCode) & "'", False) = True Then
                ReturnError = "gothere"

                If Find("SELECT * FROM ProductMaster WHERE " & _
                " ProductName = '" & Trim(strProductName) & _
                "' AND ProductTextCode = '" & Trim(strProductTextCode) & _
                "' AND ProductStatus = " & bProductStatus & _
                "  AND DefaultSupplierID = " & lDefaultSupplierID & _
                "  AND DefaultPriceID = " & lDefaultPriceID & _
                "  AND COASalesID = " & lCOABalanceSheet & _
                "  AND COAExportID = " & lCOAExportID & _
                "  AND COAProfitLoss = " & lCOAProfitLoss & _
                "  AND ProductType = '" & Trim(strProductType) & _
                "' AND ProductManufacturer = '" & Trim(strProductManufacturer) & _
                "' AND UnitName = '" & Trim(strUnitName) & _
                "' AND ProductSize = " & Trim(lProductSize) & _
                "  AND ProductModelNoSerialNo = '" & Trim(strProductModelNoSerialNo) & _
                "' AND ProductModelName = '" & Trim(strProductModelName) & _
                "' AND ProductColor = '" & Trim(strProductColor) & "'", _
                False) = True Then

                    Return True

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If


                If DisplayConfirmMessages = True Then
                    If MsgBox("The Product has already exists." & _
                        Chr(10) & "Do you want to update the details?", _
                                MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                        MsgBoxResult.Yes Then

                        Update("UPDATE ProductMaster SET " & _
                        " ProductStatus = " & bProductStatus & _
                        " , DefaultSupplierID = " & lDefaultSupplierID & _
                        " , DefaultPriceID = " & lDefaultPriceID & _
                        " , COASalesID = " & lCOABalanceSheet & _
                        " , COAExportID = " & lCOAExportID & _
                        " , COAProfitLoss = " & lCOAProfitLoss & _
                        " , ProductType = '" & Trim(strProductType) & _
                        " , ProductManufacturer = '" & Trim(strProductManufacturer) & _
                        "' , UnitName = '" & Trim(strUnitName) & _
                        "' , ProductSize = " & Trim(lProductSize) & _
                        " , ProductModelNoSerialNo = '" & Trim(strProductModelNoSerialNo) & _
                        "' , ProductModelName = '" & Trim(strProductModelName) & _
                        "' , ProductColor = '" & Trim(strProductColor) & _
                        "' WHERE  ProductName = '" & Trim(strProductName) & _
                        "' AND ProductTextCode = '" & Trim(strProductTextCode) & "'")

                    End If

                Else

                    Update("UPDATE ProductMaster SET " & _
                    " ProductStatus = " & bProductStatus & _
                    " , DefaultSupplierID = " & lDefaultSupplierID & _
                    " , DefaultPriceID = " & lDefaultPriceID & _
                    " , COASalesID = " & lCOABalanceSheet & _
                    " , COAExportID = " & lCOAExportID & _
                    " , COAProfitLoss = " & lCOAProfitLoss & _
                    " , ProductType = '" & Trim(strProductType) & _
                    "' , ProductManufacturer = '" & Trim(strProductManufacturer) & _
                    "' , UnitName = '" & Trim(strUnitName) & _
                    "' , ProductSize = " & Trim(lProductSize) & _
                    " , ProductModelNoSerialNo = '" & Trim(strProductModelNoSerialNo) & _
                    "' , ProductModelName = '" & Trim(strProductModelName) & _
                    "' , ProductColor = '" & Trim(strProductColor) & _
                    "' WHERE  ProductName = '" & Trim(strProductName) & _
                    "' AND ProductTextCode = '" & Trim(strProductTextCode) & "'")

                End If


                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            If DisplayConfirmMessages = True Then
                If MsgBox("Are you sure you want to save this Product Text Code?", _
                        MsgBoxStyle.YesNo, _
                            "iManagement - Add Product?") = _
                                MsgBoxResult.No Then

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If


            'lProductID = objLogin.ReturnMaxLongValue _
            '            (strOrgAccessConnString, _
            '                "SELECT Max(ProductID) As MaxValue " & _
            '                " FROM ProductMaster WHERE ProductName = '" & _
            '                strProductName & "'") + 1

            strInsertInto = "INSERT INTO ProductMaster (" & _
                "ProductName," & _
                "ProductTextCode," & _
                "ProductStatus," & _
                "DefaultSupplierID," & _
                "DefaultPriceID," & _
                "COASalesID," & _
                "COAExportID," & _
                "COAProfitLoss," & _
                "ProductType," & _
                "ProductSize," & _
                "ProductManufacturer," & _
                "ProductModelName," & _
                "ProductModelNoSerialNo," & _
                "ProductColor," & _
                "UnitName" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & UCase(Trim(strProductName)) & _
                    "','" & UCase(Trim(strProductTextCode)) & _
                    "'," & bProductStatus & _
                    "," & lDefaultSupplierID & _
                    "," & lDefaultPriceID & _
                    "," & lCOABalanceSheet & _
                    "," & lCOAExportID & _
                    "," & lCOAProfitLoss & _
                    ",'" & Trim(strProductType) & _
                    "'," & Trim(lProductSize) & _
                    ",'" & Trim(strProductManufacturer) & _
                    "','" & Trim(strProductModelName) & _
                    "','" & Trim(strProductModelNoSerialNo) & _
                    "','" & Trim(strProductColor) & _
                    "','" & Trim(strUnitName) & _
                    "')"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, datSaved)

            AmountInStock = 0
            SaveStockAmount(False, False, False, False)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement - New Products Saved")
                End If
                Return True

            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save New Product' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save New Product Failed")
                End If
            End If

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception
            If DisplayErrorMessages = True Then
                ReturnError += " -- iManagement - Database or " & _
                "system error '" & "Message [" & ex.Message & _
                        "] and Source [" & ex.Source & _
                            "]. Please contact the Systems Administrator"

            End If

        End Try

    End Function

    'Find Informaiton
    Public Function Find(ByVal strQuery As String, _
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
                strQuery, _
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
                    objLogin = Nothing
                    datRetData = Nothing

                    Exit Function
                End If


                If bReturnValues = True Then
                    For Each myDataRows In myDataTables.Rows

                        If IsDBNull(myDataRows("ProductID")) = False Then
                            ProductID = _
                                    myDataRows("ProductID")
                        End If

                        If IsDBNull(myDataRows("ProductName")) = False Then
                            strProductName = _
                                    myDataRows("ProductName").ToString
                        End If


                        If IsDBNull(myDataRows("ProductTextCode")) = False Then
                            strProductTextCode = _
                                    myDataRows("ProductTextCode").ToString
                        End If


                        If IsDBNull(myDataRows("ProductStatus")) = False Then
                            bProductStatus = _
                                    myDataRows("ProductStatus")
                        End If


                        If IsDBNull(myDataRows("DefaultPriceID")) = False Then
                            lDefaultPriceID = _
                                    myDataRows("DefaultPriceID")
                        End If


                        If IsDBNull(myDataRows("DefaultSupplierID")) = False Then
                            lDefaultSupplierID = _
                                    myDataRows("DefaultSupplierID")
                        End If


                        If IsDBNull(myDataRows("COAExportID")) = False Then
                            lCOAExportID = _
                                    myDataRows("COAExportID")
                        End If


                        If IsDBNull(myDataRows("COAProfitLoss")) = False Then
                            lCOAProfitLoss = _
                                    myDataRows("COAProfitLoss")
                        End If


                        If IsDBNull(myDataRows("COASalesID")) = False Then
                            lCOABalanceSheet = _
                                    myDataRows("COASalesID")
                        End If

                        If IsDBNull(myDataRows("ProductType")) = False Then
                            strProductType = _
                                    myDataRows("ProductType").ToString
                        End If

                        If IsDBNull(myDataRows("UnitName")) = False Then
                            strUnitName = _
                                    myDataRows("UnitName").ToString
                        End If



                        '=============================
                        If IsDBNull(myDataRows("ProductSize")) = False Then
                            lProductSize = _
                                    myDataRows("ProductSize")
                        End If


                        If IsDBNull(myDataRows("ProductColor")) = False Then
                            strProductColor = _
                                    myDataRows("ProductColor").ToString
                        End If


                        If IsDBNull(myDataRows("ProductModelNoSerialNo")) = False Then
                            strProductModelNoSerialNo = _
                                    myDataRows("ProductModelNoSerialNo").ToString
                        End If


                        If IsDBNull(myDataRows("ProductModelName")) = False Then
                            strProductModelName = _
                                    myDataRows("ProductModelName").ToString
                        End If

                        If IsDBNull(myDataRows("ProductManufacturer")) = False Then
                            strProductManufacturer = _
                                    myDataRows("ProductManufacturer").ToString
                        End If



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
    Public Function Delete() As Boolean

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try


            If ProductID = 0 Then
                ReturnError += "Cannot Delete. Please select an existing Product."

            End If

            strDeleteQuery = "DELETE * FROM ProductMaster" & _
            " WHERE ProductID = " & lProductID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            objLogin = Nothing
            datDelete = Nothing

            If bDelSuccess = True Then
                ReturnError += "Product record Deleted Successfully"

                Return True

            Else
                ReturnError += "'Product delete' action failed. Please make " & _
                    "sure you selected and provided the appropriate details"
            End If




        Catch ex As Exception

        End Try

    End Function

    Public Function Update(ByVal strUpQuery As String) As Boolean

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
                ReturnError += "Record Updated Successfully"
                    

                Return True

            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception
            ReturnError += " -- Update Error" & ex.Message & _
                " " & strOrgAccessConnString & " " & strUpQuery

        End Try

    End Function

#End Region


    Public Function ReturnProductNameAndProductTextCodeThatHaveReachedStockLimit _
        () As Object

        Dim strQueryToUse As String
        Dim arProducts(,,,,,) As String
        Dim objLogin As IMLogin

        Try

            strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE (ProductStockAmount.AmountInStock <= " & _
            " ProductStockManager.ReorderLevel Or " & _
            " ProductStockAmount.AmountInStock Is Null) AND " & _
            " ProductStockManager.StockManagementStatus = True AND " & _
            " ProductMaster.ProductType = 'Purchased for Reselling (Stock)' " & _
            " Order By ProductMaster.ProductName,ProductMaster.ProductTextCode ASC"

            objLogin = New IMLogin

            With objLogin
                arProducts = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "", 6)

            End With

            objLogin = Nothing

            Return arProducts

        Catch ex As Exception

        End Try


    End Function

    Public Function ReturnProductNameAndProductTextCodeThatHaveReachedStockLimit _
    (ByVal dbProductLimitAmount As Long) As Object

        Dim strQueryToUse As String
        Dim arProducts(,,,,,) As String
        Dim objLogin As IMLogin

        Try

            strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE (ProductStockAmount.AmountInStock <= " & _
            " ProductStockManager.ReorderLevel Or " & _
            " ProductStockAmount.AmountInStock Is Null) AND " & _
            " ProductStockManager.StockManagementStatus = True AND " & _
            " ProductMaster.ProductType = 'Purchased for Reselling (Stock)' " & _
            " Order By ProductMaster.ProductName,ProductMaster.ProductTextCode ASC"

            objLogin = New IMLogin

            With objLogin
                arProducts = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "", 6)

            End With

            objLogin = Nothing

            Return arProducts

        Catch ex As Exception

        End Try


    End Function

    Public Function ReturnAllInventoryProductDetails() As Object

        Dim strQueryToUse As String
        Dim objLogin As IMLogin = New IMLogin
        Dim arItems(,,,,,) As String

        Try

            strQueryToUse = "SELECT ProductMaster.ProductName, " & _
        " ProductMaster.ProductTextCode, " & _
        " (ProductStockAmount.AmountInStock - " & _
        " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
        " ProductStockAmount.AmountInStock," & _
        " ProductStockManager.ReorderLevel, " & _
        " ProductStockManager.ReorderQuantity, " & _
        " ProductStockManager.ProductID, " & _
        " ProductStockManager.StockManagementStatus " & _
        " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
        " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
        " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
        " ProductStockManager.ProductID "


            With objLogin

                arItems = .FillArray _
                    (strOrgAccessConnString, strQueryToUse, "", 6)

                If .TheReturnError <> "" Then
                    ReturnError = "Error 1: " & .TheReturnError
                End If

            End With

            objLogin = Nothing

            Return arItems

        Catch ex As Exception
            ReturnError = "System Error Details (" & _
                ex.Message & "-" & "Source: " & ex.Source & ")"
        End Try

    End Function

    Public Function ReturnAllInventoryProductDetailsForProductsOfType _
        (ByVal strProductType As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductType = '" & strProductType & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnAllInventoryProductDetailsForProductsOfName _
        (ByVal strProductName As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnAllTrackingProductsAssignedForThePurposeOf _
        (ByVal bEnabledProductsOnly As Boolean) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    'By cost centre
    Public Function ReturnAllTrackingInventoryProductDetailsForCostCentreName _
       (ByVal strProductName As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    'By User Name
    Public Function ReturnAllTrackingInventoryProductDetailsForUserName _
       (ByVal strProductName As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    'For x days
    Public Function ReturnAllTrackingInventoryProductDetailsNotassignedForXDays _
       (ByVal strProductName As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    'On the date
    Public Function ReturnAllTrackingInventoryProductDetailsAssignedOnTheDateX _
       (ByVal strProductName As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    'On the date
    Public Function ReturnAllTrackingInventoryProductDetailsAssignedBroughtBySupplierX _
       (ByVal strProductName As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    'On the date
    Public Function ReturnAllTrackingInventoryProductDetailsAssignedDeliveredOnDateX _
       (ByVal strProductName As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    'On the date
    Public Function ReturnAllDeliveryInventoryProductDetailsProductsToBeDeliveredOnDateX _
       (ByVal strProductName As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    'On the date
    Public Function ReturnAllDeliveryInventoryProductDetailsProductsToBeDeliveredBeforeDateX _
       (ByVal strProductName As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    'On the date
    Public Function ReturnAllDeliveryInventoryProductDetailsProductsToBeDeliveredBySupplierX _
       (ByVal strProductName As String) As Object

        Try

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems(,,,,,) As String

            Try

                strQueryToUse = "SELECT ProductMaster.ProductName, " & _
            " ProductMaster.ProductTextCode, " & _
            " (ProductStockAmount.AmountInStock - " & _
            " ProductStockManager.ReorderLevel) AS AmountToStockLimit, " & _
            " ProductStockAmount.AmountInStock," & _
            " ProductStockManager.ReorderLevel, " & _
            " ProductStockManager.ReorderQuantity, " & _
            " ProductStockManager.ProductID, " & _
            " ProductStockManager.StockManagementStatus " & _
            " FROM (ProductMaster LEFT JOIN ProductStockAmount ON " & _
            " ProductMaster.ProductID = ProductStockAmount.ProductID) " & _
            " INNER JOIN ProductStockManager ON ProductMaster.ProductID = " & _
            " ProductStockManager.ProductID " & _
            " WHERE ProductName = '" & strProductName & "'"


                With objLogin

                    arItems = .FillArray _
                        (strOrgAccessConnString, strQueryToUse, "", 6)

                    If .TheReturnError <> "" Then
                        ReturnError = "Error 1: " & .TheReturnError
                    End If

                End With

                objLogin = Nothing

                Return arItems

            Catch ex As Exception
                ReturnError = "System Error Details (" & _
                    ex.Message & "-" & "Source: " & ex.Source & ")"
            End Try

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnAllProducts _
    (ByVal bEnabledProductsOnly As Boolean) As String()

        Dim strQueryToUse As String
        Dim objLogin As IMLogin = New IMLogin
        Dim arItems() As String

        Try

            If bEnabledProductsOnly = False Then
                strQueryToUse = "SELECT DISTINCT ProductName FROM ProductMaster"

            Else
                strQueryToUse = "SELECT DISTINCT ProductName " & _
                    "FROM ProductMaster WHERE ProductStatus = TRUE"

            End If

            With objLogin

                arItems = .FillArray _
                    (strOrgAccessConnString, strQueryToUse, "", "")

                If .TheReturnError <> "" Then
                    ReturnError = "Error 1: " & .TheReturnError
                End If

            End With

            objLogin = Nothing

            Return arItems

        Catch ex As Exception
            ReturnError = "System Error Details (" & _
                ex.Message & "-" & "Source: " & ex.Source & ")"
        End Try

    End Function

    Public Function ReturnAllProductsTextCodesForProduct _
        (ByVal strValProductName As String) As String()

        Dim strQueryToUse As String
        Dim objLogin As IMLogin = New IMLogin
        Dim arItems() As String

        Try

            strQueryToUse = "SELECT ProductTextCode FROM ProductMaster " & _
                "WHERE ProductName = '" & strValProductName & "'"

            With objLogin

                arItems = .FillArray _
                    (strOrgAccessConnString, strQueryToUse, "", "")

                'If .ReturnError <> "" Then
                '    ReturnError = "Error 1: " & .ReturnError
                'End If

            End With

            objLogin = Nothing

            Return arItems

        Catch ex As Exception

        End Try

    End Function

End Class


