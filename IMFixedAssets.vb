
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMFixedAssets


#Region "PrivateVariables"

    Private lFixedAssetID As Long
    Private lDepreciationType As Long '1 Double-Declining; 2 Single-Line; 3 Sum-of-Year Digits
    Private strFixedAssetName As String
    Private dtDepreciationStarts As Date
    Private decCost As Decimal
    Private lLife As Long 'In months
    Private decSalvage As Decimal
    Private decCurrentValue As Decimal
    Private decTotalDepreciation As Decimal
    Private dtLastDepreciated As Date
    Private strFixedAssetType As String  '1 Plant;2 Property;3 Equipment;4 Vehicle;5;Other
    Private strFixedAssetSerialNumber As String
    Private strFixedAssetDescription As String
    Private lFixedAssetTypeIMCategory As Long 'Plot,Building,Water Meter,Electricity Meter,Customer,Plot Beacon,Vehicle,Fire Hydrant,
    Private lFixedAssetObjectID As Long
    Private bFixedAssetStatus As Boolean
    Private strFixedAssetStatusDescription As String
    Private lRentProductID As Long
    Private lMeterProductID As Long
    Private strRentType As String
    Private decSalePrice As Decimal
    Private lRentPaymentMode As Long

    Private lFixedAssetSize As Double
    Private strFixedAssetColor As String
    Private strFixedAssetModelNoSerialNo As String
    Private strFixedAssetModelName As String
    Private strFixedAssetManufacturer As String


#End Region


#Region "Properties"

    Public Property FixedAssetSize() As Long

        Get
            Return lFixedAssetSize
        End Get

        Set(ByVal Value As Long)
            lFixedAssetSize = Value
        End Set

    End Property

    Public Property FixedAssetColor() As String

        Get
            Return strFixedAssetColor
        End Get

        Set(ByVal Value As String)
            strFixedAssetColor = Value
        End Set

    End Property

    'Serial number for that particular model
    Public Property FixedAssetModelNoSerialNo() As String

        Get
            Return strFixedAssetModelNoSerialNo
        End Get

        Set(ByVal Value As String)
            strFixedAssetModelNoSerialNo = Value
        End Set

    End Property

    'Name of the model, Compaq Presario
    Public Property FixedAssetModelName() As String

        Get
            Return strFixedAssetModelName
        End Get

        Set(ByVal Value As String)
            strFixedAssetModelName = Value
        End Set

    End Property

    Public Property FixedAssetManufacturer() As String

        Get
            Return strFixedAssetManufacturer
        End Get

        Set(ByVal Value As String)
            strFixedAssetManufacturer = Value
        End Set

    End Property

    Public Property ReturnError() As Long

        Get
            Return ReturnError
        End Get

        Set(ByVal Value As Long)
            ReturnError = Value
        End Set

    End Property

    Public Property FixedAssetID() As Long

        Get
            Return lFixedAssetID
        End Get

        Set(ByVal Value As Long)
            lFixedAssetID = Value
        End Set

    End Property

    Public Property Life() As Long

        Get
            Return lLife
        End Get

        Set(ByVal Value As Long)
            lLife = Value
        End Set

    End Property

    Public Property DepreciationType() As Long

        Get
            Return lDepreciationType
        End Get

        Set(ByVal Value As Long)
            lDepreciationType = Value
        End Set

    End Property

    Public Property FixedAssetName() As String

        Get
            Return strFixedAssetName
        End Get

        Set(ByVal Value As String)
            strFixedAssetName = Value
        End Set

    End Property

    Public Property DepreciationStarts() As Date

        Get
            Return dtDepreciationStarts
        End Get

        Set(ByVal Value As Date)
            dtDepreciationStarts = Value
        End Set

    End Property

    Public Property Cost() As Decimal

        Get
            Return decCost
        End Get

        Set(ByVal Value As Decimal)
            decCost = Value
        End Set

    End Property

    Public Property Salvage() As Decimal

        Get
            Return decSalvage
        End Get

        Set(ByVal Value As Decimal)
            decSalvage = Value
        End Set

    End Property

    Public Property CurrentValue() As Decimal

        Get
            Return decCurrentValue
        End Get

        Set(ByVal Value As Decimal)
            decCurrentValue = Value
        End Set

    End Property

    Public Property TotalDepreciation() As Decimal

        Get
            Return decTotalDepreciation
        End Get

        Set(ByVal Value As Decimal)
            decTotalDepreciation = Value
        End Set

    End Property

    Public Property LastDepreciated() As Date

        Get
            Return dtLastDepreciated
        End Get

        Set(ByVal Value As Date)
            dtLastDepreciated = Value
        End Set

    End Property

    Public Property FixedAssetType() As String

        Get
            Return strFixedAssetType
        End Get

        Set(ByVal Value As String)
            strFixedAssetType = Value
        End Set

    End Property

    Public Property FixedAssetSerialNumber() As String

        Get
            Return strFixedAssetSerialNumber
        End Get

        Set(ByVal Value As String)
            strFixedAssetSerialNumber = Value
        End Set

    End Property

    Public Property FixedAssetDescription() As String

        Get
            Return strFixedAssetDescription
        End Get

        Set(ByVal Value As String)
            strFixedAssetDescription = Value
        End Set

    End Property

    Public Property FixedAssetTypeIMCategory() As Long

        Get
            Return lFixedAssetTypeIMCategory
        End Get

        Set(ByVal Value As Long)
            lFixedAssetTypeIMCategory = Value
        End Set

    End Property

    Public Property FixedAssetObjectID() As Long

        Get
            Return lFixedAssetObjectID
        End Get

        Set(ByVal Value As Long)
            lFixedAssetObjectID = Value
        End Set

    End Property

    Public Property FixedAssetStatus() As Boolean

        Get
            Return bFixedAssetStatus
        End Get

        Set(ByVal Value As Boolean)
            bFixedAssetStatus = Value
        End Set

    End Property

    Public Property FixedAssetStatusDescription() As String

        Get
            Return strFixedAssetStatusDescription
        End Get

        Set(ByVal Value As String)
            strFixedAssetStatusDescription = Value
        End Set

    End Property

    Public Property RentProductID() As Long

        Get
            Return lRentProductID
        End Get

        Set(ByVal Value As Long)
            lRentProductID = Value
        End Set

    End Property

    Public Property MeterProductID() As Long

        Get
            Return lMeterProductID
        End Get

        Set(ByVal Value As Long)
            lMeterProductID = Value
        End Set

    End Property

    Public Property RentType() As String

        Get
            Return strRentType
        End Get

        Set(ByVal Value As String)
            strRentType = Value
        End Set

    End Property

    Public Property SalePrice() As Decimal

        Get
            Return decSalePrice
        End Get

        Set(ByVal Value As Decimal)
            decSalePrice = Value
        End Set

    End Property


#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "GeneralProcedures"

    Public Function ReturnAllAssetsSerialNoForAssetsOfType() As String()

    End Function

    Public Function ReturnAllAssetsOfType() As Object

    End Function

    Public Function ReturnAllAssetsSerialNumbers _
        (ByVal bIncludeDisabledAssets As Boolean, _
            ByVal strValFixedAssetType As Long, _
                ByVal strValFixedAssetIMType As Long) As Object


        Dim strQueryToUse As String
        Dim arReturnValue As Object
        Dim objLogin As IMLogin

        Try

            If bIncludeDisabledAssets = True Then

                If Trim(strValFixedAssetType) = "" And _
                    Trim(strValFixedAssetIMType) = "" Then
                    strQueryToUse = "SELECT * FROM FixedAssets " & _
                        "ORDER BY strFixedAssetSerialNumber"
                End If

                If Trim(strValFixedAssetType) <> "" And _
                    Trim(strValFixedAssetIMType) = "" Then

                    strQueryToUse = "SELECT * FROM FixedAssets " & _
                    " WHERE (FixedAssetType = '" & strValFixedAssetType & _
                    "') ORDER BY FixedAssetSerialNumber ASC"

                End If

                If Trim(strValFixedAssetType) = "" And _
                    Trim(strValFixedAssetIMType) <> "" Then

                    strQueryToUse = "SELECT * FROM FixedAssets " & _
                     " WHERE (FixedAssetIMType = '" & _
                     strValFixedAssetIMType & _
                     "') ORDER BY FixedAssetSerialNumber ASC"

                End If

                If Trim(strValFixedAssetType) <> "" And _
                    Trim(strValFixedAssetIMType) <> "" Then

                    strQueryToUse = "SELECT * FROM FixedAssets " & _
                      " WHERE (FixedAssetType = '" & strValFixedAssetIMType & _
                      "' AND FixedAssetIMType = '" & strValFixedAssetIMType & _
                      "') ORDER BY FixedAssetSerialNumber ASC"

                End If

            Else


                If Trim(strValFixedAssetType) = "" And _
                    Trim(strValFixedAssetIMType) = "" Then
                    strQueryToUse = "SELECT * FROM FixedAssets " & _
                      "WHERE FixedAssetStatus = TRUE" & _
                      " ORDER BY FixedAssetSerialNumber ASC"
                End If

                If Trim(strValFixedAssetType) <> "" And _
                    Trim(strValFixedAssetIMType) = "" Then

                    strQueryToUse = "SELECT * FROM FixedAssets " & _
                    " WHERE (FixedAssetType = '" & strValFixedAssetType & _
                      "' AND FixedAssetStatus = TRUE" & _
                      ") ORDER BY FixedAssetSerialNumber ASC"

                End If

                If Trim(strValFixedAssetType) = "" And _
                    Trim(strValFixedAssetIMType) <> "" Then

                    strQueryToUse = "SELECT * FROM FixedAssets " & _
                     " WHERE (FixedAssetIMType = '" & _
                     strValFixedAssetIMType & _
                      "' AND FixedAssetStatus = TRUE" & _
                      ") ORDER BY FixedAssetSerialNumber ASC"

                End If

                If Trim(strValFixedAssetType) <> "" And _
                    Trim(strValFixedAssetIMType) <> "" Then

                    strQueryToUse = "SELECT * FROM FixedAssets " & _
                      " WHERE (FixedAssetType = '" & strValFixedAssetIMType & _
                      "' AND FixedAssetIMType = '" & strValFixedAssetIMType & _
                      "' AND FixedAssetStatus = TRUE" & _
                      ") ORDER BY FixedAssetSerialNumber ASC"

                End If

            End If



            objLogin = New IMLogin

            With objLogin
                arReturnValue = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "", 6)

            End With

            objLogin = Nothing

            Return arReturnValue

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnAssetObjectWithSerialNumber _
        (ByVal strValFixedAssetSerial As String, _
            ByVal bIncludeDisabledAssets As Boolean) As Boolean

        Try

            If Find("SELECT * FROM FixedAssets " & _
                "WHERE FixedAssetSerialNumber = '" & _
                    strValFixedAssetSerial & "'", True) = True Then
                Return True
            End If

        Catch ex As Exception
            ReturnError += " + " & ex.Message
        End Try

    End Function

    Public Function DetermineIfqueryIsTrues _
        (ByVal strQuery As String) As String

    End Function

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

            If lDepreciationType = 0 Or Trim(strFixedAssetType) = "" _
                Or Trim(strFixedAssetSerialNumber) = "" Or _
                    lFixedAssetTypeIMCategory = 0 Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please provide the following details in" & _
                " order to save the Asset's details: " & _
                Chr(10) & "1. Depreciation Type " & _
                Chr(10) & "2. Asset's name " & _
                Chr(10) & "2. Asset's Serial Number (e.g. Vehicle's chasis number, electricity/water meter serial number, computer's serial number " & _
                Chr(10) & "3. Asset's general type (e.g. Property, Vehicle, Other etc) " & _
                Chr(10) & "4. Asset's iManagement type (e.g. Plot,Building,Water Meter, Other etc) "

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM FixedAssets " & _
                " WHERE (FixedAssetSerialNumber = '" & strFixedAssetSerialNumber & _
                "' AND FixedAssetType = '" & strFixedAssetType & "') OR " & _
                " (FixedAssetSerialNumber = '" & strFixedAssetSerialNumber & _
                "' AND FixedAssetType = '" & strFixedAssetType & "')", _
                False) = True Then

                Update("UPDATE FixedAssets SET " & _
                " DepreciationType = " & lDepreciationType & _
                ",FixedAssetName = '" & strFixedAssetName & _
                "',DepreciationStarts = #" & dtDepreciationStarts & _
                "#,Cost = " & decCost & _
                ",Salvage = " & decSalvage & _
                ",Life = " & lLife & _
                ",CurrentValue = " & decCurrentValue & _
                ",TotalDepreciation = " & decTotalDepreciation & _
                ",LastDepreciated = #" & dtLastDepreciated & _
                "#,FixedAssetType = '" & strFixedAssetType & _
                "',FixedAssetSerialNumber = '" & strFixedAssetSerialNumber & _
                "',FixedAssetDescription = '" & strFixedAssetDescription & _
                "',FixedAssetTypeIMCategory = " & lFixedAssetTypeIMCategory & _
                ",FixedAssetObjectID = " & lFixedAssetObjectID & _
                ",FixedAssetStatus = " & bFixedAssetStatus & _
                ",FixedAssetStatusDescription  = '" & strFixedAssetStatusDescription & _
                ",RentProductID = " & lRentProductID & _
                ",MeterProductID = " & lMeterProductID, _
                True, True, True, True)

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO FixedAssets (" & _
                "FixedAssetID," & _
                "DepreciationType," & _
                "FixedAssetName," & _
                "DepreciationStarts," & _
                "Cost," & _
                "Salvage," & _
                "CurrentValue," & _
                "TotalDepreciation," & _
                "LastDepreciated," & _
                "FixedAssetType," & _
                "FixedAssetSerialNumber," & _
                "FixedAssetDescription," & _
                "FixedAssetTypeIMCategory," & _
                "FixedAssetObjectID," & _
                "FixedAssetStatus," & _
                "FixedAssetStatusDescription," & _
                "RentProductID," & _
                "MeterProductID," & _
                "RentType," & _
                "SalePrice," & _
                "RentPaymentMode," & _
                "Life" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lFixedAssetID & _
                    "," & lDepreciationType & _
                    ",'" & strFixedAssetName & _
                    "',#" & dtDepreciationStarts & _
                    "#," & decCost & _
                    "," & decSalvage & _
                    "," & decCurrentValue & _
                    "," & decTotalDepreciation & _
                    ",#" & dtLastDepreciated & _
                    "#,'" & strFixedAssetType & _
                    "','" & strFixedAssetSerialNumber & _
                    "','" & strFixedAssetDescription & _
                    "'," & lFixedAssetTypeIMCategory & _
                    "," & lFixedAssetObjectID & _
                    "," & bFixedAssetStatus & _
                    ",'" & strFixedAssetStatusDescription & _
                    "'," & lRentProductID & _
                    "," & lMeterProductID & _
                    ",'" & strRentType & _
                    "'," & decSalePrice & _
                    "," & lRentPaymentMode & _
                    "," & lLife & _
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
                    ReturnSuccess += "Fixed Asset Saved Successfully."

                End If
            Else

                If DisplayFailure = True Then
                    ReturnError = "'Save Fixed Asset' action failed." & _
            " Make sure all mandatory details are entered"

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
                ReturnError += ex.Message.ToString

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

                            lFixedAssetID = _
                                myDataRows("FixedAssetID")
                            lDepreciationType = _
                                myDataRows("DepreciationType")
                            strFixedAssetName = _
                                myDataRows("FixedAssetName")
                            dtDepreciationStarts = _
                                myDataRows("DepreciationStarts")
                            decCost = _
                                myDataRows("Cost")
                            lLife = _
                                myDataRows("Life")

                            '=========================
                            decSalvage = _
                                myDataRows("Salvage")
                            decCurrentValue = _
                                myDataRows("CurrentValue")
                            decTotalDepreciation = _
                                myDataRows("TotalDepreciation")
                            dtLastDepreciated = _
                                myDataRows("LastDepreciated")
                            strFixedAssetType = _
                                myDataRows("FixedAssetType")

                            '=========================
                            strFixedAssetSerialNumber = _
                                myDataRows("FixedAssetSerialNumber")
                            strFixedAssetDescription = _
                                myDataRows("FixedAssetDescription")
                            lFixedAssetTypeIMCategory = _
                                myDataRows("FixedAssetTypeIMCategory")
                            lFixedAssetObjectID = _
                                myDataRows("FixedAssetObjectID")
                            bFixedAssetStatus = _
                                myDataRows("FixedAssetStatus")

                            '=========================
                            strFixedAssetStatusDescription = _
                                myDataRows("FixedAssetStatusDescription")
                            lRentProductID = _
                                myDataRows("RentProductID")
                            lMeterProductID = _
                                myDataRows("MeterProductID")
                            strRentType = _
                                myDataRows("RentType")
                            decSalePrice = _
                                myDataRows("SalePrice")
                            lRentPaymentMode = _
                                myDataRows("RentPaymentMode")

                        Next

                    End If

                Next
                Return True

            Else
                Return False

            End If

            datRetData = Nothing
            objLogin = Nothing

        Catch ex As Exception
            ReturnError += ex.Message.ToString

        End Try

    End Function


    Public Function Delete() As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = "DELETE * FROM FixedAssets " & _
            "WHERE FixedAssetSerialNumber = '" _
            & strFixedAssetSerialNumber & _
            "' AND FixedAssetType = '" & strFixedAssetType & "'"

            If strFixedAssetSerialNumber = 0 Or _
                Trim(strFixedAssetType) = "" Or _
                lFixedAssetTypeIMCategory = 0 Then

                ReturnError += "You must provide a fixed asset serial " & _
                        "number, Fixed Asset Type, and" & _
                        " System Category Type in order to delete " & _
                        "the fixed asset"
                Exit Function
            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                ReturnSuccess += "Fixed Asset details deleted"
                Return True
            Else

                ReturnError += "'Delete Fixed Asset' action failed"


            End If

            datDelete = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Function


    Public Function Update(ByVal strUpQuery As String, _
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

            If Trim(strFixedAssetSerialNumber) = "" Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                    strUpdateQuery, datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    If DisplaySuccess = True Then
                        ReturnSuccess += "Fixed Asset's record updated successfully"
                        Return True

                    End If

                End If

            End If

            objLogin = Nothing
            datUpdated = Nothing


        Catch ex As Exception

        End Try
    End Function

#End Region


End Class
