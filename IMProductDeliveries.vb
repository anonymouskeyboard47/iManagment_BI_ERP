Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMProductDeliveries
    Inherits IMProductStockAmount


#Region "PrivateVariables"

    Private lDeliveryID As Long
    Private dtDeliveryDate As Date
    Private lOrderID As Long
    Private dbQuantityDelivered As Double
    Private dbQuantityAccepted As Double
    Private dbQuantityRejected As Double
    Private strDeliverySummaryDetails As String
    Private strDeliverySerialNo As String
    Private dtDateRegistered As String
    Private lApprovalOfficerUserID As Long

#End Region

#Region "Properties"

    Public Property ApprovalOfficeUserID() As Long

        Get
            Return lApprovalOfficerUserID
        End Get

        Set(ByVal Value As Long)
            lApprovalOfficerUserID = Value
        End Set

    End Property

    Public Property DeliverySummaryDetails() As String

        Get
            Return strDeliverySummaryDetails
        End Get

        Set(ByVal Value As String)
            strDeliverySummaryDetails = Value
        End Set

    End Property

    Public Property DeliverySerialNo() As String

        Get
            Return strDeliverySerialNo
        End Get

        Set(ByVal Value As String)
            strDeliverySerialNo = Value
        End Set

    End Property

    Public Property DeliveryID() As Long

        Get
            Return lDeliveryID
        End Get

        Set(ByVal Value As Long)
            lDeliveryID = Value
        End Set

    End Property

    Public Property DeliveryDate() As Date

        Get
            Return dtDeliveryDate
        End Get

        Set(ByVal Value As Date)
            dtDeliveryDate = Value
        End Set

    End Property

    Public Property OrderID() As Long

        Get
            Return lOrderID
        End Get

        Set(ByVal Value As Long)
            lOrderID = Value
        End Set

    End Property

    Public Property QuantityDelivered() As Double

        Get
            Return dbQuantityDelivered
        End Get

        Set(ByVal Value As Double)
            dbQuantityDelivered = Value
        End Set

    End Property

    Public Property QuantityAccepted() As Double

        Get
            Return dbQuantityAccepted
        End Get

        Set(ByVal Value As Double)
            dbQuantityAccepted = Value
        End Set

    End Property

    Public Property QuantityRejected() As Double

        Get
            Return dbQuantityRejected
        End Get

        Set(ByVal Value As Double)
            dbQuantityRejected = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "GeneralProcedures"

    Public Function CalculateNextDeliverySerialNo() As String
        Try

            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String
            Dim strProposedSerialNo As String

            Dim objLogin As IMLogin = New IMLogin


            MyMaxValue = objLogin.FillArray(strOrgAccessConnString, _
                        "SELECT COUNT(*) AS TotalRecords FROM" & _
                            " ProductDeliveries WHERE DateRegistered = Now()", "", "")

            objLogin = Nothing

            If Not MyMaxValue Is Nothing Then
                For Each strItem In MyMaxValue
                    If Not strItem Is Nothing Then

                        MaxValue = CLng(Val(strItem))

                    End If
                Next
            End If

            MaxValue = MaxValue + 1

            strProposedSerialNo = "Delivery" & Now.Day.ToString _
                & Now.Month.ToString & _
                    Now.Year.ToString & _
                            MaxValue.ToString

            Return strProposedSerialNo

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, _
                "iManagement - System Error")
        End Try

    End Function

#End Region


#Region "DatabaseProcedures"

    'Save informaiton
    Public Function SaveDelivery(ByVal DisplayErrorMessages As Boolean, _
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


            If lOrderID = 0 Or _
                dbQuantityDelivered = 0 Or _
                    dbQuantityAccepted = 0 _
                        Then

                MsgBox("You must provide an appropriate Existing Order" & _
                    ", the Quantity Delivered, and the Quantity Accepted. " _
                                , MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if there is an existing series with this name
            If FindStockAmount("SELECT * FROM ProductDeliveries WHERE DeliveryID = " _
                & lDeliveryID & " OR DeliverySerialNo = '" _
                    & Trim(strDeliverySerialNo) & "'", False) = True Then

                If MsgBox("The Product Delivery has already exists." & _
                    Chr(10) & "Do you want to update the details?", _
                            MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                    MsgBoxResult.Yes Then


                    UpdateStockAmount("UPDATE ProductDeliveries SET " & _
                        "DeliveryDate = '" & dtDeliveryDate & _
                        "' , OrderID = " & lOrderID & _
                        " , QuantityDelivered = " & dbQuantityDelivered & _
                        " , QuantityAccepted = " & dbQuantityAccepted & _
                        " , QuantityRejected = " & dbQuantityRejected & _
                        " , DeliverySummaryDetails = '" & _
                        Trim(strDeliverySummaryDetails) & _
                        "' , ApprovalOfficerUserID = " & lApprovalOfficerUserID & _
                        " WHERE  DeliveryID = " & lDeliveryID & _
                        " AND DeliverySerialNo = '" & Trim(strDeliverySerialNo) & "'", _
                        False, False, False, False)

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            strInsertInto = "INSERT INTO ProductDeliveries (" & _
                "DeliveryDate," & _
                "OrderID," & _
                "QuantityDelivered," & _
                "QuantityAccepted," & _
                "QuantityRejected," & _
                "DeliverySummaryDetails," & _
                "DeliverySerialNo," & _
                "DateRegistered," & _
                "ApprovalOfficerUserID" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(#" & dtDeliveryDate & _
                    "#," & lOrderID & _
                    "," & dbQuantityDelivered & _
                    "," & dbQuantityAccepted & _
                    "," & dbQuantityRejected & _
                    ",'" & strDeliverySummaryDetails & _
                    "','" & strDeliverySerialNo & _
                    "',#" & dtDateRegistered & _
                    "#," & lApprovalOfficerUserID & _
                    ")"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - New Product Delivery Saved")

                End If
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save New Product' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save New Product Delivery Failed")
                End If
            End If

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception
            If DisplayErrorMessages = True Then
                MsgBox(ex.Source, MsgBoxStyle.Critical, _
                    "iManagement - Database or system error")
            End If

        End Try

    End Function

    'Find Informaiton
    Public Function FindDelivery(ByVal strQuery As String, _
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


                        lDeliveryID = myDataRows("DeliveryID")
                        dtDeliveryDate = myDataRows("DeliveryDate")
                        lOrderID = myDataRows("OrderID")
                        dbQuantityDelivered = myDataRows("QuantityDelivered")
                        dbQuantityAccepted = myDataRows("QuantityAccepted")
                        dbQuantityRejected = myDataRows("QuantityRejected")

                        strDeliverySummaryDetails = _
                                myDataRows("DeliverySummaryDetails")

                        strDeliverySerialNo = myDataRows("DeliverySerialNo")
                        dtDateRegistered = myDataRows("DateRegistered")

                        lApprovalOfficerUserID = _
                                myDataRows("ApprovalOfficerUserID")


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
    Public Function DeleteDelivery(ByVal strDelQuery As String) As Boolean

        If Trim(strDeliverySerialNo) = "" Then

            MsgBox("Cannot Delete. Please select an existing Product.", _
                MsgBoxStyle.Exclamation, _
                "iManagement - invalid or incomplete information")
            Exit Function

        End If

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try


            strDeleteQuery = strDelQuery



            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", _
                MsgBoxStyle.Information, _
                "iManagement - Product Delivery Deleted")
            Else
                MsgBox("'Product Delivery delete' action failed", _
                MsgBoxStyle.Exclamation, _
                "Product Delivery Deletion failed")
            End If


            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Sub UpdateDelivery(ByVal strUpQuery As String)

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
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Product Delivery Details Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub


#End Region

End Class
