
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMSupplierOrder

#Region "Connection Properties"

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

    Private lSupplierOrganizationID As Long
    Private lOrderID As Long
    Private bOrderConfirmed As Boolean

#End Region


#Region "Properties"

    Public Property OrderConfirmed() As Boolean

        Get
            Return bOrderConfirmed
        End Get

        Set(ByVal Value As Boolean)
            bOrderConfirmed = Value
        End Set

    End Property

    Public Property SupplierOrgID() As Long

        Get
            Return lSupplierOrganizationID
        End Get

        Set(ByVal Value As Long)
            lSupplierOrganizationID = Value
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

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "GeneralProcedures"

    Public Function ReturnProductSupplierOrder _
       (ByVal lValOrderID As Long, _
           ByVal lValSupplierID As Long) As String(,)

        Dim strQueryToUse As String
        Dim objLogin As IMLogin = New IMLogin
        Dim arItems(,) As String

        Try

            If lValOrderID = 0 And lValSupplierID = 0 Then

                strQueryToUse = "SELECT Name,OrderID FROM Employers " & _
                    " INNER JOIN SupplierOrder ON " & _
                    " SupplierOrder.SupplierOrganizationID = " & _
                    " Employers.EmployerID "

            ElseIf lValOrderID = 0 And lValSupplierID <> 0 Then
                strQueryToUse = "SELECT Name,OrderID FROM Employers " & _
                    " INNER JOIN SupplierOrder ON " & _
                    " SupplierOrder.SupplierOrganizationID = " & _
                    " Employers.EmployerID " & _
                    " WHERE SupplierOrder.SupplierOrganizationID = " & _
                    lValSupplierID

            ElseIf lValOrderID <> 0 And lValSupplierID = 0 Then
                strQueryToUse = "SELECT Name,OrderID FROM Employers " & _
                  " INNER JOIN SupplierOrder ON " & _
                  " SupplierOrder.SupplierOrganizationID = " & _
                  " Employers.EmployerID " & _
                  " WHERE SupplierOrder.OrderID = " & _
                  lValOrderID

            End If


            With objLogin
                arItems = .FillArray _
                    (strOrgAccessConnString, strQueryToUse, "", "", 2)

            End With

            objLogin = Nothing

            Return arItems

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnSupplierFromOrderID _
      (ByVal lValOrderID As Long) As String

        Try

            If lValOrderID = 0 Then
                Exit Function
            End If

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems() As String


            strQueryToUse = "SELECT Name FROM Employers " & _
                " INNER JOIN SupplierOrder ON " & _
                " SupplierOrder.SupplierID = " & _
                " Employers.EmployerID " & _
                " WHERE SupplierOrder.OrderID = " & _
                lValOrderID


            With objLogin
                arItems = .FillArray _
                    (strOrgAccessConnString, strQueryToUse, "", "")

            End With

            objLogin = Nothing

            If Not arItems Is Nothing Then
                Return arItems(0)
            End If


        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnSupplierFromOrderSerialNo _
     (ByVal strValOrderSerialNo As String) As String

        Try

            If Trim(strValOrderSerialNo) = "" Then
                Exit Function
            End If

            Dim strQueryToUse As String
            Dim objLogin As IMLogin = New IMLogin
            Dim arItems() As String


            strQueryToUse = "SELECT Name FROM (Employers " & _
                " INNER JOIN SupplierOrder ON " & _
                " SupplierOrder.SupplierID = " & _
                " Employers.EmployerID) " & _
                " INNER JOIN ProductOrders ON " & _
                " SupplierOrder.OrderID = " & _
                " ProductOrders.OrderID " & _
                " WHERE ProductOrders.OrderSerialNo = '" & _
                strValOrderSerialNo & "'"


            With objLogin
                arItems = .FillArray _
                    (strOrgAccessConnString, strQueryToUse, "", "")

            End With

            objLogin = Nothing

            If Not arItems Is Nothing Then
                Return arItems(0)
            End If


        Catch ex As Exception

        End Try

    End Function


#End Region


#Region "DatabaseProcedures"

    'Save informaiton
    Public Function Save(ByVal bDisplayErrorMessages As Boolean, _
                ByVal bDisplaySuccessMessages As Boolean, _
                    ByVal bDisplayFailureMessages As Boolean, _
                        ByVal bDisplayConfirmMessages As Boolean) As Boolean

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


            If lSupplierOrganizationID = 0 Or lOrderID = 0 Then

                MsgBox("You must provide an existing Supplier and an " & _
                    "existing Order.", MsgBoxStyle.Critical, _
                        "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing
                Exit Function

            End If


            'Check if there is an existing supplier with this id
            If Find("SELECT * FROM Suppliers WHERE SupplierOrganizationID = " _
                & lSupplierOrganizationID, _
                    False) = False Then

                If bDisplayConfirmMessages = False Then
                    If MsgBox("This supplier does not exist. Cannot update or save.", _
                        MsgBoxStyle.Critical, _
                            "iManagement - Supplier Record Does Not Exists") = _
                                MsgBoxResult.No Then

                        objLogin = Nothing
                        datSaved = Nothing
                        Exit Function

                    End If
                End If
            End If


            'Check if there is an existing supplier with this id
            If Find("SELECT * FROM ProductOrders WHERE OrderID = " _
                & lOrderID, _
                    False) = False Then

                If bDisplayConfirmMessages = False Then
                    If MsgBox("This Order does not exist. Cannot update or save.", _
                        MsgBoxStyle.Critical, _
                            "iManagement - Order Record Does Not Exists") = _
                                MsgBoxResult.No Then

                        objLogin = Nothing
                        datSaved = Nothing
                        Exit Function

                    End If
                End If
            End If


            'Check if there is an existing order with this Orderid
            If Find("SELECT * FROM SupplierOrder WHERE OrderID = " & lOrderID, _
                    False) = True Then

                If bDisplayConfirmMessages = False Then
                    If MsgBox("This Order has already been added. " & _
                        "Do you want to update the supplier order details.", _
                        MsgBoxStyle.YesNo, _
                        "iManagement - Record Exists. Update it?") = _
                        MsgBoxResult.No Then

                        objLogin = Nothing
                        datSaved = Nothing
                        Exit Function

                    End If
                End If


                If Find("SELECT * FROM SupplierOrder WHERE OrderID = " & _
                    lOrderID & " AND OrderConfirmed = TRUE", False) = True Then

                    If bDisplayFailureMessages = True Then
                        MsgBox("This Order cannot be modified since it " & _
                        "has already been confirmed", _
                        MsgBoxStyle.Critical, _
                        "iManagement - Record Cannot be Updated")

                        objLogin = Nothing
                        datSaved = Nothing
                        Exit Function

                    End If
                End If

                Update("UPDATE SupplierOrder SET " & _
                " SupplierID = " & lSupplierOrganizationID & _
                ", OrderConfirmed = " & bOrderConfirmed & _
                " WHERE OrderID = " & lOrderID, False, False, False, False)

                objLogin = Nothing
                datSaved = Nothing
                Exit Function

            End If


            If bDisplayConfirmMessages = True Then
                If MsgBox("Are you sure you want to add this " & _
                "Order?", MsgBoxStyle.YesNo, "iManagement - Add Record") = _
                        MsgBoxResult.No Then

                    objLogin = Nothing
                    datSaved = Nothing
                    Exit Function
                End If
            End If


            strInsertInto = "INSERT INTO SupplierOrder (" & _
                "SupplierID," & _
                "OrderID," & _
                "OrderConfirmed" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lSupplierOrganizationID & _
                    "," & lOrderID & _
                    "," & bOrderConfirmed & _
                    ")"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If bDisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - New Order Saved")

                End If
                Return True

            Else

                If bDisplayFailureMessages = True Then
                    MsgBox("'Save New Order' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save New Order Failed")
                End If
            End If

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception
            If bDisplayErrorMessages = True Then
                MsgBox(ex.Source, MsgBoxStyle.Critical, _
                    "iManagement - Database or system error")
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

        bQuerySuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strQuery, _
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

                        If IsDBNull(myDataRows("SupplierOrganizationID")) = False Then
                            lSupplierOrganizationID = _
                                    myDataRows("SupplierOrganizationID")
                        End If

                        If IsDBNull(myDataRows("OrderID")) = False Then
                            lOrderID = myDataRows("OrderID")
                        End If

                        If IsDBNull(myDataRows("OrderConfirmed")) = False Then
                            bOrderConfirmed = myDataRows("OrderConfirmed")
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

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin


            If lOrderID = 0 Then
                MsgBox("Cannot Delete any order. Please " & _
                    "select an existing Order.", MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

                datDelete = Nothing
                objLogin = Nothing
                Exit Function

            End If


            If MsgBox("Are you sure you want to delete the Order's details?" & _
                MsgBoxStyle.YesNo, _
                    "iManagement - Delete Record?") = MsgBoxResult.No Then

                datDelete = Nothing
                objLogin = Nothing
                Exit Function
            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            'Delete Order
            strDeleteQuery = "DELETE * FROM SupplierOrder " & _
                " WHERE OrderID = " & lSupplierOrganizationID

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Order Deleted")
                Return True

            Else
                MsgBox("'Order delete' action failed", _
                    MsgBoxStyle.Exclamation, "Order Deletion failed")
            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Sub Update(ByVal strUpQuery As String, _
    ByVal bDisplayErrorMessages As Boolean, _
                ByVal bDisplaySuccessMessages As Boolean, _
                    ByVal bDisplayFailureMessages As Boolean, _
                        ByVal bDisplayConfirmMessages As Boolean)

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                If bDisplaySuccessMessages = True Then
                    MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                        "iManagement - Supplier Order Details Updated")
                End If

            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try

    End Sub

#End Region


End Class
