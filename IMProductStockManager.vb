Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMProductStockManager

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

    Private lProductID As Long
    Private lStockManagementID As Long
    Private dbReorderLevel As Double
    Private dbReorderQuantity As Double
    Private bStockManagementStatus As Boolean

#End Region

#Region "Properties"

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

    Public Property ReorderLevel() As Double

        Get
            Return dbReorderLevel
        End Get

        Set(ByVal Value As Double)
            dbReorderLevel = Value
        End Set

    End Property

    Public Property ReorderQuantity() As Double

        Get
            Return dbReorderQuantity
        End Get

        Set(ByVal Value As Double)
            dbReorderQuantity = Value
        End Set

    End Property

    Public Property StockManagementStatus() As Boolean

        Get
            Return bStockManagementStatus
        End Get

        Set(ByVal Value As Boolean)
            bStockManagementStatus = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function ReturnProductReorderDetails _
        (Optional ByVal lValProductID As Long = 0, _
            Optional ByVal bIncludeStatusIsDisabled As Boolean = False) _
                As String(,)

        Dim strQueryToUse As String
        Dim objLogin As IMLogin = New IMLogin
        Dim arItems(,) As String

        Try

            If lValProductID = 0 And bIncludeStatusIsDisabled = False Then
                strQueryToUse = "SELECT ReorderLevel, ReorderQuantity " & _
                    "FROM ProductStockManager WHERE StockManagementStatus = TRUE"

            ElseIf lValProductID = 0 And bIncludeStatusIsDisabled = True Then
                strQueryToUse = "SELECT ReorderLevel, ReorderQuantity " & _
                    "FROM ProductStockManager"

            ElseIf lValProductID <> 0 Then
                strQueryToUse = "SELECT ReorderLevel, ReorderQuantity " & _
                    "FROM ProductStockManager WHERE ProductID = " & lValProductID

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

    Public Function ReturnProductStockMngtID _
    (Optional ByVal lValProductID As Long = 0, _
        Optional ByVal bIncludeStatusIsDisabled As Boolean = False) _
            As String()

        Dim strQueryToUse As String
        Dim objLogin As IMLogin = New IMLogin
        Dim arItems() As String

        Try

            If lValProductID = 0 And bIncludeStatusIsDisabled = False Then
                strQueryToUse = "SELECT StockManagementID " & _
                    "FROM ProductStockManager WHERE StockManagementStatus = TRUE"

            ElseIf lValProductID = 0 And bIncludeStatusIsDisabled = True Then
                strQueryToUse = "SELECT StockManagementID " & _
                    "FROM ProductStockManager"

            ElseIf lValProductID <> 0 Then
                strQueryToUse = "SELECT StockManagementID " & _
                    "FROM ProductStockManager WHERE ProductID = " & lValProductID

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

            If lProductID = 0 Or _
                dbReorderLevel = 0 Or _
                    dbReorderQuantity = 0 Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide the following details in" & _
                " order to save a Stock Management detail:" & _
                Chr(10) & "1. Existing Product" & _
                Chr(10) & "2. Reorder Level" & _
                Chr(10) & "3. Reorder Quantity" & _
                MsgBoxStyle.Critical, _
            "iManagement - Save Action Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            Else

                If Find("SELECT * FROM ProductStockManager WHERE ProductID = " & _
                lProductID, False) = True Then

                    If Find("SELECT * FROM ProductStockManager WHERE " & _
                    "ProductID = " & lProductID & _
                    " AND ReorderLevel = " & dbReorderLevel & _
                    " AND ReorderQuantity = " & dbReorderQuantity & _
                    " AND StockManagementStatus = " & bStockManagementStatus, _
                    False) = True Then

                        Return True

                        objLogin = Nothing
                        datSaved = Nothing
                        Exit Function

                    End If

                    If DisplayConfirmation = True Then
                        If MsgBox("The Product Reorder Details already exists." & _
                            Chr(10) & "Do you want to update the details?", _
                                    MsgBoxStyle.YesNo, _
                                        "iManagement - Record Exists") = _
                                            MsgBoxResult.No Then

                            objLogin = Nothing
                            datSaved = Nothing
                            Exit Function

                        End If
                    End If

                    If Update("UPDATE ProductStockManager SET " & _
                                " ReorderLevel = " & dbReorderLevel & _
                                " , ReorderQuantity = " & dbReorderQuantity & _
                                " , StockManagementStatus = " & bStockManagementStatus & _
                                    " WHERE  ProductID = " _
                                        & lProductID, False, False, False, False) = True Then
                        Return True
                    End If




                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If

                strInsertInto = "INSERT INTO ProductStockManager (" & _
                    "ProductID," & _
                    "ReorderLevel," & _
                    "ReorderQuantity," & _
                    "StockManagementStatus" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "(" & lProductID & _
                        "," & dbReorderLevel & _
                        "," & dbReorderQuantity & _
                        "," & bStockManagementStatus & _
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
                        MsgBox("Stock Management Saved Successfully.", _
                            MsgBoxStyle.Information, _
                                "iManagement - Record Saved Successfully")

                    End If

                Else

                    If DisplayFailure = True Then
                        MsgBox("'Save Stock Management' action failed." & _
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

                            lProductID = _
                                myDataRows("ProductID")
                            lStockManagementID = _
                                myDataRows("StockManagementID")
                            dbReorderLevel = _
                                myDataRows("ReorderLevel")
                            dbReorderQuantity = _
                                myDataRows("ReorderQuantity")
                            bStockManagementStatus = _
                                myDataRows("StockManagementStatus")

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

    Public Function Delete(ByVal strDelQuery As String)

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = strDelQuery

            If lProductID <> 0 _
                                Then

                objLogin.ConnectString = strAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
                datDelete)



                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Stock Management Details Deleted", MsgBoxStyle.Information, _
                        "iManagement - Record Deleted Successfully")

                Else

                    MsgBox("'Delete Stock Management' action failed", _
                        MsgBoxStyle.Exclamation, "Stock Management Deletion failed")

                    objLogin.RollbackTheTrans()

                End If

            Else

                MsgBox("Cannot Delete. Please select an existing Stock Management Detail", _
                        MsgBoxStyle.Exclamation, "iManagement - Missing Information")

                objLogin.RollbackTheTrans()

            End If
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

            If lProductID <> 0 _
                            Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    If DisplaySuccess = True Then
                        MsgBox("Record Updated Successfully", _
                            MsgBoxStyle.Information, _
                                "iManagement -  Stock Management Updated")
                    End If

                    Return True

                End If


            End If

        Catch ex As Exception

        End Try
    End Function

#End Region

End Class
