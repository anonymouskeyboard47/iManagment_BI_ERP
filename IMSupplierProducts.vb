Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMSupplierProducts

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
    Private lSupplierOrganizationID As Long
    Private bSupplyingStatus As Boolean

#End Region


#Region "Properties"


    Public Property SupplierOrganizationID() As Long

        Get
            Return lSupplierOrganizationID
        End Get

        Set(ByVal Value As Long)
            lSupplierOrganizationID = Value
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

    Public Property SupplyingStatus() As Boolean

        Get
            Return bSupplyingStatus
        End Get

        Set(ByVal Value As Boolean)
            bSupplyingStatus = Value
        End Set

    End Property

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "GeneralProcedures"

    Public Function ReturnProductSupplierName _
        (ByVal lValProductID As Long, _
            ByVal bIncludeDisabledSupplyingStatus As Boolean, _
                Optional ByVal strValProductName As String = "") As String(,)

        Dim strQueryToUse As String
        Dim objLogin As IMLogin = New IMLogin
        Dim arItems(,) As String

        Try


            'Returns all suppliers as long as they are enabled
            If lValProductID = 0 And Trim(strValProductName) = "" _
                And bIncludeDisabledSupplyingStatus = False Then

                strQueryToUse = "SELECT Employers.Name, ProductMaster.ProductID " & _
                            " FROM (ProductMaster INNER JOIN SupplierProducts " & _
                            " ON ProductMaster.ProductID = " & _
                            " SupplierProducts.ProductID) INNER JOIN " & _
                            " Employers ON " & _
                            " SupplierProducts.SupplierOrganizationID = " & _
                            " Employers.EmployerID " & _
                            " WHERE SupplierProducts.SupplyingStatus = TRUE "


                'Returns all suppliers regardless of their status 
            ElseIf lValProductID = 0 And Trim(strValProductName) = "" _
                And bIncludeDisabledSupplyingStatus = True Then

                strQueryToUse = "SELECT Employers.Name, ProductMaster.ProductID " & _
                            " FROM (ProductMaster INNER JOIN SupplierProducts " & _
                            " ON ProductMaster.ProductID = " & _
                            " SupplierProducts.ProductID) INNER JOIN " & _
                            " Employers ON " & _
                            " SupplierProducts.SupplierOrganizationID = " & _
                            " Employers.EmployerID "


                'Filters suppliers By product name but status must be true
            ElseIf lValProductID = 0 And Trim(strValProductName) <> "" And _
                bIncludeDisabledSupplyingStatus = False Then

                strQueryToUse = "SELECT Employers.Name, ProductMaster.ProductID " & _
                            " FROM (ProductMaster INNER JOIN SupplierProducts " & _
                            " ON ProductMaster.ProductID = " & _
                            " SupplierProducts.ProductID) INNER JOIN " & _
                            " Employers ON " & _
                            " SupplierProducts.SupplierOrganizationID = " & _
                            " Employers.EmployerID " & _
                            " WHERE(((ProductMaster.ProductName) = '" & _
                            strValProductName & " ')) " & _
                            " AND SupplierProducts.SupplyingStatus = TRUE "

                'Filters suppliers By product name regardless of status
            ElseIf lValProductID = 0 And Trim(strValProductName) <> "" And _
                bIncludeDisabledSupplyingStatus = True Then

                strQueryToUse = "SELECT Employers.Name, ProductMaster.ProductID " & _
                            " FROM (ProductMaster INNER JOIN SupplierProducts " & _
                            " ON ProductMaster.ProductID = " & _
                            " SupplierProducts.ProductID) INNER JOIN " & _
                            " Employers ON " & _
                            " SupplierProducts.SupplierOrganizationID = " & _
                            " Employers.EmployerID " & _
                            " WHERE(((ProductMaster.ProductName) = '" & _
                            strValProductName & " ')) "


                'Filters suppliers By Product ID but status must be true
            ElseIf lValProductID <> 0 And _
                bIncludeDisabledSupplyingStatus = False Then

                strQueryToUse = "SELECT Employers.Name, ProductMaster.ProductID " & _
                            " FROM (ProductMaster INNER JOIN SupplierProducts " & _
                            " ON ProductMaster.ProductID = " & _
                            " SupplierProducts.ProductID) INNER JOIN " & _
                            " Employers ON " & _
                            " SupplierProducts.SupplierOrganizationID = " & _
                            " Employers.EmployerID " & _
                            " WHERE SupplierProducts.SupplyingStatus = TRUE " & _
                            " And ProductMaster.ProductID = " & lValProductID


                'Filters suppliers By ProductID regardless of status
            ElseIf lValProductID <> 0 And _
                bIncludeDisabledSupplyingStatus = True Then

                strQueryToUse = "SELECT Employers.Name, ProductMaster.ProductID " & _
                            " FROM (ProductMaster INNER JOIN SupplierProducts " & _
                            " ON ProductMaster.ProductID = " & _
                            " SupplierProducts.ProductID) INNER JOIN " & _
                            " Employers ON " & _
                            " SupplierProducts.SupplierOrganizationID = " & _
                            " Employers.EmployerID " & _
                            " WHERE ProductMaster.ProductID = " & lValProductID

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

#End Region


#Region "DatabaseProcedures"

    Public Function UnSetSupplierProduct() As Boolean

        Try

            Return Update("UPDATE SupplierProducts SET " & _
                    " SupplyingStatus = FALSE " & _
                    " WHERE ProductID = " & lProductID, _
                    False, False, False, False)

        Catch ex As Exception

        End Try

    End Function

    Public Function SetSupplierProduct() As Boolean

        Try

            If UnSetSupplierProduct() = False Then
                Exit Function
            End If

            Return Update("UPDATE SupplierProducts SET " & _
                    " SupplyingStatus = TRUE " & _
                    " WHERE ProductID = " & lProductID & _
                    " AND SupplierOrganizationID = " & lSupplierOrganizationID, _
                    False, False, False, False)

        Catch ex As Exception

        End Try

    End Function

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
                lSupplierOrganizationID = 0 Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide the following details in" & _
                " order to save a Supplier's Product Price Range" & _
                Chr(10) & "1. Existing Product" & _
                Chr(10) & "2. Supplier Organization", MsgBoxStyle.Critical, _
                "iManagement - Save Action Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM SupplierProducts WHERE ProductID = " & _
                lProductID & " AND SupplierOrganizationID = " & _
                        lSupplierOrganizationID, False) = True Then

                If DisplayConfirmation = True Then
                    If MsgBox("The product is already linked to this supplier." & _
                    " Do you want to update the details?", MsgBoxStyle.YesNo + _
                            MsgBoxStyle.Exclamation, "iManagement - Record Exists") = MsgBoxResult.Yes Then

                        If bSupplyingStatus = True Then
                            Update("UPDATE SupplierProducts SET " & _
                            " SupplyingStatus = FALSE WHERE SupplyingStatus = TRUE", _
                False, False, False, False)
                        End If
                    End If


                    Update("UPDATE SupplierProducts SET " & _
                          " SupplyingStatus = " & bSupplyingStatus & _
                    " WHERE ProductID = " & lProductID & _
                            " AND SupplierOrganizationID = " & lSupplierOrganizationID, _
                False, False, False, False)

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            If DisplayConfirmation = True Then
                If MsgBox _
                    ("Do you want to save this new Supplier's Product?", _
                        MsgBoxStyle.YesNo, _
                            "iManagement - Save Record?") = _
                                MsgBoxResult.No Then

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If
            End If



            strInsertInto = "INSERT INTO SupplierProducts (" & _
                "ProductID," & _
                "SupplierOrganizationID," & _
                "SupplyingStatus" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lProductID & _
                    "," & lSupplierOrganizationID & _
                    ",FALSE)"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)


            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccess = True Then
                    MsgBox("Supplier Price Ranges Saved Successfully.", _
                        MsgBoxStyle.Information, _
                            "iManagement - Record Saved Successfully")

                End If

            Else

                If DisplayFailure = True Then
                    MsgBox("'Save Supplier Price Ranges' action failed." & _
            " Make sure all mandatory details are entered", _
            MsgBoxStyle.Exclamation, _
            "iManagement -  Addition Failed")

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
                        datRetData = Nothing
                        objLogin = Nothing
                        Return False
                        Exit Function

                    End If

                    'Whether to fill properties with values or not
                    If ReturnStatus = True Then

                        For Each myDataRows In myDataTables.Rows

                            lProductID = _
                                myDataRows("ProductID")
                            lSupplierOrganizationID = _
                                myDataRows("SupplierOrganizationID")
                            bSupplyingStatus = _
                                myDataRows("SupplyingStatus")

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


            If lProductID = 0 Or lSupplierOrganizationID = 0 _
                                Then
                MsgBox("Cannot Delete the supplier product linkage." & _
                Chr(10) & "Please select an existing Supplier and a " & _
                "product they supply.", _
                                      MsgBoxStyle.Exclamation, _
                                      "iManagement - invalid or incomplete Information")
                datDelete = Nothing
                objLogin = Nothing
                Exit Function

            End If


            If MsgBox("Are you sure you want to delete this Supplier's " & _
                "product likage?", MsgBoxStyle.YesNo, _
                    "iManagement - Delete Supplier Product Record?") = _
                        MsgBoxResult.No Then

                datDelete = Nothing
                objLogin = Nothing
                Exit Function

            End If

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            strDeleteQuery = "DELETE * FROM SupplierProducts " & _
                " WHERE SupplierOrganizationID = " & lSupplierOrganizationID & _
                    " AND ProductID = " & lProductID

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strDeleteQuery, _
            datDelete)


            strDeleteQuery = "DELETE * FROM SupplierPriceRanges " & _
                " WHERE SupplierOrganizationID = " & lSupplierOrganizationID & _
                    " AND ProductID = " & lProductID

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Supplier Price Ranges Details Deleted", _
                    MsgBoxStyle.Information, _
                        "iManagement - Record Deleted Successfully")

            Else

                MsgBox("'Delete Supplier Price Ranges' action failed", _
                    MsgBoxStyle.Exclamation, _
                        "Supplier Price Ranges Deletion failed")

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
                                "iManagement -  Supplier Product Updated")
                    End If


                End If

            End If

            datUpdated = Nothing
            objLogin = Nothing

            If bUpdateSuccess = True Then
                Return True
            End If

        Catch ex As Exception

        End Try
    End Function

#End Region


End Class

