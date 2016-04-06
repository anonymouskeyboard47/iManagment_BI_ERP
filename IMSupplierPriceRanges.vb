Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMSupplierPriceRanges

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
    Private lRangeID As Double
    Private dbPricePerUnit As Decimal
    Private dbMinNumOfUnits As Double
    Private dtCommencementDate As Date
    Private dtDateCreated As Date
    Private dtExpiryDate As Date
    Private bPriceRangeStatus As Boolean

#End Region

#Region "Properties"

    Public Property PriceRangeStatus() As Boolean

        Get
            Return bPriceRangeStatus
        End Get

        Set(ByVal Value As Boolean)
            bPriceRangeStatus = Value
        End Set

    End Property

    Public Property SupplierOrganizationID() As Long

        Get
            Return lSupplierOrganizationID
        End Get

        Set(ByVal Value As Long)
            lSupplierOrganizationID = Value
        End Set

    End Property

    Public Property RangeID() As Double

        Get
            Return lRangeID
        End Get

        Set(ByVal Value As Double)
            lRangeID = Value
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

    Public Property PricePerUnit() As Decimal

        Get
            Return dbPricePerUnit
        End Get

        Set(ByVal Value As Decimal)
            dbPricePerUnit = Value
        End Set

    End Property

    Public Property MinNumOfUnits() As Double

        Get
            Return dbMinNumOfUnits
        End Get

        Set(ByVal Value As Double)
            dbMinNumOfUnits = Value
        End Set

    End Property

    Public Property DateEntered() As Date

        Get
            Return dtDateCreated
        End Get

        Set(ByVal Value As Date)
            dtDateCreated = Value
        End Set

    End Property

    Public Property ExpiryDate() As Date

        Get
            Return dtExpiryDate
        End Get

        Set(ByVal Value As Date)
            dtExpiryDate = Value
        End Set

    End Property

    Public Property CommencementDate() As Date

        Get
            Return dtCommencementDate
        End Get

        Set(ByVal Value As Date)
            dtCommencementDate = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

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
                lSupplierOrganizationID = 0 Or _
                    dbPricePerUnit = 0 Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide the following details in" & _
                " order to save a Supplier's Product Price Range" & _
                Chr(10) & "1.Existing Product" & _
                Chr(10) & "2.Supplier Organization" & _
                Chr(10) & "3.Price Per Unit", _
                MsgBoxStyle.Critical, _
            "iManagement - Save Action Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM SupplierPriceRanges WHERE RangeID = " & _
            lRangeID, False) = True Then

                If MsgBox("The Supplier's Price Ranges Details for this product already exists." & _
                    Chr(10) & "Do you want to update the details?", _
                            MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                    MsgBoxResult.Yes Then

                    Update("UPDATE SupplierPriceRanges SET " & _
                                " ProductID = " & lProductID & _
                                " , SupplierOrganizationID = " & lSupplierOrganizationID & _
                                " , PricePerUnit = " & dbPricePerUnit & _
                                " , MinNumOfUnits = " & dbMinNumOfUnits & _
                                " , CommencementDate = #" & dtCommencementDate & _
                                "# , ExpiryDate = #" & dtExpiryDate & _
                                "# , PriceRangeStatus = " & bPriceRangeStatus & _
                                    " WHERE  RangeID = " & lRangeID)

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            If MsgBox("Do you want to save this new Supplier's Price Range?", _
                    MsgBoxStyle.YesNo, _
                        "iManagement - Save Record?") = _
                            MsgBoxResult.No Then

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            strInsertInto = "INSERT INTO SupplierPriceRanges (" & _
                "ProductID," & _
                "SupplierOrganizationID," & _
                "PricePerUnit," & _
                "MinNumOfUnits," & _
                "CommencementDate," & _
                "ExpiryDate," & _
                "PriceRangeStatus" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lProductID & _
                    "," & lSupplierOrganizationID & _
                    "," & dbPricePerUnit & _
                    "," & dbMinNumOfUnits & _
                    ",#" & dtCommencementDate & _
                    "#,#" & dtExpiryDate & _
                    "#," & bPriceRangeStatus & _
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
                            lRangeID = _
                                myDataRows("RangeID")
                            dbPricePerUnit = _
                                myDataRows("PricePerUnit")
                            dtDateCreated = _
                                myDataRows("DateCreated")
                            dtCommencementDate = _
                                myDataRows("CommencementDate")
                            dtExpiryDate = _
                                myDataRows("ExpiryDate")
                            dbMinNumOfUnits = _
                               myDataRows("MinNumOfUnits")
                            bPriceRangeStatus = _
                               myDataRows("PriceRangeStatus")

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

    Public Function Delete() As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin


            If lRangeID = 0 Then

                MsgBox("Cannot Delete Record. Please select an existing Supplier Price Ranges Detail", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - invalid or incomplete Information")

                datDelete = Nothing
                objLogin = Nothing
                Exit Function

            End If

            strDeleteQuery = "DELETE * FROM SupplierPriceRanges" & _
                " WHERE RangeID = " & lRangeID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, _
            datDelete)



            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Supplier Price Ranges Details Deleted", _
                MsgBoxStyle.Information, _
                    "iManagement - Record Deleted Successfully")
                Return True

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

    Public Sub Update(ByVal strUpQuery As String)

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
                    MsgBox("Record Updated Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement -  Supplier Price Ranges Updated")
                End If

            End If

        Catch ex As Exception

        End Try
    End Sub


#End Region

End Class
