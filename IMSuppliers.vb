Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMSuppliers

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
    Private lPriorityID As Long
    Private bSupplierStatus As Boolean

#End Region

#Region "Properties"

    Public Property SupplierOrgID() As Long

        Get
            Return lSupplierOrganizationID
        End Get

        Set(ByVal Value As Long)
            lSupplierOrganizationID = Value
        End Set

    End Property

    Public Property PriorityStatus() As Long

        Get
            Return lPriorityID
        End Get

        Set(ByVal Value As Long)
            lPriorityID = Value
        End Set

    End Property

    Public Property SupplierStatus() As Boolean

        Get
            Return bSupplierStatus
        End Get

        Set(ByVal Value As Boolean)
            bSupplierStatus = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function ReturnSupplierNames _
        (Optional ByVal lValSupplierID As Long = 0, _
            Optional ByVal bIncludeDisbledSuppliers As Boolean = True) As String()

        Dim strQueryToUse As String
        Dim objLogin As IMLogin = New IMLogin
        Dim arItems() As String

        Try

            If lValSupplierID = 0 Then
                If bIncludeDisbledSuppliers = False Then
                    strQueryToUse = "SELECT Name FROM Employers " & _
                        " INNER JOIN Suppliers ON " & _
                        " Suppliers.SupplierOrganizationID = " & _
                        " Employers.EmployerID " & _
                        " WHERE Suppliers.SupplierStatus = TRUE "

                Else
                    strQueryToUse = "SELECT Name FROM Employers " & _
                        " INNER JOIN Suppliers ON " & _
                        " Suppliers.SupplierOrganizationID = " & _
                        " Employers.EmployerID "

                End If


            ElseIf lValSupplierID <> 0 Then
                strQueryToUse = "SELECT Name FROM Employers " & _
                    " INNER JOIN Suppliers ON " & _
                    " Suppliers.SupplierOrganizationID = " & _
                    " Employers.EmployerID " & _
                    " WHERE Suppliers.SupplierOrganizationID = " & lValSupplierID


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

    Public Function ReturnDefaultOrderSent _
        (Optional ByVal lValSupplierID As Long = 0, _
            Optional ByVal bIncludeDisbledSuppliers As Boolean = True) As String()

        Dim strQueryToUse As String
        Dim objLogin As IMLogin = New IMLogin
        Dim arItems() As String

        Try

            If lValSupplierID = 0 Then
                If bIncludeDisbledSuppliers = False Then
                    strQueryToUse = "SELECT Name FROM Employers " & _
                        " INNER JOIN Suppliers ON " & _
                        " Suppliers.SupplierOrganizationID = " & _
                        " Employers.EmployerID " & _
                        " WHERE Supplier.SupplierStatus = TRUE "

                Else
                    strQueryToUse = "SELECT Name FROM Employers " & _
                        " INNER JOIN Suppliers ON " & _
                        " Suppliers.SupplierOrganizationID = " & _
                        " Employers.EmployerID "

                End If


            ElseIf lValSupplierID <> 0 Then
                strQueryToUse = "SELECT Name FROM Employers " & _
                    " INNER JOIN Suppliers ON " & _
                    " Suppliers.SupplierOrganizationID = " & _
                    " Employers.EmployerID " & _
                    " WHERE Supplier.SupplierOrganizationID = " & lValSupplierID


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


            If lSupplierOrganizationID = 0 Then

                MsgBox("You must provide an appropriate Supplier's Organization." _
                , MsgBoxStyle.Critical, _
                "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing
                Exit Function

            End If


            'Check if there is an existing supplier with this id
            If Find("SELECT * FROM Suppliers WHERE SupplierOrganizationID = " _
                & lSupplierOrganizationID, _
                    False) = True Then

                If bDisplayConfirmMessages = False Then
                    If MsgBox("This supplier has already been added and duplicate supplier entries are not allowed.", _
                        MsgBoxStyle.YesNo, _
                            "iManagement - Record Exists") = _
                                MsgBoxResult.No Then

                        objLogin = Nothing
                        datSaved = Nothing
                        Exit Function

                    End If
                End If


                Update("UPDATE Suppliers SET" & _
                    " PriorityID = " & lPriorityID _
                        & " , SupplierStatus = " & bSupplierStatus & _
                            " WHERE SupplierOrganizationID = " _
                                & lSupplierOrganizationID, _
                                    bDisplayErrorMessages, _
                                        bDisplaySuccessMessages, _
                                            bDisplayFailureMessages, _
                                                bDisplayConfirmMessages)



                objLogin = Nothing
                datSaved = Nothing
                Exit Function
            End If

            If bDisplayConfirmMessages = True Then
                If MsgBox("Are you sure you want to add this " & _
                "Organization as a member of your authorised possible Suppliers?", _
                    MsgBoxStyle.YesNo, _
                    "iManagement - Add Record") = MsgBoxResult.No Then

                    objLogin = Nothing
                    datSaved = Nothing
                    Exit Function
                End If

            End If
            strInsertInto = "INSERT INTO Suppliers (" & _
                "SupplierOrganizationID," & _
                "PriorityID," & _
                "SupplierStatus" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lSupplierOrganizationID & _
                    "," & lPriorityID & _
                    "," & bSupplierStatus & _
                    ")"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If bDisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - New Supplier Saved")

                End If
                Return True

            Else

                If bDisplayFailureMessages = True Then
                    MsgBox("'Save New Supplier' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save New Supplier Failed")
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

                        lSupplierOrganizationID = _
                                myDataRows("SupplierOrganizationID")
                        lPriorityID = _
                                myDataRows("PriorityID")
                        bSupplierStatus = _
                                myDataRows("SupplierStatus")

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

            If lSupplierOrganizationID = 0 Then
                MsgBox("Cannot Delete the existing supplier. Please " & _
                "select an existing Supplier.", _
                                    MsgBoxStyle.Exclamation, _
                "iManagement - invalid or incomplete information")

                datDelete = Nothing
                objLogin = Nothing
                Exit Function

            End If


            If MsgBox("Are you sure you want to delete the Supplier's details, " & _
            "inclusive of " & Chr(10) & _
            "the price ranges for the supplier's products?", _
                MsgBoxStyle.YesNo, _
                    "iManagement - Delete Record?") = MsgBoxResult.No Then

                datDelete = Nothing
                objLogin = Nothing
                Exit Function
            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            'Delete Suppliers
            strDeleteQuery = "DELETE * FROM Suppliers " & _
            " WHERE SupplierOrganizationID = " & lSupplierOrganizationID

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strDeleteQuery, _
                    datDelete)

            'Delete Supplier Products
            strDeleteQuery = "DELETE * FROM SupplierProducts " & _
            " WHERE SupplierOrganizationID = " & lSupplierOrganizationID

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, _
                datDelete)

            'Delete Price Ranges
            strDeleteQuery = "DELETE * FROM SupplierPriceRanges " & _
            " WHERE SupplierOrganizationID = " & lSupplierOrganizationID

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, _
                datDelete)



            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Supplier Deleted")
                Return True

            Else
                MsgBox("'Supplier delete' action failed", _
                    MsgBoxStyle.Exclamation, "Supplier Deletion failed")
            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function Update(ByVal strUpQuery As String, _
        ByVal bDisplayErrorMessages As Boolean, _
            ByVal bDisplaySuccessMessages As Boolean, _
                ByVal bDisplayFailureMessages As Boolean, _
                    ByVal bDisplayConfirmMessages As Boolean) As Boolean

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
                If bDisplayConfirmMessages = True Then
                    MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                        "iManagement -  Supplier Details Updated")
                End If

                Return True
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Function

#End Region

End Class


