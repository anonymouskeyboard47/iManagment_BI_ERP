Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMInvoiceTypes

#Region "PrivateVariables"

    Private strInvoiceType As String
    Private strInvoiceCategory As String
    Private strInvoiceTypeSeries As String
    Private bIsDebit As Boolean
    Private strDefaultText As String
    Private lDefaultShippingEmployerID As Long
    Private bNeedsApproval As Boolean
    Private bAdhereToDefaultText As Boolean

#End Region

#Region "Properties"

    Public Property InvoiceType() As String

        Get
            Return strInvoiceType
        End Get

        Set(ByVal Value As String)
            strInvoiceType = Value
        End Set

    End Property

    Public Property InvoiceCategory() As String

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return strInvoiceCategory
        End Get

        Set(ByVal Value As String)
            strInvoiceCategory = Value
        End Set

    End Property

    Public Property InvoiceTypeSeries() As String

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return strInvoiceTypeSeries
        End Get

        Set(ByVal Value As String)
            strInvoiceTypeSeries = Value
        End Set

    End Property

    Public Property IsDebit() As Boolean

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return bIsDebit
        End Get

        Set(ByVal Value As Boolean)
            bIsDebit = Value
        End Set

    End Property

    Public Property DefaultText() As String

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return strDefaultText
        End Get

        Set(ByVal Value As String)
            strDefaultText = Value
        End Set

    End Property

    Public Property AdhereToDefaultText() As Boolean

        Get
            Return bAdhereToDefaultText
        End Get

        Set(ByVal Value As Boolean)
            bAdhereToDefaultText = Value
        End Set

    End Property

    Public Property NeedsApproval() As Boolean

        Get
            Return bNeedsApproval
        End Get

        Set(ByVal Value As Boolean)
            bNeedsApproval = Value
        End Set

    End Property

    Public Property DefaultShippingEmployerID() As Long

        Get
            Return lDefaultShippingEmployerID
        End Get

        Set(ByVal Value As Long)
            lDefaultShippingEmployerID = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Sub Save(ByVal DisplayMessages As Boolean)
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        With objLogin

            Try

                If Trim(strInvoiceType) = "" Or _
                    Trim(strInvoiceCategory) = "" Or _
                        Trim(strInvoiceTypeSeries) = "" Or _
                            Trim(strDefaultText) = "" _
                                    Then

                    MsgBox("Cannot save the Invoice Type. Missing information", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - invalid or incomplete Information")

                    datSaved = Nothing
                    objLogin = Nothing
                    Exit Sub

                End If

                'Make sure the invoice category posted 
                If Trim(strInvoiceCategory) = "Product Invoice" And _
                       Trim(strInvoiceCategory) = "General Invoice" _
                                   Then

                    MsgBox("Cannot Save details. Please provide an appropriate Invoice Category.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - invalid or incomplete information")

                    datSaved = Nothing
                    objLogin = Nothing
                    Exit Sub

                End If

                'Check if the invoice type exists
                If Find("SELECT * FROM InvoiceTypes WHERE InvoiceType = '" & _
                strInvoiceType & "'", False) = True Then

                    If MsgBox("This invoice type already exists. Do you want to update its details?", _
                        MsgBoxStyle.YesNo, _
                            "IManagement - Record Exists. Update it?") _
                            = MsgBoxResult.Yes Then

                        Update("Update InvoiceTypes SET " & _
                        " InvoiceCategory = '" & Trim(strInvoiceCategory) & _
                        "' , InvoiceTypeSeries = '" & Trim(strInvoiceTypeSeries) & _
                        "' , DefaultText = '" & Trim(strDefaultText) & _
                        "' , IsDebit = " & bIsDebit & _
                        " , AdhereToDefaultText = " & Trim(bAdhereToDefaultText) & _
                        " , NeedsApproval = " & bNeedsApproval & _
                        " , DefaultShippingEmployerID = " & _
                            lDefaultShippingEmployerID & _
                       " WHERE InvoiceType = '" & strInvoiceType & "'")

                    End If

                    datSaved = Nothing
                    objLogin = Nothing
                    Exit Sub

                End If

                If MsgBox(" Are you sure you want to Save this Invoice Type?", _
                                     MsgBoxStyle.YesNo, _
                                     "iManagement - Delete Record?") = _
                                      MsgBoxResult.No Then

                    datSaved = Nothing
                    objLogin = Nothing
                    Exit Sub
                End If


                strInsertInto = "INSERT INTO InvoiceTypes (" & _
                        "InvoiceType," & _
                        "InvoiceCategory," & _
                        "InvoiceTypeSeries," & _
                        "DefaultText," & _
                        "IsDebit," & _
                        "AdhereToDefaultText," & _
                        "NeedsApproval," & _
                        "DefaultShippingEmployerID" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                            "(" & _
                        "'" & Trim(strInvoiceType) & _
                        "', '" & Trim(strInvoiceCategory) & _
                        "', '" & Trim(strInvoiceTypeSeries) & _
                        "', '" & Trim(strDefaultText) & _
                        "', " & bIsDebit & _
                        ", " & bAdhereToDefaultText & _
                        ", " & bNeedsApproval & _
                        ", " & lDefaultShippingEmployerID & _
                            ")"

                .ConnectString = strOrgAccessConnString
                .ConnectToDatabase()

                bSaveSuccess = .ExecuteQuery(strOrgAccessConnString, _
                                    strSaveQuery, _
                                            datSaved)

                If bSaveSuccess = True And DisplayMessages = True Then
                    MsgBox("Record Saved Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement - Invoice Type Details Saved")

                ElseIf bSaveSuccess = False And DisplayMessages = True Then

                    MsgBox("'Save Invoice Type' action failed." & _
                        " Make sure all mandatory details are entered", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Invoice Type Details Failed")

                    '.RollbackTheTrans()

                End If

                .CloseDb()

            Catch ex As Exception

                'If Not objLogin Is Nothing Then
                '    .RollbackTheTrans()
                'End If

            End Try

        End With

    End Sub

    Public Function Find(ByVal strQuery As String, _
            ByVal bReturnValue As Boolean) As Boolean

        Try


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
                objLogin = Nothing
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


                    For Each myDataRows In myDataTables.Rows

                        If bReturnValue = True Then
                            strInvoiceType = myDataRows("InvoiceType").ToString()
                            strInvoiceTypeSeries = myDataRows("InvoiceTypeSeries").ToString()
                            strInvoiceCategory = myDataRows("InvoiceCategory").ToString()
                            strDefaultText = myDataRows("DefaultText").ToString()
                            bIsDebit = myDataRows("IsDebit")
                            bAdhereToDefaultText = myDataRows("AdhereToDefaultText")
                            bNeedsApproval = myDataRows("NeedsApproval")
                            lDefaultShippingEmployerID = myDataRows _
                                ("DefaultShippingEmployerID")
                        End If
                    Next

                Next

                Return True
            Else
                Return False
            End If

            datRetData = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Sub Delete()

        Try

            'Deletes the country details of the country with the country code
            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin



            If Trim(strInvoiceType) = "" _
                        Then
                MsgBox("Cannot Delete due to missing information", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

                datDelete = Nothing
                objLogin = Nothing
                Exit Sub

            End If

            strDeleteQuery = "DELETE * FROM InvoiceTypes WHERE InvoiceType = '" & _
            strInvoiceType & "'"

            If MsgBox(" Are you sure you want to Delete this Invoice Type?", _
                                                MsgBoxStyle.YesNo, _
                                                "iManagement - Delete Record?") = _
                                                 MsgBoxResult.No Then

                datDelete = Nothing
                objLogin = Nothing
                Exit Sub
            End If


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strDeleteQuery, _
                            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Invoice Type Details Deleted")
            Else
                MsgBox("'Invoice Type delete' action failed", _
                    MsgBoxStyle.Exclamation, " Invoice Type Deletion failed")
            End If

            datDelete = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Update(ByVal strUpQuery As String)
        'Updates country details of the country with the country code

        Try


            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If Trim(strInvoiceType) = "" _
                         Then
                MsgBox("Cannot Update due to missing information.", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

                datUpdated = Nothing
                objLogin = Nothing
                Exit Sub

            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
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

            datUpdated = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try


    End Sub

#End Region


End Class

