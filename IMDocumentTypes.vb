Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMDocumentTypes

#Region "PrivateAccountingPeriodVariables"
    Private lDocumentTypeID As Long
    Private strDocumentType As String
    Private strDocumentPurpose As String
    Private lNoOfPrintOuts As Long
    Private lNoOfScanIns As Long
    Private bTypeStatus As Boolean
    Private strDefaultPrinter As String
    Private dtDateRegisterd As Date
    Private strProducedFor As String
    Private strProducedBy As String

    
#End Region

#Region "Properties"

    Public Property ProducedBy() As String

        Get
            Return strProducedBy
        End Get

        Set(ByVal Value As String)
            strProducedBy = Value
        End Set

    End Property

    Public Property ProducedFor() As String

        Get
            Return strProducedFor
        End Get

        Set(ByVal Value As String)
            strProducedFor = Value
        End Set

    End Property

    Public Property DocumentTypeID() As Long

        Get
            Return lDocumentTypeID
        End Get

        Set(ByVal Value As Long)
            lDocumentTypeID = Value
        End Set

    End Property

    Public Property DateRegisterd() As Date

        Get
            Return dtDateRegisterd
        End Get

        Set(ByVal Value As Date)
            dtDateRegisterd = Value
        End Set

    End Property

    Public Property DocumentType() As String

        Get
            Return strDocumentType
        End Get

        Set(ByVal Value As String)
            strDocumentType = Value
        End Set

    End Property

    Public Property DocumentPurpose() As String

        Get
            Return strDocumentPurpose
        End Get

        Set(ByVal Value As String)
            strDocumentPurpose = Value
        End Set

    End Property

    Public Property NoOfPrintOuts() As Long

        Get
            Return lNoOfPrintOuts
        End Get

        Set(ByVal Value As Long)
            lNoOfPrintOuts = Value
        End Set

    End Property

    Public Property NoOfScanIns() As Long

        Get
            Return lNoOfScanIns
        End Get

        Set(ByVal Value As Long)
            lNoOfScanIns = Value
        End Set

    End Property

    Public Property TypeStatus() As Boolean

        Get
            Return bTypeStatus
        End Get

        Set(ByVal Value As Boolean)
            bTypeStatus = Value
        End Set

    End Property

    Public Property DefaultPrinter() As String

        Get
            Return strDefaultPrinter
        End Get

        Set(ByVal Value As String)
            strDefaultPrinter = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Sub NewRecord()

        lDocumentTypeID = 0
        strDocumentType = ""
        strDocumentPurpose = ""

    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Sub Save()

        'Saves a new country name
        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String

            'check if periodstatus is provided
            If Trim(strDocumentType) = "" _
                Or Trim(strDocumentPurpose) = "" Then

                MsgBox("Please provide the document type and the document purpose", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

                Exit Sub

            End If


            If Find("SELECT * FROM DocumentTypes WHERE DocumentType = '" & _
                Trim(strDocumentType) & "'", True) = True Then

                If MsgBox("The document type provided exists. Do you" & _
                        Chr(10) & " want to update the detials?", _
                            MsgBoxStyle.YesNo, _
                                "iManagement - Record Exists. Update?") _
                                    = MsgBoxResult.Yes Then

                    Update("UPDATE DocumentTypes SET " & _
                    " DocumentPurpose = '" & Trim(strDocumentPurpose) & _
                    "' , NoOfPrintOuts = " & lNoOfPrintOuts & _
                    " , NoOfScanIns = " & lNoOfScanIns & _
                    " , TypeStatus = " & bTypeStatus & _
                    " , ProducedBy = '" & Trim(strProducedBy) & _
                    "' , ProducedFor = '" & Trim(strProducedFor) & _
                    "' , DefaultPrinter = '" & Trim(strDefaultPrinter) & _
                    "' WHERE DocumentTypeID = " & lDocumentTypeID)
                End If

                objLogin = Nothing
                datSaved = Nothing
                Exit Sub

            End If


            If MsgBox("Do you want to add this new Document Type?", _
                MsgBoxStyle.YesNo, "iManagement - Add New Document Type") _
                    = MsgBoxResult.No Then

                objLogin = Nothing
                datSaved = Nothing
                Exit Sub
            End If

            strInsertInto = "INSERT INTO DocumentTypes (" & _
                    "DocumentType," & _
                    "DocumentPurpose," & _
                    "NoOfPrintOuts," & _
                    "NoOfScanIns," & _
                    "TypeStatus," & _
                    "ProducedBy," & _
                    "ProducedFor," & _
                    "DefaultPrinter" & _
                        ") VALUES "

            strSaveQuery = strInsertInto & _
                        "('" & Trim(strDocumentType) & _
                        "', '" & Trim(strDocumentPurpose) & _
                        "', " & lNoOfPrintOuts & _
                        " , " & lNoOfScanIns & _
                        " , " & bTypeStatus & _
                        " , '" & Trim(strProducedBy) & _
                        "' , '" & Trim(strProducedFor) & _
                        "', '" & Trim(strDefaultPrinter) & _
                        "')"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                            strSaveQuery, _
                            datSaved)

            objLogin.CloseDb()


            If bSaveSuccess = True Then
                MsgBox("Accounting Period Record Saved Successfully", _
                    MsgBoxStyle.Information, _
                        "iManagement - Record Saved")

            Else

                MsgBox("'Save Document Type' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Save Record Failed")

            End If

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception


        End Try

    End Sub

    Public Function Find(ByVal strQuery As String, _
            ByVal bReturnDetails As Boolean) As Boolean

        Try

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
                        datRetData = Nothing
                        objLogin = Nothing
                        Return False
                        Exit Function

                    End If

                    If bReturnDetails = True Then
                        For Each myDataRows In myDataTables.Rows

                            lDocumentTypeID = myDataRows("DocumentTypeID")
                            strDocumentPurpose = myDataRows("DocumentPurpose").ToString()
                            strDocumentType = myDataRows("DocumentType").ToString()
                            strDefaultPrinter = myDataRows("DefaultPrinter").ToString()
                            lNoOfPrintOuts = myDataRows("NoOfPrintOuts")
                            lNoOfScanIns = myDataRows("NoOfScanIns")
                            bTypeStatus = myDataRows("TypeStatus")
                            dtDateRegisterd = myDataRows("DateRegistered")
                            strProducedBy = myDataRows("ProducedBy").ToString
                            strProducedFor = myDataRows("ProducedFor").ToString

                        Next
                    End If


                Next

                Return True
            Else
                Return False
            End If

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

            If lDocumentTypeID = 0 Then
                MsgBox("Cannot Delete Document Type Details. Please select an existing Document Type", _
                        MsgBoxStyle.Exclamation, _
                        "iManagement - Cannot Delete Record")

                datDelete = Nothing
                objLogin = Nothing
                Exit Sub
            End If

            If MsgBox("Do you want to delete this document type?", _
                MsgBoxStyle.YesNo, _
                    "iManagement - Delete Record?") = MsgBoxResult.No Then

                datDelete = Nothing
                objLogin = Nothing
                Exit Sub
            End If


            If Find("SELECT * FROM DocumentTypes WHERE" & _
                    " DocumentTypeID = " & lDocumentTypeID, True) = False Then

                MsgBox("The Document Type provided for deletion is does not exist." _
                    , MsgBoxStyle.Exclamation, _
                        "iManagament - invalid or incomplete informaiton")

                datDelete = Nothing
                objLogin = Nothing
                Exit Sub
            End If

            strDeleteQuery = "DELETE * FROM DocumentTypes WHERE" & _
            " DocumentTypeID = " & lDocumentTypeID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Document Type Details Record Deleted Successfully", _
                MsgBoxStyle.Information, _
                    "iManagement - Record Deleted")
            Else
                MsgBox("'Document Type Delete' action failed", _
                    MsgBoxStyle.Exclamation, _
                    " Record Deletion failed")
            End If



            datDelete = Nothing
            objLogin = Nothing


        Catch ex As Exception
            MsgBox("Error during deletion. Please contact the " & _
            "Systems Administrator." & _
                Chr(10) & "(" & ex.Message & ")", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Deletion Error")
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

            If lDocumentTypeID <> 0 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strUpdateQuery, _
                datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    MsgBox("Document Type Details Record Updated Successfully", MsgBoxStyle.Information, _
                        "iManagement - Record Updated")
                End If

            End If

        Catch ex As Exception
            MsgBox("Error occured while Updating an existing record." & _
                    Chr(10) & "Please contact the Systems Administrator." & _
                           Chr(10) & "(" & ex.Message & ")", _
                               MsgBoxStyle.Critical, _
                                   "iManagement - Deletion Error")
        End Try

    End Sub

#End Region

End Class
