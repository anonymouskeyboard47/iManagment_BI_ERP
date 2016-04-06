
Option Explicit On 
'Option Strict On
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMTransactionMaster


#Region "PrivateVariables"

    Private lTransactionID As Long
    Private lSequenceGroupID As Long
    Private dtTransactionDate As Date
    Private dtTransactionBookDate As Date
    Private lCostCentreID As Long
    Private strTransactionDetails As String
    Private lContraEntryTransactionID As Long
    Private strContraEntryPurpose As String

#End Region


#Region "Properties"

    Public Property TransactionID() As Long

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return lTransactionID
        End Get

        Set(ByVal Value As Long)
            lTransactionID = Value
        End Set

    End Property

    Public Property SequenceGroupID() As Long

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return lSequenceGroupID
        End Get

        Set(ByVal Value As Long)
            lSequenceGroupID = Value
        End Set

    End Property

    Public Property TransactionDate() As Date

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return dtTransactionDate
        End Get

        Set(ByVal Value As Date)
            dtTransactionDate = Value
        End Set

    End Property

    Public Property TransactionBookDate() As Date

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return dtTransactionBookDate
        End Get

        Set(ByVal Value As Date)
            dtTransactionBookDate = Value
        End Set

    End Property


    Public Property ContraEntryTransactionID() As Long

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return lContraEntryTransactionID
        End Get

        Set(ByVal Value As Long)
            lContraEntryTransactionID = Value
        End Set

    End Property

    Public Property ContraEntryPurpose() As String

        'USED TO SET AND RETRIEVE THE POST COUNTRY CODE (STRING)
        Get
            Return strContraEntryPurpose
        End Get

        Set(ByVal Value As String)
            strContraEntryPurpose = Value
        End Set

    End Property

    Public Property CostCentreID() As Long

        'USED TO SET AND RETRIEVE THE BANK ID (STRING)
        Get
            Return lCostCentreID
        End Get

        Set(ByVal Value As Long)
            lCostCentreID = Value
        End Set

    End Property

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "DatabaseProcedures"

    Public Function SaveMainDetails(ByVal bDisplayErrorMessages As Boolean, _
        ByVal bDisplayConfirmation As Boolean, _
            ByVal bDisplayFailure As Boolean, _
                ByVal bDisplaySuccess As Boolean) As Boolean

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String
        Dim i As Long


        Try

            If lTransactionID = 0 Or lSequenceGroupID = 0 Then
                ReturnError += "This transaction cannot be saved due " & _
                    "to invalid data"

                datSaved = Nothing
                objLogin = Nothing

                Exit Function
            End If

            'If bDisplayConfirmation = True Then
            '    If MsgBox("Do you want to add this new Transaction?", _
            '        MsgBoxStyle.YesNo, _
            '            "iManagement - Add Transaction Details?") _
            '                = MsgBoxResult.No Then

            '        objLogin = Nothing
            '        datSaved = Nothing

            '        Exit Function

            '    End If
            'End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            strInsertInto = "INSERT INTO TransactionMaster (" & _
                "TransactionID," & _
                "SequenceGroupID," & _
                "TransactionDate," & _
                "TransactionBookDate," & _
                "TransactionDetails," & _
                "CostCentreID," & _
                "ContraEntryTransaction," & _
                "ContraEntryPurpose" & _
                    ") VALUES "


            strSaveQuery = strInsertInto & _
                    "(" & lTransactionID & _
                    "," & lSequenceGroupID & _
                    ",#" & dtTransactionDate & _
                    "#,#" & dtTransactionBookDate & _
                    "#,'" & strTransactionDetails & _
                    "'," & CostCentreID & _
                    "," & lContraEntryTransactionID & _
                    ",'" & strContraEntryPurpose & _
                            "')"

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, strSaveQuery, datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If bDisplaySuccess = True Then

                    ReturnError += " Main Transaction Saved Successfully."

                End If
            Else

                If bDisplayFailure = True Then

                    ReturnError += "'Save Main Transaction' action failed." & _
              " Make sure all mandatory details are entered."

                End If
            End If


            objLogin = Nothing
            datSaved = Nothing

            If bSaveSuccess = True Then
                Return True
            End If


        Catch ex As Exception
            If bDisplayErrorMessages = True Then
                ReturnError += ex.Message.ToString

            End If
        End Try

    End Function

    Public Function FindMainDetails(ByVal strQuery As String, _
                        ByVal ReturnStatus As Boolean) As Boolean
        'Query must contain at least rows from Sequence

        Try

            Dim datRetData As DataSet = New DataSet
            Dim bQuerySuccess As Boolean
            Dim myDataTables As DataTable
            Dim myDataColumns As DataColumn
            Dim myDataRows As DataRow
            Dim objLogin As IMLogin = New IMLogin
            Dim i As Long

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

                            lTransactionID = _
                                    myDataRows("TransactionID")
                            dtTransactionDate = _
                                    myDataRows("TransactionDate")
                            lSequenceGroupID = _
                                    myDataRows("SequenceGroupID")
                            lCostCentreID = _
                                    myDataRows("CostCentreID")
                            dtTransactionBookDate = _
                                    myDataRows("TransactionBookDate")
                            lContraEntryTransactionID  = _
                                    myDataRows("lContraEntryTransactionID")
                            strTransactionDetails = _
                                    myDataRows("TransactionDetails").ToString
                            

                        Next
                    End If
                Next

                datRetData = Nothing
                objLogin = Nothing

                Return True

            Else

                datRetData = Nothing
                objLogin = Nothing

                Return False

            End If

        Catch ex As Exception
            returnerror += ex.Message.ToString

        End Try

    End Function

    Public Function DeleteMainDetails(ByVal strDelQuery As String) As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = strDelQuery

            If TransactionID = 0 Then
                ReturnError += "Cannot Delete. Please select an " & _
                    "existing Transaction Detail."

                objLogin = Nothing
                datDelete = Nothing
                Exit Function

            End If


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                ReturnError += "Transaction Details Deleted"

                datDelete = Nothing
                objLogin = Nothing
                Return True
            Else

                returnerror += "'Delete Transaction action failed"


            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Function


#End Region


End Class
