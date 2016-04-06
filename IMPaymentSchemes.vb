Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMPaymentSchemes

#Region "PrivatePaySchemesVariables"

    Private lPaymentSchemeID As Long
    Private strPaymentSchemeTitle As String
    Private strFrequency As String
    Private strModeOfPayment As String
    Private strPaySchemeDescription As String

#End Region

#Region "Properties"

    Public Property PaymentSchemeID() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lPaymentSchemeID
        End Get

        Set(ByVal Value As Long)
            lPaymentSchemeID = Value
        End Set

    End Property

    Public Property PaymentSchemeTitle() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPaymentSchemeTitle
        End Get

        Set(ByVal Value As String)
            strPaymentSchemeTitle = Value
        End Set

    End Property

    Public Property PaymentSchemeDescription() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPaySchemeDescription
        End Get

        Set(ByVal Value As String)
            strPaySchemeDescription = Value
        End Set

    End Property


    Public Property Frequency() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeDescription (STRING)
        Get
            Return strFrequency
        End Get

        Set(ByVal Value As String)
            strFrequency = Value
        End Set

    End Property
    Public Property ModeOfPayment() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeDescription (STRING)
        Get
            Return strModeOfPayment
        End Get

        Set(ByVal Value As String)
            strModeOfPayment = Value
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
        lPaymentSchemeID = 0
        strPaymentSchemeTitle = ""
        strFrequency = ""
        strModeOfPayment = ""
        strPaySchemeDescription = ""
    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Sub Save()
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        If Trim(strPaymentSchemeTitle) <> "" Or _
                   Trim(strFrequency) <> "" Or _
                    Trim(strModeOfPayment) <> "" Then

            strInsertInto = "INSERT INTO PaymentScheme (" & _
                "PaymentSchemeTitle," & _
                "SchemeDescription," & _
                "Frequency," & _
                "ModeOfPayment)" & _
                " VALUES "

            strSaveQuery = strInsertInto & _
                        "('" & strPaymentSchemeTitle & _
                        "', '" & strPaySchemeDescription & _
                        "', '" & strFrequency & _
                        "', '" & strModeOfPayment & _
                        "')"

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Payment Scheme Saved")

            Else

                MsgBox("'Save Payment Scheme' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Save Payment Schme Failed")

            End If

        End If

    End Sub

    Public Function Find(ByVal strQuery As String) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.connectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bQuerySuccess = objLogin.ExecuteQuery(strAccessConnString, strQuery, _
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


                For Each myDataRows In myDataTables.Rows

                    lPaymentSchemeID = myDataRows _
                            ("PaymentSchemeID")

                    strPaySchemeDescription = myDataRows _
                            ("SchemeDescription").ToString()

                    strPaymentSchemeTitle = myDataRows _
                            ("PaymentSchemeTitle").ToString()

                    strFrequency = myDataRows _
                            ("Frequency").ToString()

                    strModeOfPayment = myDataRows _
                            ("ModeOfPayment").ToString()

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strDeleteQuery = strDelQuery

        If Trim(strPaymentSchemeTitle) <> "" Or _
                  Trim(strFrequency) <> "" Or _
                    Trim(strModeOfPayment) <> "" Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                strDeleteQuery, _
                        datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Payment Scheme Lookup Details Deleted")
            Else
                MsgBox("'Payment Scheme ' action failed", _
                    MsgBoxStyle.Exclamation, " Payment Scheme Deletion failed")
            End If
        Else
            MsgBox("Cannot Delete. Please select an existing Payment Scheme", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

        End If

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If Trim(strPaymentSchemeTitle) <> "" Or _
                  Trim(strFrequency) <> "" Or _
                    Trim(strModeOfPayment) <> "" Then


            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Lookup Details Updated")
            End If

        End If

    End Sub

    Public Function FillControl(ByVal strFillConnString As String, _
                ByVal strTSQL As String, ByVal strValueField As String, _
                    ByVal strTextField As String) As String()

        Dim datFillData As DataSet
        Dim bReturnedSuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim strTextFieldData() As String
        Dim i As Integer
        Dim objLogin As IMLogin = New IMLogin

        Try

            datFillData = New DataSet

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            'The db is okay now get the recordset
            bReturnedSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                strTSQL, datFillData)

            objLogin.CloseDb()

            If datFillData Is Nothing Then
                Exit Function
            End If

            For Each myDataTables In datFillData.Tables

                'Check if there is any data. If not exit
                If myDataTables.Rows.Count = 0 Then
                    'Return an empty array
                    ReDim strTextFieldData(1)
                    strTextFieldData(0) = ""
                    Return strTextFieldData

                    Exit Function
                Else
                    'Resize the array
                    ReDim strTextFieldData(myDataTables.Rows.Count)

                End If

                i = 0
                For Each myDataRows In myDataTables.Rows
                    strTextFieldData(i) = myDataRows(0).ToString()
                    i = i + 1
                Next

            Next

            Return strTextFieldData
            datFillData.Dispose()

        Catch ex As Exception

        End Try

    End Function


#End Region

End Class
