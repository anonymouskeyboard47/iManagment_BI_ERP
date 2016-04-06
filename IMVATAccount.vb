Option Explicit On 
'Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMVATAccount

    Private lVATID As Long
    Private strVATTypeID As String
    Private dbVATPercentage As Decimal
    Private lVATAccount As Long
    Private lVATAnnualIntervals As Long
    Private bVATStatus As Boolean

#Region "Properties"

    Public Property VATID() As Long

        Get
            Return lVATID
        End Get

        Set(ByVal Value As Long)
            lVATID = Value
        End Set

    End Property

    Public Property VATTypeID() As String

        Get
            Return strVATTypeID
        End Get

        Set(ByVal Value As String)
            strVATTypeID = Value
        End Set

    End Property

    Public Property VATPercentage() As Decimal

        Get
            Return dbVATPercentage
        End Get

        Set(ByVal Value As Decimal)
            dbVATPercentage = Value
        End Set

    End Property

    Public Property VATAccount() As Long

        Get
            Return lVATAccount
        End Get

        Set(ByVal Value As Long)
            lVATAccount = Value
        End Set

    End Property

    Public Property VATAnnualIntervals() As Long

        Get
            Return lVATAnnualIntervals
        End Get

        Set(ByVal Value As Long)
            lVATAnnualIntervals = Value
        End Set

    End Property

    Public Property VATStatus() As Boolean

        Get
            Return bVATStatus
        End Get

        Set(ByVal Value As Boolean)
            bVATStatus = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "DatabaseProcedures"

    'Save informaiton
    Public Sub Save(ByVal DisplayErrorMessages As Boolean, _
            ByVal DisplaySuccessMessages As Boolean, _
                ByVal DisplayFailureMessages As Boolean)

        'Saves a new base organization
        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String

            If Trim(strOrganizationName) = "" Then

                MsgBox("Please open an existing company.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Select an existing company")
                objLogin = Nothing
                datSaved = Nothing

                Exit Sub
            End If

            'Check if VAT text, Percentage, and VAT Account is provided
            If Trim(strVATTypeID) = "" Or dbVATPercentage = 0 Or lVATAccount = 0 Then

                MsgBox("You must provide a valid VAT Name, its Percentage value, and" & _
                    Chr(10) & " the Chart of Account it is related to." _
                        , MsgBoxStyle.Critical, _
                            "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub
            End If

            'Check if COA exists
            If Find("SELECT * FROM COA WHERE COAAccountNr = " _
                & Trim(lVATAccount), _
                    False) = False Then

                MsgBox("The Chart Of Account provided does not exist", _
                    MsgBoxStyle.Critical, _
                        "iManagement - invalid or incomplete informaiton")

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub

            End If


            'Check if there is an existing format with this name
            If Find("SELECT * FROM VATCodes WHERE VATTypeID = '" _
                & Trim(strVATTypeID) & "'", _
                    False) = True Then



                If MsgBox("The VAT Type you have provided exists." & _
                        Chr(10) & " Do you want to update this VAT Code's details?", _
                            MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, _
                                "iManagement - Set this format as the default?") _
                                    = MsgBoxResult.Yes Then

                    'Check if there is an existing format with this name
                    If Find("SELECT * FROM VATCodes WHERE VATTypeID = '" _
                        & Trim(strVATTypeID) & "' AND VATStatus = FALSE", _
                            False) = True Then

                        If bVATStatus = False Then
                            MsgBox("The current VAT Code is disabled. Enable it in order to update it", _
                                MsgBoxStyle.Critical, _
                                    "iManagement - invalid or incomplete information")

                        End If
                    End If


                    Update("UPDATE VATCodes SET " & _
                        "VATPercentage = " & dbVATPercentage & _
                            "VATAccount = " & lVATAccount & _
                                "VATAnnualIntervals = " & lVATAnnualIntervals & _
                                    "VATPercentage = " & False & _
                                        " WHERE VATTypeID = '" & VATTypeID & "'")


                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub
            End If


            strInsertInto = "INSERT INTO VATCodes (" & _
                " VATTypeID ," & _
                ", VATPercentage," & _
                ", VATIntervals," & _
                ", VATStatus" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & strVATTypeID & _
                    "'," & dbVATPercentage & _
                     "," & lVATAnnualIntervals & _
                      "," & bVATStatus & _
                    ")"

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement - VAT Code Saved")

                End If
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save New VAT Code' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save New VAT Code Failed")
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

    End Sub

    'Find Informaiton
    Public Function Find(ByVal strQuery As String, _
        ByVal bReturnValues As Boolean) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.connectString = strOrgAccessConnString
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

                        lVATID = _
                                myDataRows("VATID")
                        strVATTypeID = _
                                myDataRows("VATTypeID").ToString
                        lVATAccount = _
                            myDataRows("VATAccount")
                        lVATAnnualIntervals = _
                            myDataRows("VATAnnualIntervals")
                        dbVATPercentage = _
                            myDataRows("VATPercentage")
                        bVATStatus = _
                            myDataRows("VATStatus")

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
    Public Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try


            strDeleteQuery = strDelQuery

            If lVATID = 0 Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - VAT Code Deleted")
                Else
                    MsgBox("'Chart Of Account Format delete' action failed", _
                        MsgBoxStyle.Exclamation, " VAT Code Deletion failed")
                End If
            Else

                MsgBox("Cannot Delete. Please select an existing VAT Code.", _
                        MsgBoxStyle.Exclamation, "iManagement - invalid or incomplete information")

            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  VAT Code Details Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


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

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            'The db is okay now get the recordset
            bReturnedSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
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


            objLogin = Nothing
            datFillData = Nothing

            Return strTextFieldData
            datFillData.Dispose()

        Catch ex As Exception

        End Try

    End Function

#End Region


End Class
