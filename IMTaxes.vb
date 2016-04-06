Option Explicit On 
'Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections


Public Class IMTaxes

    Private lTaxID As Long
    Private strTaxName As String
    Private dbTaxPercentage As Double
    Private lTaxAccount As Long
    Private lTaxAnnualIntervals As Long
    Private strTaxCategory As String
    Private bTaxStatus As Boolean
    Private strTaxCodeID As String

#Region "Properties"

    Public Property TaxCodeID() As String

        Get
            Return strTaxCodeID
        End Get

        Set(ByVal Value As String)
            strTaxCodeID = Value
        End Set

    End Property

    Public Property TaxCategory() As String

        Get
            Return strTaxCategory
        End Get

        Set(ByVal Value As String)
            strTaxCategory = Value
        End Set

    End Property


    Public Property TaxID() As Long

        Get
            Return lTaxID
        End Get

        Set(ByVal Value As Long)
            lTaxID = Value
        End Set

    End Property

    Public Property TaxTypeID() As String

        Get
            Return strTaxName
        End Get

        Set(ByVal Value As String)
            strTaxName = Value
        End Set

    End Property

    Public Property TaxPercentage() As Double

        Get
            Return dbTaxPercentage
        End Get

        Set(ByVal Value As Double)
            dbTaxPercentage = Value
        End Set

    End Property

    Public Property TaxAccount() As Long

        Get
            Return lTaxAccount
        End Get

        Set(ByVal Value As Long)
            lTaxAccount = Value
        End Set

    End Property

    Public Property TaxAnnualIntervals() As Long

        Get
            Return lTaxAnnualIntervals
        End Get

        Set(ByVal Value As Long)
            lTaxAnnualIntervals = Value
        End Set

    End Property

    Public Property TaxStatus() As Boolean

        Get
            Return bTaxStatus
        End Get

        Set(ByVal Value As Boolean)
            bTaxStatus = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function NewDetails()

        Try

            lTaxID = 0
            strTaxName = ""
            dbTaxPercentage = 0
            lTaxAccount = 0
            lTaxAnnualIntervals = 0
            strTaxCategory = ""
            bTaxStatus = False
            strTaxCodeID = ""

        Catch ex As Exception

        End Try

    End Function

#End Region

#Region "DatabaseProcedures"

    'Save informaiton
    Public Function Save(ByVal DisplayErrorMessages As Boolean, _
            ByVal DisplaySuccessMessages As Boolean, _
                ByVal DisplayFailureMessages As Boolean) As Boolean

        'Saves a new base organization
        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String
            Dim objCOA As IMChartOfAccount = New IMChartOfAccount

            If Trim(strOrganizationName) = "" Then

                MsgBox("Please open an existing company.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Select an existing company")
                objLogin = Nothing
                datSaved = Nothing
                objCOA = Nothing

                Exit Function
            End If


            'Check if Tax text, Percentage, and Tax Account is provided
            If Trim(strTaxName) = "" Or _
                lTaxAccount = 0 Then

                MsgBox("You must provide a valid Tax Name " & _
                    Chr(10) & ", and the Chart of Account the Tax is related to." _
                        , MsgBoxStyle.Critical, _
                            "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing
                objCOA = Nothing

                Exit Function
            End If


            'Check if COA exists
            If Find("SELECT * FROM COA WHERE COAAccountNr = " _
                & Trim(lTaxAccount), _
                    False) = False Then

                MsgBox("The Chart Of Account provided does not exist.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - invalid or incomplete informaiton")

                objLogin = Nothing
                datSaved = Nothing
                objCOA = Nothing

                Exit Function
            End If


            'Check if there is an existing Tax with this name
            If Find("SELECT * FROM TaxCodes WHERE TaxTypeID = '" _
                & Trim(strTaxName) & "'", _
                    False) = True Then

                If MsgBox("The Tax Type you have provided exists." & _
                        Chr(10) & " Do you want to update this Tax Code's details?", _
                            MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, _
                                "iManagement - Set this format as the default?") _
                                    = MsgBoxResult.Yes Then


                    'Check if there is an existing format with this name
                    If Find("SELECT * FROM TaxCodes WHERE TaxTypeID = '" _
                        & Trim(strTaxName) & "' AND TaxStatus = FALSE", _
                            False) = True Then

                        If bTaxStatus = False Then
                            MsgBox("The current Tax Code is disabled. Enable it in order to update it.", _
                                MsgBoxStyle.Critical, _
                                    "iManagement - invalid or incomplete information")
                            objLogin = Nothing
                            datSaved = Nothing
                            objCOA = Nothing

                            Exit Function

                        End If
                    End If

                    Dim strItem As String
                    Dim arItem() As String
                    Dim lValCOAAccountNr As Long

                    arItem = objLogin.FillArray(strOrgAccessConnString, _
                        "SELECT TaxAccount FROM TaxCodes " & _
                    " WHERE TaxTypeID = '" & TaxTypeID & "'", "", "")

                    If Not arItem Is Nothing Then
                        For Each strItem In arItem
                            If Not strItem Is Nothing Then
                                lValCOAAccountNr = CLng(Val(strItem))

                            End If
                        Next
                    End If


                    'Check if the TaxCodeID exists
                    If Find("SELECT * FROM TaxCodes WHERE TaxCodeID = '" _
                        & Trim(strTaxCodeID) & _
                            "' AND TaxTypeId = '" & strTaxName & "'", _
                                False) = True Then

                        If bTaxStatus = False Then
                            MsgBox("The Tax Code ID already exists. The Tax Code ID must be unique", _
                                MsgBoxStyle.Critical, _
                                    "iManagement - Record Exists. Cannot Update")
                            objLogin = Nothing
                            datSaved = Nothing
                            objCOA = Nothing

                            Exit Function

                        End If
                    End If


                    bSaveSuccess = Update("UPDATE TaxCodes SET " & _
                    "TaxCategory = '" & Trim(strTaxCategory) & _
                        "', TaxPercentage = " & dbTaxPercentage & _
                            ", TaxAccount = " & lTaxAccount & _
                                ", TaxAnnualIntervals = " & lTaxAnnualIntervals & _
                                    ", TaxStatus = " & bTaxStatus & _
                                        " WHERE TaxTypeID = '" & _
                                                Trim(strTaxName) & "'")


                    If bSaveSuccess = True Then

                        objCOA.COAAccountNr = lValCOAAccountNr
                        objCOA.UnReserveAccount()

                        objCOA.COAAccountNr = lTaxAccount
                        objCOA.ReservedBy = Trim(strTaxName)
                        objCOA.ReserveAccount(False)

                    End If
                End If

                objLogin = Nothing
                datSaved = Nothing
                objCOA = Nothing

                Exit Function
            End If


            'Check if the TaxCodeID exists
            If Find("SELECT * FROM TaxCodes WHERE TaxCodeID = '" _
                & Trim(strTaxCodeID) & "'", False) = True Then

                If bTaxStatus = False Then
                    MsgBox("The Tax Code ID already exists. " & _
                        "The Tax Code ID must be unique.", _
                            MsgBoxStyle.Critical, _
                                "iManagement - Record Exists. Cannot Save")

                    objLogin = Nothing
                    datSaved = Nothing
                    objCOA = Nothing

                    Exit Function

                End If
            End If


            If MsgBox("Are you sure you want to save this new Tax Structure?", _
                MsgBoxStyle.YesNo, _
                    "Save the Tax Structure") = MsgBoxResult.No Then
                objLogin = Nothing
                datSaved = Nothing
                objCOA = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO TaxCodes (" & _
                "TaxTypeID," & _
                "TaxPercentage," & _
                "TaxAnnualIntervals," & _
                "TaxStatus," & _
                "TaxAccount," & _
                "TaxCategory," & _
                "TaxCodeID" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & Trim(strTaxName) & _
                    "'," & dbTaxPercentage & _
                    "," & lTaxAnnualIntervals & _
                    "," & bTaxStatus & _
                    "," & lTaxAccount & _
                    ",'" & Trim(strTaxCategory) & _
                    "','" & Trim(strTaxCodeID) & _
                    "')"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)


            objCOA.COAAccountNr = lTaxAccount
            objCOA.ReservedBy = Trim(strTaxName)
            objCOA.ReserveAccount(False)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement - Tax Code Saved")

                End If
                Return True
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save New Tax Code' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save New Tax Code Failed")
                End If
            End If

            objLogin = Nothing
            datSaved = Nothing
            objCOA = Nothing

        Catch ex As Exception
            If DisplayErrorMessages = True Then
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

                        lTaxID = _
                                myDataRows("TaxID")
                        strTaxName = _
                                myDataRows("TaxTypeID").ToString
                        lTaxAccount = _
                            myDataRows("TaxAccount")
                        lTaxAnnualIntervals = _
                            myDataRows("TaxAnnualIntervals")
                        dbTaxPercentage = _
                            myDataRows("TaxPercentage")
                        bTaxStatus = _
                            myDataRows("TaxStatus")
                        strTaxCategory = _
                            myDataRows("TaxCategory").ToString
                        strTaxCodeID = _
                            myDataRows("TaxCodeID").ToString

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

            If lTaxID = 0 Then

                MsgBox("Cannot Delete. Please select an existing Tax Code.", _
                        MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

                Exit Function

            End If

            strDeleteQuery = "DELETE * FROM TaxCodes WHERE TaxID = " & lTaxID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strDeleteQuery, _
                    datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", _
                    MsgBoxStyle.Information, _
                        "iManagement - Tax Code Deleted")
                Return True
            Else
                MsgBox("'Chart Of Account Format delete' action failed", _
                    MsgBoxStyle.Exclamation, " Tax Code Deletion failed")
            End If



            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function Update(ByVal strUpQuery As String) As Boolean

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
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Tax Code Details Updated")
                Return True
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Function

#End Region


End Class

