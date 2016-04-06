Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

'DEFAULT COA CONSTANTS
'Exchange Rate Profits Default
'Exchange Rate Losses Default
'VAT COA Default
'Bank Charges Default
'Trade Creditors Default
'Trade Debtors Default
'Cash Payments Default
'Credit Card Payments Default
'Stock Account Default
'Stock Write-Off Account Default

Public Class IMChartOfAccount

#Region "PrivateVariables"

    Private lCOAAccountNr As Long
    Private strAccountTitle As String
    Private bCreditWarning As Boolean
    Private bDebitWarning As Boolean
    Private strTextDescription As String
    Private lVATCodeID As Long
    Private lProductID As Long
    Private lShortcutNo As Long
    Private lHeaderID As Long
    Private MaxValue As Long
    Private dtDateCreated As Date
    Private bIsReserved As Boolean
    Private strReservedBy As String

#End Region

#Region "Properties"

    Public Property ReservedBy() As String

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return strReservedBy
        End Get

        Set(ByVal Value As String)
            strReservedBy = Value
        End Set

    End Property


    Public Property IsReserved() As Boolean

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return bIsReserved
        End Get

        Set(ByVal Value As Boolean)
            bIsReserved = Value
        End Set

    End Property

    Public Property AccountTitle() As String

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return strAccountTitle
        End Get

        Set(ByVal Value As String)
            strAccountTitle = Value
        End Set

    End Property

    Public Property COAAccountNr() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lCOAAccountNr
        End Get

        Set(ByVal Value As Long)
            lCOAAccountNr = Value
        End Set

    End Property

    Public Property CreditWarning() As Boolean

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return bCreditWarning
        End Get

        Set(ByVal Value As Boolean)
            bCreditWarning = Value
        End Set

    End Property

    Public Property DebitWarning() As Boolean

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return bDebitWarning
        End Get

        Set(ByVal Value As Boolean)
            bDebitWarning = Value
        End Set

    End Property

    Public Property TextDescription() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strTextDescription
        End Get

        Set(ByVal Value As String)
            strTextDescription = Value
        End Set

    End Property

    Public Property VATCodeID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lVATCodeID
        End Get

        Set(ByVal Value As Long)
            lVATCodeID = Value
        End Set

    End Property

    Public Property ProductID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lProductID
        End Get

        Set(ByVal Value As Long)
            lProductID = Value
        End Set

    End Property

    Public Property ShortcutNo() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lShortcutNo
        End Get

        Set(ByVal Value As Long)
            lShortcutNo = Value
        End Set

    End Property

    Public Property HeaderID() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lHeaderID
        End Get

        Set(ByVal Value As Long)
            lHeaderID = Value
        End Set

    End Property

    Public Property DateCreated() As Date

        Get
            Return dtDateCreated
        End Get

        Set(ByVal Value As Date)
            dtDateCreated = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Sub NewChartRecord()

        lCOAAccountNr = 0
        strAccountTitle = ""
        bCreditWarning = False
        bDebitWarning = False
        strTextDescription = ""
        lVATCodeID = 0
        lProductID = 0
        lShortcutNo = 0
        lHeaderID = 0
        DateCreated = Now()

    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Function ReserveAccount(ByVal _
        bDisplayConfirmation As Boolean) As Boolean

        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String


            If lCOAAccountNr = 0 Or _
                Trim(strReservedBy) = "" Then
                MsgBox("You cannot update the Reservation details. " & _
                Chr(10) & "Provide the reservation Purpose as well as the " & _
                Chr(10) & "Chart Of Account to reserve with", _
                MsgBoxStyle.Exclamation, _
                    "iManagement - invalid or incomplete information")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM COA WHERE " & _
               " COAAccountNr = " & lCOAAccountNr, False) = False Then

                MsgBox("You cannot use this account number" & _
                " since it is not inserted in the chart of accounts.", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM COA WHERE " & _
            " COAAccountNr = " & lCOAAccountNr & _
                " AND IsReserved = TRUE", False) = True Then

                MsgBox("The Chart Of Account Number is already reserved." _
                , MsgBoxStyle.Exclamation, _
                        "iManagement - Record Exists")

                datSaved = Nothing
                objLogin = Nothing
                Exit Function
            End If


            If Find("SELECT * FROM COA WHERE " & _
            " ReservedBy = '" & strReservedBy & "'", False) = True Then

                MsgBox("Either the default account already has an " & _
                "entry or the Reserd By text already exists. " & _
                "You cannot have duplicate 'Reserved By' texts.", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Record Exists")

                datSaved = Nothing
                objLogin = Nothing
                Exit Function
            End If


            If bDisplayConfirmation = True Then
                If MsgBox("Do you want to set this default Chart Of Account " & _
                "Number and Reserve it?", _
                MsgBoxStyle.YesNo, _
                "iManagement - Add New Record?") = MsgBoxResult.No Then

                    datSaved = Nothing
                    objLogin = Nothing
                    Exit Function
                End If
            End If


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            strInsertInto = "UPDATE COA SET " & _
                    " ReservedBy = '" & Trim(strReservedBy) & "' , " & _
                        "IsReserved = TRUE " & " WHERE  COAAccountNr = " & _
                            lCOAAccountNr

            bSaveSuccess = objLogin.ExecuteQuery _
                       (strOrgAccessConnString, _
                               strInsertInto, _
                                       datSaved)

            objLogin.CloseDb()

            objLogin = Nothing
            datSaved = Nothing

            Return bSaveSuccess

        Catch ex As Exception

        End Try

    End Function

    Public Function UnReserveAccount() As Boolean

        Try

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String

            If lCOAAccountNr = 0 Then
                MsgBox("You cannot update the Reservation details. " & _
                Chr(10) & "Please provide the Chart Of Account to Unreserve.", _
                MsgBoxStyle.Exclamation, _
                    "iManagement - invalid or incomplete information")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()


            strInsertInto = "UPDATE COA SET " & _
                  " ReservedBy = '', IsReserved = FALSE " & _
                      " WHERE  COAAccountNr = " & lCOAAccountNr

            bSaveSuccess = objLogin.ExecuteQuery _
                       (strOrgAccessConnString, _
                               strInsertInto, _
                                       datSaved)

            objLogin.CloseDb()

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function CheckIfRangeIsValidForHeader( _
         ByVal lValHeaderID As Long, _
             ByVal lValCOAAccountNr As Long, _
                 ByVal bZeroLowerAccepted As Boolean, _
                     ByVal bZeroUpperAccepted As Boolean) As Boolean

        Try

            Dim arReturn() As String
            Dim strItemRet As String
            Dim lValLower As Long
            Dim lValUpper As Long

            Dim objL As IMLogin = New IMLogin

            With objL

                arReturn = .FillArray(strOrgAccessConnString, _
                " SELECT UpperRange FROM ChartHeaders WHERE " & _
                " HeaderID= " & lValHeaderID _
                , "", "")


                If Not arReturn Is Nothing Then
                    For Each strItemRet In arReturn
                        If Not strItemRet Is Nothing Then
                            lValUpper = CLng(Val(arReturn(0)))

                        End If
                    Next
                End If


                arReturn = .FillArray(strOrgAccessConnString, _
                " SELECT LowerRange FROM ChartHeaders WHERE " & _
                " HeaderID= " & lValHeaderID _
                , "", "")

                If Not arReturn Is Nothing Then
                    For Each strItemRet In arReturn
                        If Not strItemRet Is Nothing Then
                            lValLower = CLng(Val(arReturn(0)))

                        End If
                    Next
                End If

            End With

            objL = Nothing


            'Make sure upper and lower are present
            If bZeroLowerAccepted = False Then
                If lValLower = 0 Then
                    Return False
                    Exit Function

                End If
            End If


            'Make sure upper and lower are present
            If bZeroUpperAccepted = False Then
                If lValUpper = 0 Then
                    Return False
                    Exit Function

                End If
            End If


            If lValCOAAccountNr > lValUpper Then
                Return False
                Exit Function

            End If


            If lValCOAAccountNr < lValLower Then
                Return False
                Exit Function

            End If

            Return True

        Catch ex As Exception

        End Try

    End Function

    Public Function CheckIfValueIsBetweenAcceptedRanges( _
            ByVal lValHeaderID As Long, _
                ByVal lValCOAAccountNr As Long, _
                    ByVal bZeroLowerAccepted As Boolean, _
                        ByVal bZeroUpperAccepted As Boolean) As Boolean

        Try

            Dim arReturn() As String
            Dim strItemRet As String
            Dim lValLower As Long
            Dim lValUpper As Long

            Dim objL As IMLogin = New IMLogin

            With objL

                arReturn = .FillArray(strOrgAccessConnString, _
                " SELECT LowerRange FROM ChartHeaders WHERE " & _
                " HeaderID= " & lValHeaderID _
                , "", "")

                If Not arReturn Is Nothing Then

                    For Each strItemRet In arReturn

                        If Not strItemRet Is Nothing Then
                            lValLower = CLng(Val(arReturn(0)))

                        End If

                    Next
                End If


                arReturn = .FillArray(strOrgAccessConnString, _
                " SELECT LowerRange FROM ChartHeaders WHERE " & _
                " HeaderID= " & lValHeaderID _
                , "", "")

                If Not arReturn Is Nothing Then

                    For Each strItemRet In arReturn

                        If Not strItemRet Is Nothing Then
                            lValLower = CLng(Val(arReturn(0)))

                        End If

                    Next
                End If

            End With

            objL = Nothing


            'Make sure upper and lower are present
            If bZeroLowerAccepted = False Then
                If lValLower = 0 Then
                    Return False
                    Exit Function

                End If
            End If


            'Make sure upper and lower are present
            If bZeroUpperAccepted = False Then
                If lValUpper = 0 Then
                    Return False
                    Exit Function


                End If
            End If

            If lValCOAAccountNr > lValUpper Then
                Return False
                Exit Function

            End If

            If lValCOAAccountNr < lValLower Then
                Return False
                Exit Function

            End If

            Return True


        Catch ex As Exception

        End Try

    End Function

    Public Function Save(ByVal DisplayMessages As Boolean) As Boolean
        'Saves a new sequence group details

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String


        Try

            If lCOAAccountNr = 0 Or _
                       Trim(strAccountTitle) = "" Or _
                               lHeaderID = 0 _
                                           Then

                MsgBox("Please provide the following details in order to " & _
                     "save an Account in the Chart Of Account." & _
                    Chr(10) & "1.Chart Account Number." & _
                        Chr(10) & "2.Account's Name (Short Description)." & _
                        Chr(10) & "3.Select the Chart Account" & _
                        " Header (Range) under which the new Account will be added", _
                                    MsgBoxStyle.Critical, "Save Action Failed")

                objLogin = Nothing
                datSaved = Nothing
                Exit Function

            End If


            If InStr(strAccountTitle, " - Range") = True Then

                MsgBox("The Chart Of Account Name cannot use the combination " & _
                "'- Range.'", _
                MsgBoxStyle.Exclamation, _
                    "iManagement - invalid or incomplete iformation")

                strAccountTitle = ""

                objLogin = Nothing
                datSaved = Nothing
                Exit Function

            End If

            If IsNumeric(Trim(strAccountTitle)) = True Then

                MsgBox("The Chart Of Account Name cannot start with or be a " & _
                "numeric value.", _
                MsgBoxStyle.Exclamation, _
                    "iManagement - invalid or incomplete iformation")

                strAccountTitle = ""

                objLogin = Nothing
                datSaved = Nothing
                Exit Function

            End If

            If IsReserved = True Then
                If Trim(strReservedBy) = "" Then
                    MsgBox("You must provide a valid reserved by text.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - invalid or incomplete information")

                    objLogin = Nothing
                    datSaved = Nothing
                    Exit Function
                End If
            End If

            '[Check if the HeaderID exists
            If Find("SELECT HeaderID FROM ChartHeaders " & _
                        " WHERE " & _
                            " HeaderID = " & _
                                lHeaderID _
                                    , False) = False Then
                MsgBox("The Chart Heading provided does not exist", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

                objLogin = Nothing
                datSaved = Nothing
                Exit Function
            End If

            '[Check if there is another COA with this title]
            If Find("SELECT COAAccountNr FROM COA " & _
                        " WHERE " & _
                            " AccountTitle = '" & _
                                strAccountTitle & _
                                    "' AND COAAccountNr <> " & _
                                    lCOAAccountNr, False) = True Then

                MsgBox("The Account Name provided for the Chart Of Account already exists." & _
                            Chr(10) & "Provide another name for your account.", _
                                MsgBoxStyle.Exclamation, _
                                    "iManagement - invalid or incomplete information")

                objLogin = Nothing
                datSaved = Nothing
                Exit Function
            End If


            '[Check for range of values]
            If CheckIfRangeIsValidForHeader( _
                lHeaderID, lCOAAccountNr, _
                    False, False) = False Then

                MsgBox("The selected Range cannot take this value. Type " & _
                        "in an acceptable value" & _
                            Chr(10) & _
                                " found within the Upper and Lower " & _
                                    "Ranges of the selected header." _
                                        , MsgBoxStyle.Exclamation, _
                                            "iManagement - invalid or incomplete information")

                objLogin = Nothing
                datSaved = Nothing
                Exit Function
            End If


            '[Check if the there is another COA with this number]
            If Find("SELECT COAAccountNr FROM COA " & _
                        " WHERE " & _
                            " COAAccountNr = " & _
                                lCOAAccountNr _
                                    , False) = True Then

                'Confirm Update. If yes, update
                If MsgBox("The Chart Of Account number provided exists." & _
                            Chr(10) & "Do you want to update its details?", _
                                MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, _
                    "iManagement - invalid or incomplete information") _
                        = MsgBoxResult.Yes Then

                    Update("UPDATE COA SET " & _
                       " AccountTitle = '" & Trim(strAccountTitle) & _
                       "' , CreditWarning = " & bCreditWarning & _
                       " , DebitWarning = " & bDebitWarning & _
                       " , TextDescription = '" & Trim(strTextDescription) & _
                       "' , VATCodeID = " & lVATCodeID & _
                       " , ProductID = " & lProductID & _
                       " , ShortcutNo = " & lShortcutNo & _
                       " , HeaderID = " & lHeaderID & _
                       " , IsReserved = " & bIsReserved & _
                       " , ReservedBy = '" & Trim(strReservedBy) & _
                           "' WHERE  COAAccountNr = " & lCOAAccountNr)

                End If

                objLogin = Nothing
                datSaved = Nothing
                Exit Function
            End If




            If MsgBox("Are you sure you Add this Chart Of" & _
     " Account Number '" & lCOAAccountNr & " - " & strAccountTitle & "'?", _
         MsgBoxStyle.YesNo + MsgBoxStyle.Information, _
             "iManagement - Delete Record?") = MsgBoxResult.No Then

                datSaved = Nothing
                objLogin = Nothing
                Exit Function
            End If


            'If the HeaderID does not exist in ChartHeaders then
            strInsertInto = "INSERT INTO COA (" & _
                "COAAccountNr," & _
                "AccountTitle," & _
                "CreditWarning," & _
                "DebitWarning," & _
                "TextDescription," & _
                "VATCodeID," & _
                "ProductID," & _
                "ShortcutNo," & _
                "HeaderID," & _
                "IsReserved," & _
                "ReservedBy" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lCOAAccountNr & _
                    ",'" & Trim(strAccountTitle) & _
                    "'," & bCreditWarning & _
                    "," & bDebitWarning & _
                    ",'" & Trim(strTextDescription) & _
                    "'," & lVATCodeID & _
                    "," & lProductID & _
                    "," & lShortcutNo & _
                    "," & lHeaderID & _
                    "," & bIsReserved & _
                     ",'" & Trim(strReservedBy) & _
                            "')"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True And DisplayMessages = True Then
                MsgBox("New Chart Of Account Number Saved Successfully", _
                    MsgBoxStyle.Information, _
                        "iManagement - Record Saved Successfully")

                Return True

            ElseIf bSaveSuccess = False And DisplayMessages = True Then

                MsgBox("'Save Chart Account action failed." & _
                    " Make sure all mandatory details are entered.", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Chart Of Account Addition Failed")

                Exit Function

            End If

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception
            MsgBox(ex.Message.ToString, _
                MsgBoxStyle.Exclamation, _
                    "iManagement System Error")

        End Try

    End Function

    Public Function Find(ByVal strQuery As String, _
                         ByVal ReturnStatus As Boolean) As Boolean

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

                    If ReturnStatus = True Then

                        For Each myDataRows In myDataTables.Rows

                            lCOAAccountNr = myDataRows("COAAccountNr")
                            strAccountTitle = Trim(myDataRows("AccountTitle").ToString())
                            bCreditWarning = myDataRows("CreditWarning")
                            bDebitWarning = myDataRows("DebitWarning")
                            strTextDescription = Trim(myDataRows("TextDescription").ToString())
                            lVATCodeID = myDataRows("VATCodeID")
                            lProductID = myDataRows("ProductID")
                            lShortcutNo = myDataRows("ShortcutNo")
                            lHeaderID = myDataRows("HeaderID")
                            dtDateCreated = myDataRows("DateCreated")
                            bIsReserved = myDataRows("IsReserved")
                            strReservedBy = Trim(myDataRows("ReservedBy").ToString())


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
                    "iManagement System Error")

        End Try

    End Function

    Public Sub Delete()

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        If lCOAAccountNr = 0 Then
            MsgBox("Cannot Delete. Please select an existing Chart Of Acocunt number", _
                MsgBoxStyle.Exclamation, _
                    "iManagement -Missing Information")

            datDelete = Nothing
            objLogin = Nothing
            Exit Sub

        End If


        If MsgBox("Are you sure you want to delete this Chart Of" & _
            " Account Number '" & lCOAAccountNr & " - " & _
                strAccountTitle & "'?", _
                MsgBoxStyle.YesNo + MsgBoxStyle.Information, _
                    "iManagement - Delete Record?") = MsgBoxResult.No Then

            datDelete = Nothing
            objLogin = Nothing
            Exit Sub
        End If

        strDeleteQuery = "DELETE * FROM COA WHERE " & _
                " COAAccountNr = " & lCOAAccountNr

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, _
                datDelete)

        objLogin.CloseDb()

        If bDelSuccess = True Then
            MsgBox("Sequence Details Deleted", _
                MsgBoxStyle.Information, _
                    "iManagement - Record Deleted Successfully")
        Else
            MsgBox("'Delete Chart Of Account' action failed", _
                MsgBoxStyle.Exclamation, "Deletion failed")
        End If


    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If lHeaderID <> 0 _
                        Then

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery _
                        (strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", _
                    MsgBoxStyle.Information, _
                        "iManagement -  Chart Of Account Details Updated")
            End If

        End If

    End Sub

#End Region

End Class
