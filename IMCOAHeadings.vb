Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCOAHeadings


#Region "PrivateVariables"

    Private lHeaderID As Long
    Private strChartTitle As String
    Private lLowerRange As Long
    Private lUpperRange As Long
    Private strRelationshipTitle As String
    Private strLevelInRelation As String
    Private strChartOfAccountName As String
    Private strAccountingCategory As String
    Private dtDateCreated As Date
    Private MaxValue As Long
    Private strRootNode As Long

#End Region


#Region "Properties"

    Public Property AccountingCategory() As String

        Get
            Return strAccountingCategory
        End Get

        Set(ByVal Value As String)
            strAccountingCategory = Value
        End Set

    End Property

    Public Property DateCreated() As Date

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return dtDateCreated
        End Get

        Set(ByVal Value As Date)
            dtDateCreated = Value
        End Set

    End Property

    Public Property ChartOfAccountName() As String

        'USED TO SET AND RETRIEVE THE Chart Of Account Name 
        'Related to this Heading (STRING)
        Get
            Return strChartOfAccountName
        End Get

        Set(ByVal Value As String)
            If Trim(Value) = "" Then
                'Default is the standard value
                strChartOfAccountName = "Default"
            Else
                strChartOfAccountName = Value
            End If

        End Set

    End Property

    Public Property HeaderID() As Long

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lHeaderID
        End Get

        Set(ByVal Value As Long)
            lHeaderID = Value
        End Set

    End Property

    Public Property LowerRange() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lLowerRange
        End Get

        Set(ByVal Value As Long)
            lLowerRange = Value
        End Set

    End Property

    Public Property UpperRange() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lUpperRange
        End Get

        Set(ByVal Value As Long)
            lUpperRange = Value
        End Set

    End Property

    Public Property ChartTitle() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strChartTitle
        End Get

        Set(ByVal Value As String)
            strChartTitle = Value
        End Set

    End Property

    'The header under which the item is to be placed. If it is empty
    'then the node is placed as a root node
    Public Property RelationshipTitle() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strRelationshipTitle
        End Get

        Set(ByVal Value As String)
            strRelationshipTitle = Value
        End Set

    End Property

    Public Property LevelInRelation() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strLevelInRelation
        End Get

        Set(ByVal Value As String)
            strLevelInRelation = Value
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

        lHeaderID = 0
        strChartTitle = ""
        lLowerRange = 0
        lUpperRange = 0
        strRelationshipTitle = ""
        strLevelInRelation = ""


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

            If Trim(strChartTitle) = "" Or _
                Trim(strAccountingCategory) = "" Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide following details in" & _
                " order to save a Chart Range Header." & _
                Chr(10) & "1. Chart Title" & _
                Chr(10) & "2. Accounting Category" & _
                MsgBoxStyle.Critical, _
            "iManagement - Save Action Failed")

                End If

                Exit Function
            Else

                If Trim(strRelationshipTitle) = "" Then
                    If lUpperRange <> 0 Or lLowerRange <> 0 Then

                        MsgBox("A root node cannot have ranges. On" & _
                        Chr(10) & " the other hand, a child node must" & _
                        " have both sets of ranges")

                        Exit Function
                    End If

                End If


                'Check if the node to be added is a root node
                'And make sure it has no ranges
                'tis means it has no RelatinshipTitle or LevelInRelationship
                'should not have ranges
                If Trim(strRelationshipTitle) = "" Then
                    If lUpperRange <> 0 Or lLowerRange <> 0 Then

                        If DisplayErrorMessages = True Then
                            MsgBox("A root node cannot have ranges. On" & _
                            Chr(10) & " the other hand, a child node must" & _
                            " have both sets of ranges")


                        End If
                        Exit Function
                    End If
                End If

                If lUpperRange = 0 And lLowerRange <> 0 Then
                    If DisplayErrorMessages = True Then
                        MsgBox("You must provide both an upper and" & _
                        "a lower range." & _
                        Chr(10) & " The Upper Range is missing" _
                        , MsgBoxStyle.Exclamation, _
                        "iManagement - Invalid or incomplete information")


                    End If
                    Exit Function
                End If

                If lUpperRange <> 0 And lLowerRange = 0 Then
                    If DisplayErrorMessages = True Then
                        MsgBox("You must provide both an upper and" & _
                        "a lower range." & _
                        Chr(10) & " The Lower Range is missing" _
                        , MsgBoxStyle.Exclamation, _
                        "iManagement - Invalid or incomplete information")


                    End If
                    Exit Function
                End If

                'If the ranges are not empty, make sure that the 
                'parent does not have child nodes without ranges
                If lUpperRange = 0 Or lLowerRange = 0 And _
                    Trim(strRelationshipTitle) <> "" Then
                    If Find("SELECT * FROM ChartHeaders" & _
                    " WHERE ChartTitle = '" & Trim(strRelationshipTitle) & _
                        "' AND LowerRange <> 0 OR UpperRange <> 0", _
                            False) = True Then

                        If DisplayErrorMessages = True Then
                            MsgBox("The selected header already posseses" _
                            & Chr(10) & _
                            " child nodes with upper and lower ranges." _
                            & Chr(10) & _
                            " Only those nodes that do not have child" _
                            & Chr(10) & _
                            " nodes with ranges can hold those root nodes" _
                            & Chr(10) & _
                            " without ranges", _
                            MsgBoxStyle.Critical, _
                            "iManagement - Invalid or incomplete" & _
                            " information provided")
                        End If

                        Exit Function
                    End If
                End If

                '[Check if the Chart Title exists within this specific
                'chart e.g. If there is another chart title in
                'the 'Default' chart account
                If Find("SELECT * FROM ChartHeaders " & _
                            " WHERE " & _
                            " ChartTitle = '" & _
                            Trim(strChartTitle) & _
                            "' AND ChartOfAccountName " & _
                            " = '" & _
                            strChartOfAccountName _
                            & "'" _
                            , False) = True Then

                    If DisplayErrorMessages = True Then
                        MsgBox("This chart of account Title Exists." & _
                            " Type in another name for it" _
                            , MsgBoxStyle.Exclamation, _
                            "iManagement - Addition Failed")
                    End If

                    Return False
                    Exit Function
                End If

                'Check if the upper or lower ranges are used up (Part 1)
                If lLowerRange <> 0 Or lUpperRange <> 0 Then
                    If Find("SELECT * FROM ChartHeaders " & _
                                " WHERE " & _
                                " LowerRange <=" & _
                                lLowerRange & _
                                " AND UpperRange >=" & _
                                lLowerRange & _
                                " AND ChartOfAccountName " & _
                                " = '" & _
                                strChartOfAccountName _
                                & "'" _
                                , False) = True Then

                        If DisplayErrorMessages = False Then
                            MsgBox("This Lower Range Exists. Type in" & _
                                "  Lower range" _
                                , MsgBoxStyle.Exclamation, _
                                "iManagement - Addition Failed")
                        End If

                        Return False
                        Exit Function
                    End If

                    'Check if the upper or lower ranges are used up(Part 2)
                    If Find("SELECT * FROM ChartHeaders " & _
                                    " WHERE " & _
                                    " LowerRange <=" & _
                                    lUpperRange & _
                                    " AND UpperRange >=" & _
                                    lUpperRange & _
                                    " AND ChartOfAccountName " & _
                                    " = '" & _
                                    strChartOfAccountName _
                                    & "'" _
                                    , False) = True Then

                        If DisplayErrorMessages = False Then
                            MsgBox("This Upper Range Exists. Type" _
                            & Chr(10) & _
                            " in an Upper Range that is not already used." _
                            , MsgBoxStyle.Exclamation, _
                            "iManagement - Addition Failed")
                        End If

                        Return False
                        Exit Function
                    End If


                End If


                strInsertInto = "INSERT INTO ChartHeaders (" & _
                    "ChartTitle," & _
                    "LowerRange," & _
                    "UpperRange," & _
                    "RelationshipTitle," & _
                    "LevelInRelation," & _
                    "ChartOfAccountName," & _
                    "AccountingCategory" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "('" & strChartTitle & _
                        "'," & lLowerRange & _
                        "," & lUpperRange & _
                        ",'" & strRelationshipTitle & _
                        "','" & strLevelInRelation & _
                        "','" & strChartOfAccountName & _
                        "','" & strAccountingCategory & _
                                "')"

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bSaveSuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, _
                strSaveQuery, _
                datSaved)


                objLogin.CloseDb()



                If bSaveSuccess = True Then
                    If DisplaySuccess = True Then
                        MsgBox("Chart Header Saved Successfully", _
                        MsgBoxStyle.Information, _
                        "iManagement - Record Saved Successfully")

                    End If

                Else

                    If DisplayFailure = True Then
                        MsgBox("'Save Chart Header' action failed." & _
                " Make sure all mandatory details are entered", _
                MsgBoxStyle.Exclamation, _
                "iManagement -  Addition Failed")

                    End If

                    Exit Function

                End If
            End If

            objLogin = Nothing
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

            objLogin.connectString = strOrgAccessConnString
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
                        Return False
                        Exit Function

                    End If

                    'Whether to fill properties with values or not
                    If ReturnStatus = True Then

                        For Each myDataRows In myDataTables.Rows

                            lHeaderID = _
                                myDataRows("HeaderID").ToString()
                            strChartTitle = _
                                myDataRows("ChartTitle").ToString()
                            lLowerRange = _
                                myDataRows("LowerRange").ToString()
                            lUpperRange = _
                                myDataRows("UpperRange").ToString()
                            strRelationshipTitle = _
                                myDataRows("RelationshipTitle").ToString()
                            strLevelInRelation = _
                                myDataRows("LevelInRelation").ToString()
                            strAccountingCategory = _
                                myDataRows("AccountingCategory").ToString()
                            strChartOfAccountName = _
                                myDataRows("ChartOfAccountName").ToString()

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

    Public Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strDeleteQuery = strDelQuery

        If lHeaderID <> 0 _
                            Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

           

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Chart Of Account Heading Details Deleted", MsgBoxStyle.Information, _
                    "iManagement - Record Deleted Successfully")

            Else

                MsgBox("'Delete Chart Of Account Heading' action failed", _
                    MsgBoxStyle.Exclamation, "Deletion failed")

                objLogin.RollbackTheTrans()

            End If

        Else

            MsgBox("Cannot Delete. Please select an existing Sequence's Detail", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

            objLogin.RollbackTheTrans()

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

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Chart Of Account Format Updated")
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

            Return strTextFieldData
            datFillData.Dispose()

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnMaxValue(ByVal strQuery As String) As Boolean
        'Query must contain at least rows from Sequence

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

                    MaxValue = myDataRows(0)

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Function ReturnChartFormatDesignerTree _
        (ByVal strFormatName As String, ByVal myTrControl As Object, _
            ByVal DisplayErrorsMessages As Boolean, _
                ByVal bReturnCOATrees As Boolean)

        Try
            Dim arTreeNodes() As Object
            Dim trnItem As Object

            myTrControl.Nodes.Clear()

            If myTrControl Is Nothing Then
                Exit Function
            End If

            'Dim arOpeningBalance() As String = Split _
            '            ("", ",")
            myTrControl.Nodes.Add("Opening Balance")
            '-------------------------------------------

            'Dim arAssets() As String = Split _
            '            ("Short Term Assets or Current Assets,Long Term Assets", ",")

            myTrControl.Nodes.Add("Assets")
            myTrControl.Nodes(1).Nodes.Add("Short Term Assets or Current Assets")
            myTrControl.Nodes(1).Nodes.Add("Long Term Assets")

            ReDim arTreeNodes(2)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Short Term Assets or Current Assets", True, _
                    myTrControl.Nodes(1).Nodes(0), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Long Term Assets", True, _
                    myTrControl.Nodes(1).Nodes(1), bReturnCOATrees)
            '********************************************
            '--------------------------------------------

            'Dim arLiabilities() As String = Split _
            '            ("Short Term Liabilities or Current Liabilities" & _
            '             ",Long Term Liabilities" & _
            '             ",Shareholder Assets", _
            '                    ",")

            myTrControl.Nodes.Add("Liabilities")
            myTrControl.Nodes(2).Nodes.Add("Short Term Liabilities or Current Liabilities")
            myTrControl.Nodes(2).Nodes.Add("Long Term Liabilities")
            myTrControl.Nodes(2).Nodes.Add("Shareholder Assets")

            ReDim arTreeNodes(3)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Short Term Liabilities or Current Liabilities", True, _
                    myTrControl.Nodes(2).Nodes(0), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Long Term Liabilities", True, _
                    myTrControl.Nodes(2).Nodes(1), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Shareholder Assets", True, _
                    myTrControl.Nodes(2).Nodes(2), bReturnCOATrees)
            '********************************************
            '--------------------------------------------

            'Dim arProfitAndLoss() As String = Split _
            '            ("Income,Costs or Expenditures", ",")

            myTrControl.Nodes.Add("Profit and Loss")

            'Profit and loss main sub folders
            myTrControl.Nodes(3).Nodes.Add("Income")
            myTrControl.Nodes(3).Nodes.Add("Costs or Expenditures")

            'Dim arProfitAndLossIncome() As String = Split _
            '         ("Trading Income or Operating Income,Grants", ",")

            'Income sub folders in Profit and loss
            myTrControl.Nodes(3).Nodes(0).Nodes.Add("Trading Income or Operating Income")
            myTrControl.Nodes(3).Nodes(0).Nodes.Add("Grants")

            ReDim arTreeNodes(2)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Trading Income or Operating Income", True, _
                    myTrControl.Nodes(3).Nodes(0).Nodes(0), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                    "Grants", True, _
                        myTrControl.Nodes(3).Nodes(0).Nodes(1), bReturnCOATrees)


            ''''''''''''''''******************************************** 

            'Dim arProfitAndLossCostsorExpenditures() As String = Split _
            '          ("Labour Costs,Raw Materials,Other Overhead Costs," & _
            '              "Tax,Financial Costs,Financial Income," & _
            '                  "Unexpected Income,Unexpected Costs", ",")


            'Costs or Expenditures sub folders in Profit and loss
            myTrControl.Nodes(3).Nodes(1).Nodes.Add("Labour Costs")
            myTrControl.Nodes(3).Nodes(1).Nodes.Add("Raw Materials")
            myTrControl.Nodes(3).Nodes(1).Nodes.Add("Other Overhead Costs")
            myTrControl.Nodes(3).Nodes(1).Nodes.Add("Tax")
            myTrControl.Nodes(3).Nodes(1).Nodes.Add("Financial Costs")
            myTrControl.Nodes(3).Nodes(1).Nodes.Add("Financial Income")
            myTrControl.Nodes(3).Nodes(1).Nodes.Add("Unexpected Income")
            myTrControl.Nodes(3).Nodes(1).Nodes.Add("Unexpected Costs")

            ReDim arTreeNodes(8)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Labour Costs", True, _
                    myTrControl.Nodes(3).Nodes(1).Nodes(0), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Raw Materials", True, _
                    myTrControl.Nodes(3).Nodes(1).Nodes(1), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Other Overhead Costs", True, _
                    myTrControl.Nodes(3).Nodes(1).Nodes(2), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Tax", True, _
                     myTrControl.Nodes(3).Nodes(1).Nodes(3), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Financial Costs", True, _
                    myTrControl.Nodes(3).Nodes(1).Nodes(4), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Financial Income", True, _
                    myTrControl.Nodes(3).Nodes(1).Nodes(5), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Unexpected Income", True, _
                    myTrControl.Nodes(3).Nodes(1).Nodes(6), bReturnCOATrees)

            ReturnChartFormatDesignerTreeWithLevels(strFormatName, _
                "Unexpected Costs", True, _
                    myTrControl.Nodes(3).Nodes(1).Nodes(7), bReturnCOATrees)
            '--------------------------------------------


        Catch ex As Exception

            If DisplayErrorsMessages = True Then
                MsgBox(ex.Message.ToString, _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Critical System Error")
            End If

        End Try

    End Function

    Public Function ReturnCOAHeaderIDFromRange _
           (ByVal strRange As String) As Long

        Try

            Dim arReturn() As String
            Dim strItemRet As String

            Dim objL As IMLogin = New IMLogin

            With objL

                arReturn = .FillArray(strOrgAccessConnString, _
                " SELECT HeaderID FROM ChartHeaders WHERE " & _
                " LowerRange = " & CLng(Val(strRange)) _
                , "", "")

                If Not arReturn Is Nothing Then

                    For Each strItemRet In arReturn

                        If Not strItemRet Is Nothing Then
                            Return CLng(Val(arReturn(0)))

                        End If

                    Next
                End If

            End With

            objL = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnChartFormatDesignerTreeWithLevels _
           (ByVal strFormatName As String, ByVal strRelationshipTitle As String, _
                ByVal DisplayErrorsMessages As Boolean, _
                    ByVal TrNodeRef As Object, _
                        ByVal bReturnCOA As Boolean)

        Try

            Dim datRetDataWithRanges As DataSet = New DataSet
            Dim bQuerySuccessWithRanges As Boolean
            Dim myDataTablesWithRanges As DataTable
            Dim myDataColumnsWithRanges As DataColumn
            Dim myDataRowsWithRanges As DataRow
            Dim objLogin As IMLogin = New IMLogin
            Dim strTreeBaseQueryWithRanges As String
            Dim strCurrentRelationshipWithRanges As String
            Dim i As Long

            Dim TrNodeWithRanges As Object 'For the Ranges
            Dim TrNodeWithCOA As Object 'For COA 

            If TrNodeRef Is Nothing Then
                datRetDataWithRanges = Nothing
                objLogin = Nothing
                Exit Function
            End If

            'Check if there is a tree control passed.
            'If not exit
            If strRelationshipTitle = "" Then
                datRetDataWithRanges = Nothing
                objLogin = Nothing
                Exit Function
            End If


            'Select all child nodes for this parent node
            '------------------Get all values that have ranges (WithRanges)
            strTreeBaseQueryWithRanges = _
                "SELECT ChartHeaders.HeaderID, " & _
    "(ChartHeaders.ChartTitle+' - Ranges'+str(LowerRange)" & _
    "+' to'+str(UpperRange)) AS RangesText " & _
    " FROM ChartHeaders " & _
    " WHERE ChartHeaders.LowerRange <> 0 And " & _
    " ChartHeaders.UpperRange <> 0 And " & _
    " ChartHeaders.RelationshipTitle = '" & _
    Trim(strRelationshipTitle) & _
    "' " & _
    " AND ChartHeaders.ChartOfAccountName = '" & Trim(strFormatName) & "'" & _
    " ORDER BY ChartHeaders.RelationshipTitle, ChartHeaders.LowerRange"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bQuerySuccessWithRanges = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strTreeBaseQueryWithRanges, _
                                        datRetDataWithRanges)

            objLogin.CloseDb()
            '---------------------------------------------

            '*********************************
            '***************Start of nodes with children but no ranges
            If bQuerySuccessWithRanges = True Then

                Dim iWithRanges As Integer
                TrNodeWithRanges = New Object

                For Each myDataTablesWithRanges In _
                                datRetDataWithRanges.Tables

                    'Check if there is any data. If not exit
                    If myDataTablesWithRanges.Rows.Count = 0 Then

                        'Exit the loop for WithNoRangesButChilds indicating 
                        ' that the search for the values with no ranges
                        ' but have children was not successful
                        Exit For

                    End If

                    'Set Node identifier to 0 for each record
                    i = 0

                    For Each myDataRowsWithRanges In _
                            myDataTablesWithRanges.Rows

                        TrNodeRef.Nodes.Add _
                            (myDataRowsWithRanges("RangesText").ToString())

                        If bReturnCOA = True Then
                            Dim arItem() As String
                            Dim strRange As String
                            Dim strItemRange As String


                            arItem = objLogin.FillArray _
                                (strOrgAccessConnString, _
                                    "SELECT (str(COAAccountNr) + ' - ' + " & _
                                        " trim(AccountTitle)) FROM COA " & _
                                            " WHERE HeaderID = " & _
                                                myDataRowsWithRanges _
                                                ("HeaderID") _
                                                        , "", "")

                            If Not arItem Is Nothing Then

                                For Each strItemRange In arItem
                                    If Not strItemRange Is Nothing Then
                                        TrNodeRef.Nodes(i).Nodes.Add _
                                           (strItemRange)

                                    End If

                                Next

                            End If

                        End If

                        'Select all values from COA that are between upper and lower ranges

                        'Add each value as a sub node to the above node

                        i = i + 1
                    Next
                Next
            End If
            '-----------------------End of nodes children with ranges

            datRetDataWithRanges = Nothing
            objLogin = Nothing

            If Not (TrNodeWithRanges Is Nothing) Then
                Return TrNodeWithRanges
            End If


        Catch ex As Exception

            If DisplayErrorsMessages = True Then
                MsgBox(ex.Message.ToString, _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Critical System Error")

            End If

        End Try

    End Function

#End Region


End Class
