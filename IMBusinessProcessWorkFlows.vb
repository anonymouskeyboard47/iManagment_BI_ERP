Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMBusinessProcessWorkFlows

#Region "PrivateVariables"

    Private lBusinessProcessID As Long
    Private lWorkFlowID As Long
    Private strWorkFlowName As String
    Private strWorkFlowDescription As String
    Private bWorkFlowStatus As Boolean
    Private lWorkFlowPosition As Long
    Private lMoveUpPosition As Long
    Private lMoveDownPosition As Long
    Private strWorkFlowType As String
    Private lHoursForCompletion As Long

#End Region

#Region "Properties"

    Public Property HoursForCompletion() As Long

        Get
            Return lHoursForCompletion
        End Get

        Set(ByVal Value As Long)
            lHoursForCompletion = Value
        End Set

    End Property

    Public Property WorkFlowType() As String

        Get
            Return strWorkFlowType
        End Get

        Set(ByVal Value As String)
            strWorkFlowType = Value
        End Set

    End Property

    Public Property MoveUpPosition() As Long

        Get
            Return lMoveUpPosition
        End Get

        Set(ByVal Value As Long)
            lMoveUpPosition = Value
        End Set

    End Property

    Public Property MoveDownPosition() As Long

        Get
            Return lMoveDownPosition
        End Get

        Set(ByVal Value As Long)
            lMoveDownPosition = Value
        End Set

    End Property

    Public Property BusinessProcessID() As Long

        Get
            Return lBusinessProcessID
        End Get

        Set(ByVal Value As Long)
            lBusinessProcessID = Value
        End Set

    End Property

    Public Property WorkFlowID() As Long

        Get
            Return lWorkFlowID
        End Get

        Set(ByVal Value As Long)
            lWorkFlowID = Value
        End Set

    End Property

    Public Property WorkFlowName() As String

        Get
            Return strWorkFlowName
        End Get

        Set(ByVal Value As String)
            strWorkFlowName = Value
        End Set

    End Property

    Public Property WorkFlowDescription() As String

        Get
            Return strWorkFlowDescription
        End Get

        Set(ByVal Value As String)
            strWorkFlowDescription = Value
        End Set

    End Property

    Public Property WorkFlowStatus() As Boolean

        Get
            Return bWorkFlowStatus
        End Get

        Set(ByVal Value As Boolean)
            bWorkFlowStatus = Value
        End Set

    End Property

    Public Property WorkFlowPosition() As Long

        Get
            Return lWorkFlowPosition
        End Get

        Set(ByVal Value As Long)
            lWorkFlowPosition = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function CalculateNextWorkFlowPositionNo(ByVal BusinessProcessID As Long) As String

        Try

            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String
            Dim strProposedApplNo As String

            Dim objLogin As IMLogin = New IMLogin

            With objLogin

                MyMaxValue = .FillArray(strOrgAccessConnString, _
                           "SELECT Max(WorkFlowPosition) AS TotalRecords FROM" & _
                               " ApplicationMaster WHERE ApplicationDate = Now()", "", "")
            End With

            objLogin = Nothing

            If Not MyMaxValue Is Nothing Then
                For Each strItem In MyMaxValue
                    If Not strItem Is Nothing Then

                        MaxValue = CLng(Val(strItem))


                    End If
                Next
            End If

            MaxValue = MaxValue + 1

            strProposedApplNo = "Appl" & Now.Day.ToString _
                & Now.Month.ToString & _
                    Now.Year.ToString & _
                            MaxValue.ToString

            Return strProposedApplNo

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, _
                "iManagement - System Error")
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


            If lMoveDownPosition <> 0 And lMoveUpPosition <> 0 Then

                MsgBox("You cannot move up and down at the same time." & _
                " Please select one direction at a time." _
                , MsgBoxStyle.Exclamation, _
                "iManagement - invalid or incomplete information provided")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            If Trim(strWorkFlowName) = "" Or _
                Trim(strWorkFlowDescription) = "" Or _
                Trim(strWorkFlowType) = "" Or _
                        lBusinessProcessID = 0 _
                Then

                MsgBox("You must provide an appropriate Business Process Name, an associated" & _
                Chr(10) & " Business Process Description, and select an existing Work Flow Type." _
                                , MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            Dim strQueryShort As String

            If lWorkFlowID = 0 Then
                strQueryShort = _
                "SELECT * FROM BusinessProcessWorkFlows WHERE (BusinessProcessID = " _
                                & lBusinessProcessID & _
                                " AND WorkFlowName = '" & strWorkFlowName & _
                                "')"
            Else
                strQueryShort = _
                "SELECT * FROM BusinessProcessWorkFlows WHERE (BusinessProcessID = " _
                & lBusinessProcessID & _
                " AND WorkFlowName = '" & strWorkFlowName & _
                "') OR WorkFlowID = " & lWorkFlowID
            End If

            'Check if there is an existing series with this name
            If Find(strQueryShort, _
                False) = True Then

                'confirm Update
                If MsgBox("This Business Process Work Flow already exists." & _
                    Chr(10) & "Do you want to update the  details?", _
                            MsgBoxStyle.YesNo, _
                                "iManagement - Record Exists. Update") = _
                                    MsgBoxResult.No Then

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If


                'If no position update is to be made. Position update is made exclusively from record update
                If lMoveDownPosition = 0 And lMoveUpPosition = 0 Then

                    Update("UPDATE BusinessProcessWorkFlows SET " & _
                        "BusinessProcessID = " & BusinessProcessID & _
                        " , WorkFlowName = '" & Trim(strWorkFlowName) & _
                        "', WorkFlowDescription = '" & _
                        Trim(strWorkFlowDescription) & _
                        "', WorkFlowStatus = " & bWorkFlowStatus & _
                        ", HoursForCompletion = " & lHoursForCompletion & _
                        ", WorkFlowType = '" & strWorkFlowType & _
                        "' WHERE BusinessProcessID = " & lBusinessProcessID & _
                        " AND WorkFlowName = '" & Trim(strWorkFlowName) & "'", _
                        True)

                    Return True

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If


                '==================================================================
                If lMoveDownPosition <> 0 Or lMoveUpPosition <> 0 Then

                    Dim lCurrentPosition As Long 'Belonging to record selected to be moved
                    Dim lUpdatePosition As Long
                    Dim lCurrentWorkFlowID As Long
                    Dim arUpdateRecs() As String
                    Dim lNewBusinessProcessID As Long
                    Dim lNewWorkFlowID As Long


                    'Get the current position
                    Find("SELECT * FROM BusinessProcessWorkFlows WHERE (BusinessProcessID = " _
                                    & lBusinessProcessID & _
                                    " AND WorkFlowName = '" & strWorkFlowName & _
                                    "') OR WorkFlowID = " & lWorkFlowID, True)

                    lCurrentPosition = lWorkFlowPosition
                    lCurrentWorkFlowID = lWorkFlowID


                    '----------------------------------lMoveUpPosition
                    '----------------------------------------------
                    '*******************************************
                    If lMoveUpPosition <> 0 Then

                        '-If up,


                        '--'Update position of the lower number to be of the retrieved position number
                        '--'Update position of the current record to number of the lower number
                        '--exit
                        '-End Up if

                        '--'if the current position = 1 then exit
                        If lCurrentPosition = 1 Then

                            objLogin = Nothing
                            datSaved = Nothing

                            Exit Function
                        End If

                        '--'If there is no other value lower than the current position exit
                        If Find("SELECT * FROM BusinessProcessWorkFlows WHERE BusinessProcessID = " _
                                        & lBusinessProcessID & _
                                        " AND WorkFlowPosition < " & lCurrentPosition, True) = False Then

                            objLogin = Nothing
                            datSaved = Nothing

                            Exit Function
                        End If


                        'Get the update position
                        arUpdateRecs = objLogin.FillArray(strOrgAccessConnString, _
                                    "SELECT WorkFlowPosition FROM BusinessProcessWorkFlows " & _
                                                        " WHERE BusinessProcessID = " & _
                                                        lBusinessProcessID & _
                                                        " AND WorkFlowPosition < " & lCurrentPosition & _
                                                        " ORDER By WorkFlowPosition DESC", "", "")

                        If arUpdateRecs Is Nothing Then

                            objLogin = Nothing
                            datSaved = Nothing

                            Exit Function
                        End If


                        'Get the update position
                        lUpdatePosition = CLng(Val(arUpdateRecs(0)))

                        If lUpdatePosition = 0 Then
                            objLogin = Nothing
                            datSaved = Nothing

                            Exit Function
                        End If

                        'Lower Number Update
                        Update("UPDATE BusinessProcessWorkFlows SET " & _
                            "WorkFlowPosition = " & lCurrentPosition & _
                                " WHERE BusinessProcessID = " _
                                    & lBusinessProcessID & _
                                    " AND WorkFlowPosition = " & _
                                        lUpdatePosition, False)


                        'Current Number Update
                        Update("UPDATE BusinessProcessWorkFlows SET " & _
                            "WorkFlowPosition = " & lUpdatePosition & _
                                " WHERE BusinessProcessID = " _
                                    & lBusinessProcessID & _
                                    " AND WorkFlowID = " & _
                                        lCurrentWorkFlowID, False)

                        Return True

                        objLogin = Nothing
                        datSaved = Nothing

                        Exit Function

                    End If

                    '----------------------------------
                    '----------------------------------------------


                    '----------------------------------
                    '----------------------------------------------
                    If lMoveDownPosition <> 0 Then

                        '-If down,
                        '--'If there is no other value higher than the current position exit
                        '--'Update position of the higher number to be of the retrieved position number
                        '--'Update position of the current record to number of the higher number
                        '--exit
                        '-End Up if



                        '--'If there is no other value lower than the current position exit
                        If Find("SELECT * FROM BusinessProcessWorkFlows WHERE BusinessProcessID = " _
                                        & lBusinessProcessID & _
                                        " AND WorkFlowPosition > " & lCurrentPosition, True) = False Then

                            objLogin = Nothing
                            datSaved = Nothing

                            Exit Function
                        End If




                        With objLogin

                            arUpdateRecs = .FillArray(strOrgAccessConnString, _
                                                "SELECT WorkFlowPosition FROM BusinessProcessWorkFlows " & _
                                                                    " WHERE BusinessProcessID = " & _
                                                                    lBusinessProcessID & _
                                                                    " AND WorkFlowPosition > " & lCurrentPosition & _
                                                                    " ORDER By WorkFlowPosition ASC", "", "")
                        End With


                        If arUpdateRecs Is Nothing Then

                            objLogin = Nothing
                            datSaved = Nothing

                            Exit Function
                        End If


                        'Get the update position
                        lUpdatePosition = CLng(Val(arUpdateRecs(0)))

                        If lUpdatePosition = 0 Then
                            objLogin = Nothing
                            datSaved = Nothing

                            Exit Function
                        End If


                        'Lower Number Update
                        Update("UPDATE BusinessProcessWorkFlows SET " & _
                            "WorkFlowPosition = " & lCurrentPosition & _
                                " WHERE BusinessProcessID = " _
                                    & lBusinessProcessID & _
                                    " AND WorkFlowPosition = " & _
                                        lUpdatePosition, False)


                        'Current Number Update
                        Update("UPDATE BusinessProcessWorkFlows SET " & _
                            "WorkFlowPosition = " & lUpdatePosition & _
                                " WHERE BusinessProcessID = " _
                                    & lBusinessProcessID & _
                                    " AND WorkFlowID = " & _
                                        lCurrentWorkFlowID, False)

                        Return True

                        objLogin = Nothing
                        datSaved = Nothing

                        Exit Function


                    End If

                    '----------------------------------if MoveDownEnd
                    '----------------------------------------------

                End If
                'If lMoveDownPosition <> 0 Or lMoveUpPosition <> 0 end
                '==================================================================


            End If


            'Confirm Addition
            If MsgBox("Do you want to add this new Business Process Work Flow?" _
                , MsgBoxStyle.YesNo, _
                    "iManagement - Add the Work Flow?") _
                        = MsgBoxResult.No Then

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO BusinessProcessWorkFlows (" & _
                    "BusinessProcessID," & _
                    "WorkFlowName," & _
                    "WorkFlowDescription," & _
                    "WorkFlowStatus," & _
                    "HoursForCompletion," & _
                    "WorkFlowType," & _
                    "WorkFlowPosition" & _
                            ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lBusinessProcessID & _
                        ",'" & Trim(strWorkFlowName) & _
                        "','" & Trim(strWorkFlowDescription) & _
                        "'," & bWorkFlowStatus & _
                        "," & lHoursForCompletion & _
                        ",'" & strWorkFlowType & _
                        "'," & _
            objLogin.ReturnMaxLongValue _
                (strOrgAccessConnString, _
    "SELECT Max(WorkFlowPosition) As MaxWFPos FROM " & _
        "BusinessProcessWorkFlows" & _
        " WHERE BusinessProcessID = " & lBusinessProcessID) + 1 & _
                            ")"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                Return True

                If DisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - Work Flow Details Saved")

                End If
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save Work Flow details' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Work Flow Details Failed")
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

    End Function

    'Find Informaiton
    Public Function Find(ByVal strQuery As String, _
        ByVal bReturnValues As Boolean) As Boolean

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
                        Return False
                        objLogin = Nothing
                        datRetData = Nothing


                        Exit Function

                    End If


                    For Each myDataRows In myDataTables.Rows

                        If bReturnValues = True Then

                            If IsDBNull(myDataRows("BusinessProcessID")) _
                                = False Then
                                lBusinessProcessID = _
                                        myDataRows("BusinessProcessID")
                            End If


                            If IsDBNull(myDataRows("WorkFlowID")) _
                                = False Then
                                lWorkFlowID = myDataRows("WorkFlowID").ToString
                            End If


                            If IsDBNull(myDataRows("WorkFlowName")) _
                                = False Then
                                strWorkFlowName = _
                                        myDataRows("WorkFlowName").ToString
                            End If


                            If IsDBNull(myDataRows("WorkFlowDescription")) _
                               = False Then
                                strWorkFlowDescription = _
                                    myDataRows("WorkFlowDescription").ToString
                            End If


                            If IsDBNull(myDataRows("WorkFlowStatus")) _
                               = False Then
                                bWorkFlowStatus = _
                                        myDataRows("WorkFlowStatus")
                            End If


                            If IsDBNull(myDataRows("WorkFlowPosition")) _
                               = False Then
                                lWorkFlowPosition = _
                                        myDataRows("WorkFlowPosition")
                            End If


                            If IsDBNull(myDataRows("WorkFlowType")) _
                               = False Then
                                strWorkFlowType = _
                                        myDataRows("WorkFlowType")
                            End If


                            If IsDBNull(myDataRows("HoursForCompletion")) _
                             = False Then
                                lHoursForCompletion = _
                                        myDataRows("HoursForCompletion")

                            End If

                        End If

                    Next

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

        Catch ex As Exception

        End Try

    End Function

    'Delete data
    Public Sub Delete()

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try

            If lBusinessProcessID = 0 Or lWorkFlowID = 0 _
                Then

                MsgBox("Cannot Delete. Please select an existing Work Flow Details.", _
                    MsgBoxStyle.Exclamation, _
                    "iManagement - invalid or incomplete information")

                objLogin = Nothing
                datDelete = Nothing

                Exit Sub

            End If


            If MsgBox("Are you sure you want to delete this record?" _
                    , MsgBoxStyle.YesNo, "iManagement - Delete Record?") _
                        = MsgBoxResult.No Then

                objLogin = Nothing
                datDelete = Nothing

                Exit Sub
            End If

            strDeleteQuery = "DELETE * FROM BusinessProcessWorkFlows" & _
                            " WHERE " & _
                            "BusinessProcessWorkFlows.BusinessProcessID = " _
                            & lBusinessProcessID & _
                                " AND BusinessProcessWorkFlows.WorkFlowID = " _
                                    & lWorkFlowID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully.", MsgBoxStyle.Information, _
                    "iManagement - Business Process Details Deleted")
            Else
                MsgBox("'Delete Work Flow' action failed.", _
                    MsgBoxStyle.Exclamation, "Work Flow Details Deletion failed")
            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Update(ByVal strUpQuery As String, _
        ByVal bDisplaySuccess As Boolean)

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
                If bDisplaySuccess = True Then
                    MsgBox("Record Updated Successfully.", MsgBoxStyle.Information, _
                        "iManagement -  Business Process Details Updated")
                End If
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub

#End Region

End Class
