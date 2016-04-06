
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMSIDMaster

#Region "PrivateVariables"

    Private strSystemSID As String
    Private strType As String
    Private strTypeUserOrGroupID As String
    Private bSIDStatus As Boolean
    Private dtDateRegistered As Date

#End Region


#Region "Properties"

    Public Property SystemSID() As String

        Get
            Return Trim(strSystemSID)
        End Get

        Set(ByVal Value As String)
            strSystemSID = Value
        End Set

    End Property

    Public Property Type() As String

        Get
            Return Trim(strType)
        End Get

        Set(ByVal Value As String)
            strType = Value
        End Set

    End Property

    Public Property TypeUserOrGroupID() As String

        Get
            Return Trim(strTypeUserOrGroupID)
        End Get

        Set(ByVal Value As String)
            strTypeUserOrGroupID = Value
        End Set

    End Property

    Public Property SIDStatus() As Boolean

        Get
            Return bSIDStatus
        End Get

        Set(ByVal Value As Boolean)
            bSIDStatus = Value
        End Set

    End Property

    Public Property DateRegistered() As Date

        Get
            Return dtDateRegistered
        End Get

        Set(ByVal Value As Date)
            dtDateRegistered = Value
        End Set

    End Property

#End Region


#Region "InitializationProcedures"

        Public Sub New()
            MyBase.New()
        End Sub

#End Region

#Region "GeneralProcedures"

    Public Function CalculateNextSID() As String
        Try

            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String
            Dim strProposedSID As String

            MyMaxValue = FillControl(strAccessConnString, _
                           "SELECT Count(*) AS TotalRecords " & _
            " FROM SIDMAster " & _
            " WHERE " & _
            " Day(SIDMAster.DateRegistered) = Day(Now()) " & _
            " AND " & _
            " Year(SIDMAster.DateRegistered) = Year(Now()) " & _
            " AND " & _
            " Month(SIDMAster.DateRegistered)=Month(Now())", "", "")

            If Not MyMaxValue Is Nothing Then
                For Each strItem In MyMaxValue
                    If Not strItem Is Nothing Then

                        MaxValue = CLng(Val(strItem))


                    End If
                Next
            End If

            MaxValue = MaxValue + 1

            strProposedSID = "SID" & Now.Day.ToString _
                & Now.Month.ToString & _
                    Now.Year.ToString & _
                            MaxValue.ToString

            Return strProposedSID

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, _
                "iManagement - System Error")
        End Try

    End Function

#End Region

#Region "DatabaseProcedures"

    Public Function Save _
        (ByVal DisplayConfirm As Boolean, _
            ByVal DisplayError As Boolean, _
                ByVal DisplaySuccess As Boolean, _
                    ByVal DisplayFailure As Boolean) As Boolean
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String


        If Trim(strType) = "" Or _
            Trim(strTypeUserOrGroupID) = "" _
                            Then

            MsgBox("Please provide an existing" & _
            Chr(10) & "1. Security ID Type" _
                        , MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

            objLogin = Nothing
            datSaved = Nothing

            Exit Function

        End If



        'Check if there is an existing series with this name
        If Find("SELECT * FROM SIDMaster WHERE  SystemSID = '" _
                    & strSystemSID & "'", False) = True Then

            If DisplayConfirm = True Then

                If MsgBox("The Security ID Name already exists." & _
                Chr(10) & "Do you want to update the details?", _
                        MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                MsgBoxResult.Yes Then

                    Update("UPDATE SIDMaster SET " & _
                                          " SIDStatus = " & bSIDStatus & _
                                              " WHERE  SystemSID = '" _
                                                  & Trim(strSystemSID) & "'", False)

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If



        End If





        If DisplayConfirm = True Then
            If MsgBox("Are you sure you want to this new Security ID?" _
            , MsgBoxStyle.YesNo, "iManagment - Add new SID record?") _
            = MsgBoxResult.No Then
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

        End If

        strSystemSID = CalculateNextSID()

        strInsertInto = "INSERT INTO SIDMaster (" & _
            "SystemSID," & _
            "Type," & _
            "TypeUserOrGroupID," & _
            "SIDStatus," & _
            "DateRegistered" & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                "('" & Trim(strSystemSID) & _
                "','" & Trim(strType) & _
                "','" & Trim(strTypeUserOrGroupID) & _
                "'," & bSIDStatus & _
                ",'" & Now() & _
                        "')"

        objLogin.ConnectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
        strSaveQuery, _
        datSaved)

        objLogin.CloseDb()

        If bSaveSuccess = True Then
            If DisplaySuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - SID Details Saved")
            End If

            Return True

        Else

            If DisplayFailure = False Then
                MsgBox("'Save SID' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - SID Addition Failed")
            End If

        End If


    End Function

    Public Function Find(ByVal strQuery As String, _
            ByVal bReturnValues As Boolean) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.ConnectString = strAccessConnString
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
                    If bReturnValues = True Then

                        strSystemSID = myDataRows("SystemSID").ToString()
                        strType = myDataRows("Type").ToString()
                        strTypeUserOrGroupID = myDataRows("TypeUserOrGroupID").ToString()
                        bSIDStatus = myDataRows("SIDStatus")
                        dtDateRegistered = myDataRows("DateRegistered")


                    End If

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Function Delete(ByVal DisplayError As Boolean, _
        ByVal DisplayConfirm As Boolean, _
            ByVal DisplayFailure As Boolean, ByVal DisplaySuccess As Boolean) As Boolean

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin


        If Trim(strSystemSID) = "" Then

            If DisplayError = True Then

                MsgBox("Cannot Delete due to missing information. Please provide an existing" & _
                " Security ID Details." _
                            , MsgBoxStyle.Exclamation, _
                            "iManagement - invalid or incomplete information")
            End If

            objLogin = Nothing
            datDelete = Nothing

            Exit Function

        End If


        If DisplayConfirm = True Then
            If MsgBox("Are you sure you want to delete this user's detaisls?" _
            , MsgBoxStyle.YesNo, "iManagement - Delete the user's details?") = MsgBoxResult.No Then

                objLogin = Nothing
                datDelete = Nothing

                Exit Function

            End If
        End If


        strDeleteQuery = "DELETE * FROM SIDMaster WHERE SystemSID = '" & _
                    strSystemSID & "'"

        objLogin.ConnectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
        datDelete)

        objLogin.CloseDb()

        If bDelSuccess = True Then
            If DisplaySuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Security ID Details Deleted")
            End If

            Return True

        Else

            If DisplayFailure = True Then
                MsgBox("'Delete Security ID' action failed", _
                    MsgBoxStyle.Exclamation, " Security ID Deletion failed")
            End If
        End If

    End Function

    Public Shadows Sub Update(ByVal strUpQuery As String, _
        ByVal bDisplayMessages As Boolean)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If Trim(strSystemSID) <> "" Then

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                If bDisplayMessages = True Then
                    MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                        "iManagement -  SID Master Details Updated")
                End If

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

            objLogin.ConnectString = strAccessConnString
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

