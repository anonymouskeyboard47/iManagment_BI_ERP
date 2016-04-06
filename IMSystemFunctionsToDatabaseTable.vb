Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMSystemFunctionsToDatabaseTable
    Inherits IMSystemWorkflowObjects
    
#Region "PrivateVariables"

    Private strTableName As String
    Private bReadTable As Boolean
    Private bInsertTable As Boolean
    Private bUpdateTable As Boolean
    Private bDeleteTable As Boolean
    Private bReadTableSchema As Boolean

#End Region

#Region "Properties"

    Public Property TableName() As String

        Get
            Return strTableName
        End Get

        Set(ByVal Value As String)
            strTableName = Value
        End Set

    End Property

    Public Property ReadTable() As Boolean

        Get
            Return bReadTable
        End Get

        Set(ByVal Value As Boolean)
            bReadTable = Value
        End Set

    End Property

    Public Property InsertTable() As Boolean

        Get
            Return bInsertTable
        End Get

        Set(ByVal Value As Boolean)
            bInsertTable = Value
        End Set

    End Property

    Public Property UpdateTable() As Boolean

        Get
            Return bUpdateTable
        End Get

        Set(ByVal Value As Boolean)
            bUpdateTable = Value
        End Set

    End Property

    Public Property DeleteTable() As Boolean

        Get
            Return bDeleteTable
        End Get

        Set(ByVal Value As Boolean)
            bDeleteTable = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Function WkFlTableSave _
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


        If Trim(ObjectType) = "" Or _
            WorkFlowID = 0 Or _
                BusinessProcessID = 0 Or _
                    Trim(strTableName) = "" _
                            Then

            MsgBox("Please provide an existing:" & _
            Chr(10) & "1. Object Type e.g. Table" & _
            Chr(10) & "2. Business Process" & _
            Chr(10) & "3. Business Process Work Flow" & _
            Chr(10) & "4. Table Name" _
                        , MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

            objLogin = Nothing
            datSaved = Nothing

            Exit Function
        End If


        If Find("SELECT * FROM SystemWorkFlowObjects WHERE " & _
            " ObjectId = " & ObjectID, False) = False Then

            ObjectID = objLogin.ReturnMaxLongValue(strOrgAccessConnString, _
            "SELECT Max(ObjectID) FROM SystemWorkFlowObjects WHERE " & _
        " WorkFlowId = " & WorkFlowID & _
        " AND BusinessProcessID = " & BusinessProcessID) + 1


            If Save(False, False, False, False) = False Then
                MsgBox("Invalid Table Rights details provided. Please contact the Systems Administrator", _
                    MsgBoxStyle.Critical, "iManagement - Invalid or incomplete information provided")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If

        End If


        If DisplayConfirm = True Then
            'Check if there is an existing series with this name
            If WkFlTableFind("SELECT * FROM SystemFunctionsToDatabaseTable" & _
            " INNER Join SystemWorkFlowObjects ON " & _
            "  SystemWorkFlowObjects.ObjectID = " & _
            " SystemFunctionsToDatabaseTable.ObjectID " & _
            " WHERE (SystemWorkFlowObjects.ObjectID = " & ObjectID & _
            ") OR (TableName = '" & strTableName & _
            "' AND BusinessProcessID = " & BusinessProcessID & _
            " AND WorkFlowID = " & WorkFlowID & ")" _
                , False, False) = True Then

                If MsgBox("The Table Rights already exists." & _
                Chr(10) & "Do you want to update the details?", _
                        MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                MsgBoxResult.Yes Then

                    WkFlTableUpdate("UPDATE SystemFunctionsToDatabaseTable SET " & _
                                ", ReadTable = " & bReadTable & _
                                ", InsertTable = " & bInsertTable & _
                                ", UpdateTable = " & bUpdateTable & _
                                ", DeleteTable = " & bDeleteTable & _
                                ", ReadTableSchema = " & bReadTableSchema & _
                                " WHERE  ObjectID = " & ObjectID)

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If
        End If


        If DisplayConfirm = True Then
            If MsgBox("Are you sure you want to add new Table Rights?" _
            , MsgBoxStyle.YesNo, "iManagment - Add Table Rights?") _
            = MsgBoxResult.No Then
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If
        End If

        strInsertInto = "INSERT INTO SystemFunctionsToDatabaseTable (" & _
            "ObjectID," & _
            "TableName," & _
            "ReadTable," & _
            "InsertTable," & _
            "UpdateTable," & _
            "DeleteTable," & _
            "ReadTableSchema" & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                "(" & ObjectID & _
                ",'" & Trim(strTableName) & _
                "'," & bReadTable & _
                "," & bInsertTable & _
                "," & bUpdateTable & _
                "," & bDeleteTable & _
                "," & bReadTableSchema & _
                        ")"

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
        strSaveQuery, _
        datSaved)

        objLogin.CloseDb()

        If bSaveSuccess = True Then
            If DisplaySuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Table Rights Details Saved")

            End If

            objLogin = Nothing
            datSaved = Nothing

            Return True

        Else

            If DisplayFailure = False Then
                MsgBox("'Save Table Rights' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Table Rights Addition Failed")
            End If

        End If
        objLogin = Nothing
        datSaved = Nothing

    End Function

    Public Function WkFlTableFind(ByVal strQuery As String, _
            ByVal bReturnValues As Boolean, _
                ByVal ReturnWorkFlowObject As Boolean) As Boolean

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
                    Exit Function

                End If

                For Each myDataRows In myDataTables.Rows
                    If bReturnValues = True Then

                        If ReturnWorkFlowObject = True Then
                            ObjectID = myDataRows("SystemWorkFlowObjects.ObjectID")

                            If Find("SELECT * FROM SystemWorkFlowObjects" & _
                           " WHERE ObjectID = " & ObjectID, True) = False Then
                                MsgBox("Critical database error. Missing Work Flow Object identifier", _
                                MsgBoxStyle.Critical, "iManagement - Database Error")

                                Exit Function
                            End If



                        End If

                        strTableName = myDataRows("TableName").ToString
                        bReadTable = myDataRows("ReadTable")
                        bInsertTable = myDataRows("InsertTable")
                        bUpdateTable = myDataRows("UpdateTable")
                        bDeleteTable = myDataRows("DeleteTable")
                        bReadTableSchema = myDataRows _
                            ("SystemFunctionsToDatabaseTable.ReadTableSchema")




                    End If
                Next
            Next

            objLogin = Nothing
            datRetData = Nothing

            Return True
        Else
            Return False
        End If


    End Function

    Public Function WkFlTableDelete(ByVal DisplayError As Boolean, _
        ByVal DisplayConfirm As Boolean, _
            ByVal DisplayFailure As Boolean, ByVal DisplaySuccess As Boolean) As Boolean

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        If ObjectID = 0 Then

            If DisplayError = True Then

                MsgBox("Cannot Delete due to missing information. Please provide an existing" & _
                " Table Name, Business Process, and Organization with the details." _
                            , MsgBoxStyle.Exclamation, _
                            "iManagement - invalid or incomplete information")
            End If

            objLogin = Nothing
            datDelete = Nothing

            Exit Function

        End If


        If DisplayConfirm = True Then
            If MsgBox("Are you sure you want to delete this Table Rights?" _
            , MsgBoxStyle.YesNo, "iManagement - Delete the Table Rights") = MsgBoxResult.No Then

                objLogin = Nothing
                datDelete = Nothing

                Exit Function

            End If
        End If


        strDeleteQuery = "DELETE * FROM SystemWorkFlowObject WHERE " & _
                    " ObjectID = " & ObjectID

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
        datDelete)

        objLogin.CloseDb()

        If bDelSuccess = True Then
            If DisplaySuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Table Rights Details Deleted")
            End If

            Return True

        Else

            If DisplayFailure = True Then
                MsgBox("'Delete Table Rights' action failed", _
                    MsgBoxStyle.Exclamation, " Table Rights Deletion failed")
            End If
        End If

    End Function

    Public Shadows Sub WkFlTableUpdate(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        If ObjectID = 0 Then

            objLogin = Nothing
            datUpdated = datUpdated
            Exit Sub

        End If


        strUpdateQuery = strUpQuery

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                            strUpdateQuery, _
                                    datUpdated)

        objLogin.CloseDb()

        If bUpdateSuccess = True Then
            MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                "iManagement -  Table Rights Details Updated")
        End If



    End Sub

#End Region

End Class
