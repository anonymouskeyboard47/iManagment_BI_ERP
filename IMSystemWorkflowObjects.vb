Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMSystemWorkflowObjects
        
#Region "PrivateVariables"

    Private lObjectID As Long
    Private strObjectType As String
    Private lWorkFlowID As Long
    Private lBusinessProcessID As Long
    Private bReadTableSchema As Boolean
    Private dtDateRegistered As Date

#End Region

#Region "Properties"

    Public Property ObjectID() As Long

        Get
            Return lObjectID
        End Get

        Set(ByVal Value As Long)
            lObjectID = Value
        End Set

    End Property

    Public Property ObjectType() As String

        Get
            Return strObjectType
        End Get

        Set(ByVal Value As String)
            strObjectType = Value
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

    Public Property BusinessProcessID() As Long

        Get
            Return lBusinessProcessID
        End Get

        Set(ByVal Value As Long)
            lBusinessProcessID = Value
        End Set

    End Property

    Public Property ReadTableSchema() As Boolean

        Get
            Return bReadTableSchema
        End Get

        Set(ByVal Value As Boolean)
            bReadTableSchema = Value
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

      

        If lObjectID = 0 Or _
            lWorkFlowID = 0 Or _
            lBusinessProcessID = 0 Or _
            Trim(strObjectType) = "" _
                            Then

            If DisplayError = True Then
                MsgBox("Please provide an existing:" & _
                Chr(10) & "1. Business Process" & _
                Chr(10) & "2. Work Flow" & _
                Chr(10) & "3. Object Type" _
                            , MsgBoxStyle.Exclamation, _
                            "iManagement - invalid or incomplete information")
            End If

            objLogin = Nothing
            datSaved = Nothing

            Exit Function

        End If

        If DisplayConfirm = True Then
            'Check if there is an existing series with this name
            If Find("SELECT * FROM SystemWorkFlowObjects WHERE" & _
            " WorkFlowID = " & lWorkFlowID & _
            " AND BusinessProcessID = " & lBusinessProcessID & _
            " AND ObjectID = " & lObjectID _
                , False) = True Then

                If MsgBox("The Security ID Name already exists." & _
                Chr(10) & "Do you want to update the details?", _
                        MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                MsgBoxResult.Yes Then

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If
        End If


        If DisplayConfirm = True Then
            If MsgBox("Are you sure you want to link object to the" & _
            Chr(10) & " Organization and Business Process Work Flow?" _
            , MsgBoxStyle.YesNo, "iManagment - Add new work flow object?") _
            = MsgBoxResult.No Then
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

        End If


        strInsertInto = "INSERT INTO SystemWorkFlowObjects (" & _
            "ObjectType," & _
            "WorkFlowID," & _
            "BusinessProcessID," & _
            "ReadTableSchema," & _
            "ObjectID" & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                "('" & strObjectType & _
                "'," & lWorkFlowID & _
                "," & lBusinessProcessID & _
                "," & bReadTableSchema & _
                "," & lObjectID & _
                        ")"

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
        strSaveQuery, _
        datSaved)

        objLogin.CloseDb()

        objLogin = Nothing
        datSaved = Nothing

        If bSaveSuccess = True Then
            If DisplaySuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Work Flow to Object Details Saved")
            End If

            Return True

        Else

            If DisplayFailure = False Then
                MsgBox("'Save Work Flow Object' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Work Flow to Object Addition Failed")
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

                        lObjectID = myDataRows("ObjectID")
                        strObjectType = myDataRows("ObjectType")
                        lWorkFlowID = myDataRows("WorkFlowID")
                        lBusinessProcessID = myDataRows("BusinessProcessID")
                        bReadTableSchema = myDataRows("ReadTableSchema")
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

        If lObjectID = 0 Then

            If DisplayError = True Then

                MsgBox("Cannot Delete due to missing information. Please provide an existing" & _
                " Work Flow Details, Business Process, and Organization with the details." _
                            , MsgBoxStyle.Exclamation, _
                            "iManagement - invalid or incomplete information")
            End If

            objLogin = Nothing
            datDelete = Nothing

            Exit Function

        End If


        If DisplayConfirm = True Then
            If MsgBox("Are you sure you want to delete this user's details?" _
            , MsgBoxStyle.YesNo, "iManagement - Delete the user's details?") = MsgBoxResult.No Then

                objLogin = Nothing
                datDelete = Nothing

                Exit Function

            End If
        End If


        strDeleteQuery = "DELETE * FROM SystemWorkFlowObject INNER JOIN " & _
        " SystemFunctionsToDatabaseTable ON " & _
        " SystemFunctionsToDatabaseTable.ObjectID = " & _
        " SystemWorkFlowObject.ObjectID WHERE " & _
                    " SystemWorkFlowObject.ObjectID = " & lObjectID

        objLogin.ConnectString = strOrgAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
        datDelete)

        objLogin.CloseDb()

        If bDelSuccess = True Then
            If DisplaySuccess = True Then
                MsgBox("Record Deleted Successfully", _
                    MsgBoxStyle.Information, _
                        "iManagement - Security ID Details Deleted")

            End If

            Return True

        Else

            If DisplayFailure = True Then
                MsgBox("'Delete Security ID' action failed", _
                    MsgBoxStyle.Exclamation, _
                        " Security ID Deletion failed")
            End If
        End If

    End Function

    Public Shadows Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        If Trim(lObjectID) = 0 Or lWorkFlowID = 0 Or _
                lBusinessProcessID = 0 Then

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
                "iManagement -  Customer's Bank Account Details Updated")
        End If



    End Sub

   
#End Region

End Class
