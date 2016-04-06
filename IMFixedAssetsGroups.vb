
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMFixedAssetsGroups


#Region "PrivateVariables"

    Private lFixedAssetGroupID As Long
    Private strFixedAssetGroupName As String '1 Double-Declining; 2 Single-Line; 3 Sum-of-Year Digits
    Private strFixedAssetGroupDescription As String
    Private dtDateCreated As Date

#End Region


#Region "Properties"

    Public Property ReturnError() As Long

        Get
            Return ReturnError
        End Get

        Set(ByVal Value As Long)
            ReturnError = Value
        End Set

    End Property

    Public Property FixedAssetGroupID() As Long

        Get
            Return lFixedAssetGroupID
        End Get

        Set(ByVal Value As Long)
            lFixedAssetGroupID = Value
        End Set

    End Property

    Public Property FixedAssetGroupName() As String

        Get
            Return strFixedAssetGroupName
        End Get

        Set(ByVal Value As String)
            strFixedAssetGroupName = Value
        End Set

    End Property

    Public Property FixedAssetGroupDescription() As String

        Get
            Return strFixedAssetGroupDescription
        End Get

        Set(ByVal Value As String)
            strFixedAssetGroupDescription = Value
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

            If Trim(strFixedAssetGroupDescription) = "" _
                Or Trim(strFixedAssetGroupName) = "" Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please provide the following details in" & _
                Chr(10) & " order to save the Fixed Asset Group's details: " & _
                Chr(10) & "1. Fixed Asset Group Description " & _
                Chr(10) & "2. Fixed Asset Group Name "

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM FixedAssetsGroups " & _
                " WHERE (FixedAssetGroupName = '" & _
                strFixedAssetGroupName & "')", _
                False) = False Then

                ReturnError += "The specified Fixed Asset Group has " & _
                    "already been saved"

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO FixedAssetsGroups (" & _
                "FixedAssetGroupName," & _
                "FixedAssetGroupDescription" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & strFixedAssetGroupName & _
                    "','" & strFixedAssetGroupDescription & _
                            "')"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)


            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccess = True Then
                    ReturnSuccess += "Fixed Assets Group Saved Successfully."

                End If
            Else

                If DisplayFailure = True Then
                    ReturnError = "'Save Fixed Assets Group' action failed." & _
            Chr(10) & " Make sure all mandatory details are entered"

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If

            objLogin = Nothing
            datSaved = Nothing

            Return True

        Catch ex As Exception
                If DisplayErrorMessages = True Then
                    ReturnError += ex.Message.ToString

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

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bQuerySuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strQuery, datRetData)

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

                            lFixedAssetGroupID = _
                                myDataRows("FixedAssetGroupID")
                            strFixedAssetGroupName = _
                                myDataRows("FixedAssetGroupName")
                            strFixedAssetGroupDescription = _
                                myDataRows("FixedAssetGroupDescription")
                            dtDateCreated = _
                                myDataRows("DateCreated")


                        Next

                    End If

                Next
                Return True

            Else
                Return False

            End If

            datRetData = Nothing
            objLogin = Nothing

        Catch ex As Exception
            ReturnError += ex.Message.ToString

        End Try

    End Function

    Public Function Delete() As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = "DELETE * FROM FixedAssetsGroups" & _
            " WHERE FixedAssetGroupName = '" _
                & strFixedAssetGroupName & "'"

            If strFixedAssetGroupName = 0 Then

                ReturnError += "You must provide a fixed asset group " & _
                        "to delete"
                Exit Function
            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                ReturnSuccess += "Fixed Assets Group details deleted"
                Return True

            Else

                ReturnError += "'Delete Fixed Assets Group' action failed"

            End If

            datDelete = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function Update(ByVal strUpQuery As String, _
        ByVal DisplayErrorMessages As Boolean, _
            ByVal DisplayConfirmation As Boolean, _
                ByVal DisplayFailure As Boolean, _
                    ByVal DisplaySuccess As Boolean) As Boolean

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If Trim(strFixedAssetGroupName) = "" Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery _
                        (strOrgAccessConnString, _
                                    strUpdateQuery, datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    If DisplaySuccess = True Then
                        ReturnSuccess += "Fixed Asset's Group record " & _
                        "updated successfully"
                        Return True

                    End If

                End If

            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try

    End Function

#End Region

    End Class
