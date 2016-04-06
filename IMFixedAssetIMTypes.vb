
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMFixedAssetIMTypes

#Region "PrivateVariables"

    Private lFixedAssetTypeID As Long
    Private strFixedAssetTypeName As String
    Private strFixedAssetTypeDescription As String
    Private dtDateCreated As Date

#End Region

#Region "Properties"

    Public Property FixedAssetTypeID() As Long

        Get
            Return lFixedAssetTypeID
        End Get

        Set(ByVal Value As Long)
            lFixedAssetTypeID = Value
        End Set

    End Property

    Public Property FixedAssetTypeName() As String

        Get
            Return strFixedAssetTypeName
        End Get

        Set(ByVal Value As String)
            strFixedAssetTypeName = Value
        End Set

    End Property

    Public Property FixedAssetTypeDescription() As String

        Get
            Return strFixedAssetTypeDescription
        End Get

        Set(ByVal Value As String)
            strFixedAssetTypeDescription = Value
        End Set

    End Property

    Public Property DateCreated() As String

        Get
            Return dtDateCreated
        End Get

        Set(ByVal Value As String)
            dtDateCreated = Value
        End Set

    End Property

    Public Property ReturnError() As Long

        Get
            Return ReturnError
        End Get

        Set(ByVal Value As Long)
            ReturnError = Value
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

            If Trim(strFixedAssetTypeDescription) = "" _
                Or Trim(strFixedAssetTypeName) = ""  Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please provide the following details in" & _
                " order to save the Fixed Asset Type details: " & _
                Chr(10) & "1. Fixed Asset Type Name (e.g. Plot, Vehicle," & _
                " Computer Notebooks, etc) " & _
                Chr(10) & "2. Fixed Asset Type Description (e.g for " & _
                "Computer Notebooks, you may have the desciption - Also " & _
                "known as laptops. Portable computers "

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM FixedAssetIMTypes " & _
                " WHERE (FixedAssetTypeName = '" & strFixedAssetTypeName & _
                "' OR FixedAssetTypeID = " & lFixedAssetTypeID & ")", _
                False) = True Then

                ReturnError += "The specified Fixed Asset Type"

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO FixedAssetIMTypes (" & _
                "FixedAssetTypeName," & _
                "FixedAssetTypeDescription" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & strFixedAssetTypeName & _
                    "','" & strFixedAssetTypeDescription & _
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
                    ReturnSuccess += "Fixed Asset Type Saved Successfully."

                End If
            Else

                If DisplayFailure = True Then
                    ReturnError = "'Save Fixed Asset Type' action failed." & _
            " Make sure all mandatory details are entered"

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

                            lFixedAssetTypeID = _
                                myDataRows("FixedAssetTypeID")
                            strFixedAssetTypeName = _
                                myDataRows("FixedAssetTypeName")
                            strFixedAssetTypeDescription = _
                                myDataRows("FixedAssetTypeDescription")
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

            strDeleteQuery = "DELETE * FROM FixedAssetIMTypes " & _
            "WHERE FixedAssetTypeID =  

            If lFixedAssetTypeID = 0 Then

                ReturnError += "You must provide a fixed " & _
                    "asset type you want to delete "
                        
                Exit Function
            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                ReturnSuccess += "Fixed Asset Type details deleted"
                Return True
            Else

                ReturnError += "'Delete Fixed Asset Type' action failed"


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

            If lFixedAssetTypeID = 0 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strUpdateQuery, datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    If DisplaySuccess = True Then
                        ReturnSuccess += " Fixed Asset Type record " & _
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
