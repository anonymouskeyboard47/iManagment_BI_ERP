
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMFixedAssetGroupToAssetLink


#Region "PrivateVariables"

    Private lFixedAssetGroupID As Long
    Private lFixedAssetID As Long
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

    Public Property FixedAssetID() As Long

        Get
            Return lFixedAssetID
        End Get

        Set(ByVal Value As Long)
            lFixedAssetID = Value
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

            If lFixedAssetGroupID = 0 _
                Or lFixedAssetID = 0 Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please provide the following details in" & _
                Chr(10) & " order to save the Fixed Asset : " & _
                Chr(10) & "1. The Fixed Asset Group you want to use " & _
                Chr(10) & "2. The Fixed Asset you want to link to the Fixed Asset Group"

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM FixedAssetGroupToAssetLink " & _
                " WHERE (FixedAssetGroupID = " & lFixedAssetGroupID & _
                " AND FixedAssetID = " & lFixedAssetID & ")", _
                False) = False Then

                ReturnError += "The specified Fixed Asset has already " & _
                    "been linked to this Fixed Asset Group "

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO FixeAssetGroupToAssetLink (" & _
                "FixedAssetGroupID, " & _
                "FixedAssetID" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lFixedAssetGroupID & _
                    "," & lFixedAssetID & _
                            ")"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, datSaved)


            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccess = True Then
                    ReturnSuccess += "Fixed Asset-Fixed Group linkage" & _
                        " saved Successfully."

                End If
            Else

                If DisplayFailure = True Then
                    ReturnError = "'Save Fixed Asset-Fixed Group linkage' " & _
                        "action failed." & _
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
                            lFixedAssetID = _
                                myDataRows("FixedAssetID")
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

            strDeleteQuery = "DELETE * FROM FixeAssetGroupToAssetLink" & _
            " WHERE FixedAssetGroupID = " & lFixedAssetGroupID & _
            " AND FixedAssetID = " & lFixedAssetID

            If lFixedAssetGroupID = 0 Or lFixedAssetID = 0 Then

                ReturnError += "You must provide a Fixed Asset Group " & _
                        " and the Fixed Asset that you want to delete " & _
                            "from the group"

                Exit Function
            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strDeleteQuery, datDelete)

            objLogin.CloseDb()


            If bDelSuccess = True Then

                ReturnSuccess += "Fixed Asset-Fixed Group linkage " & _
                    "details deleted"

                Return True

            Else

                ReturnError += "'Delete Fixed Asset-Fixed Group " & _
                    "linkage' action failed"

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

            ReturnError += " This particular set of record details " & _
                " cannot be updated"
         
        Catch ex As Exception

        End Try

    End Function


#End Region

End Class
