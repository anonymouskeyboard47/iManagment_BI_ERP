
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCostCentreTypes


#Region "PrivateVariables"

    Private lCostCentreTypeID As Long
    Private strCostCentreTypeName As String
    Private strCostCentreTypeDescription As String
    Private lCostCentreTypeHierarchyPosition As String
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

    Public Property CostCentreTypeID() As Long

        Get
            Return lCostCentreTypeID
        End Get

        Set(ByVal Value As Long)
            lCostCentreTypeID = Value
        End Set

    End Property

    Public Property CostCentreTypeDescription() As String

        Get
            Return strCostCentreTypeDescription
        End Get

        Set(ByVal Value As String)
            strCostCentreTypeDescription = Value
        End Set

    End Property

    Public Property CostCentreTypeName() As String

        Get
            Return strCostCentreTypeName
        End Get

        Set(ByVal Value As String)
            strCostCentreTypeName = Value
        End Set

    End Property

    Public Property CostCentreTypeHierarchyPosition() As Long

        Get
            Return lCostCentreTypeHierarchyPosition
        End Get

        Set(ByVal Value As Long)
            lCostCentreTypeHierarchyPosition = Value
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

    Public Function CalculateNextHierarchyPosition _
        (ByVal BusinessProcessID As Long) As Long

        Try

            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String

            Dim objLogin As IMLogin = New IMLogin

            With objLogin

                MyMaxValue = .FillArray(strOrgAccessConnString, _
                            "SELECT Max(CostCentreTypeHierarchyPosition) " & _
                                " AS TotalRecords FROM" & _
                                    " CostCentreTypes", "", "")
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

            Return MaxValue

        Catch ex As Exception

            ReturnError += ex.Message.ToString

        End Try

    End Function

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

            If Trim(strCostCentreTypeName) = "" _
                Or Trim(strCostCentreTypeDescription) = "" Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please provide the following details in" & _
                Chr(10) & " order to save the Cost Centre : " & _
                Chr(10) & "1. The Cost Centre Type name " & _
                Chr(10) & "2. The Cost Centre Type Description "

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM CostCentreTypes " & _
                " WHERE (CostCentreTypeName = '" & strCostCentreTypeName & _
                "')", _
                False) = False Then

                ReturnError += "The specified Cost Centre Name exists "

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO CostCentreTypes (" & _
                "CostCentreTypeName, " & _
                "CostCentreTypeDescription," & _
                "CostCentreTypeHierarchyPosition" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & strCostCentreTypeName & _
                    "','" & strCostCentreTypeDescription & _
                    "'," & lCostCentreTypeHierarchyPosition & _
                            ")"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, datSaved)


            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccess = True Then
                    ReturnSuccess += "Cost Centre Type" & _
                        " saved Successfully."

                End If
            Else

                If DisplayFailure = True Then
                    ReturnError = "'Save Cost Centre Type' " & _
                        "action failed." & _
            Chr(10) & " Make sure all mandatory details are entered."

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

                            lCostCentreTypeID = _
                                myDataRows("CostCentreTypeID")
                            lCostCentreTypeHierarchyPosition = _
                                myDataRows("CostCentreTypeHierarchyPosition")
                            strCostCentreTypeDescription = _
                                myDataRows("CostCentreTypeDescription")
                            strCostCentreTypeName = _
                                myDataRows("CostCentreTypeName")
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

            strDeleteQuery = "DELETE * FROM CostCentreTypes" & _
            " WHERE CostCentreTypeName = '" & strCostCentreTypeName & _
            "'"

            If Trim(strCostCentreTypeName) = "" Then

                ReturnError += "You must provide a Cost Centre type " & _
                    "in order to delete it"

                Exit Function

            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strDeleteQuery, datDelete)

            objLogin.CloseDb()


            If bDelSuccess = True Then

                ReturnSuccess += "Cost Centre Type deleted succesfully"

                Return True

            Else

                ReturnError += "'Delete Cost Centre Type" & _
                    "' action failed"

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

            If Trim(strCostCentreTypeName) = "" Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery _
                        (strOrgAccessConnString, _
                                    strUpdateQuery, datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    If DisplaySuccess = True Then
                        ReturnSuccess += "Cost Centre Type record " & _
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
