Option Explicit On 
'Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMSequenceMaster

#Region "PrivateVariables"

    Private lSequenceGroupID As Long
    Private strSequenceTitle As String
    Private dtDateCreated As Date

#End Region

#Region "Properties"

    Public Property SequenceGroupID() As Long

        Get
            Return lSequenceGroupID
        End Get

        Set(ByVal Value As Long)
            lSequenceGroupID = Value
        End Set

    End Property

    Public Property SequenceTitle() As String

        Get
            Return strSequenceTitle
        End Get

        Set(ByVal Value As String)
            strSequenceTitle = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "DatabaseProcedures"

    'Save informaiton
    Public Function Save(ByVal bDisplayErrorMessages As Boolean, _
            ByVal bDisplaySuccessMessages As Boolean, _
                ByVal bDisplayFailureMessages As Boolean) As Boolean

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

            If Trim(strSequenceTitle) = "" Then

                MsgBox("You must provide an appropriate sequnce " & _
                "text, e.g. AA, A, BB, BZ, etc." _
                , MsgBoxStyle.Critical, _
                "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If IsNumeric(strSequenceTitle) = True Then
                MsgBox("The sequence title cannot be a numeric number or start with a number.", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - ivalid or incomplete information")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if there is an existing series with this name
            If Find("SELECT * FROM SequenceMaster WHERE SequenceTitle = '" _
                    & Trim(strSequenceTitle) & "'", _
                        False) = True Then

                MsgBox("Cannot Add. There is an existing sequence with this name.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Record Exists")

                objLogin = Nothing
                datSaved = Nothing
                Exit Function

            End If


            strInsertInto = "INSERT INTO SequenceMaster " & _
                "(SequenceTitle" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & strSequenceTitle & _
                    "')"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If bDisplaySuccessMessages = True Then
                    MsgBox("New Sequence Saved Successfully.", _
                        MsgBoxStyle.Information, _
                            "iManagement - Record Saved")

                End If

                Return True
            Else

                If bDisplayFailureMessages = True Then
                    MsgBox("'Save New Sequence' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Saving Details Failed")
                End If

            End If

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception
            If bDisplayErrorMessages = True Then
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


                    If bReturnValues = True Then
                        For Each myDataRows In myDataTables.Rows

                            If bReturnValues = True Then
                                lSequenceGroupID = _
                                        myDataRows("SequenceGroupID")
                                strSequenceTitle = _
                                        myDataRows("SequenceTitle").ToString
                                dtDateCreated = _
                                       myDataRows("DateCreated")

                            End If

                        Next
                    End If
                Next

                Return True
            Else
                Return False
            End If

            objLogin = Nothing
            datRetData = Nothing

        Catch ex As Exception

        End Try

    End Function

    'Delete data
    Public Function Delete() As Boolean

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try

            If lSequenceGroupID = 0 Then

                MsgBox("Cannot Delete. Please select an existing " & _
                    "Sequence.", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - invalid or incomplete information")

                datDelete = Nothing
                objLogin = Nothing
                Exit Function

            End If

            If MsgBox("This will Delete all the existing Sequence's " & _
            " records (ONLY) that have this particular sequence", _
                MsgBoxStyle.YesNo, _
                    "iManagement - Delete the record?") = _
                        MsgBoxResult.No Then

                datDelete = Nothing
                objLogin = Nothing
                Exit Function

            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()
            'objLogin.BeginTheTrans()

            strDeleteQuery = "DELETE * FROM SequenceMaster " & _
            " WHERE SequenceGroupID = " & lSequenceGroupID


            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strDeleteQuery, _
                        datDelete)


            strDeleteQuery = "DELETE * FROM Sequence " & _
            " WHERE SequenceGroupID = " & lSequenceGroupID


            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strDeleteQuery, _
                        datDelete)

            'objLogin.CommitTheTrans()

            objLogin.CloseDb()


            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", _
                    MsgBoxStyle.Information, _
                    "iManagement - Sequence Deleted")
                Return True
            Else
                MsgBox("'Sequence delete' action failed", _
                    MsgBoxStyle.Exclamation, _
                        " Chart Of Account Format Deletion failed")
            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function Update(ByVal strUpQuery As String) As Boolean

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
                MsgBox("Sequence Record Updated Successfully", _
                    MsgBoxStyle.Information, _
                    "iManagement -  Details Updated")
                Return True
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Function

#End Region


End Class