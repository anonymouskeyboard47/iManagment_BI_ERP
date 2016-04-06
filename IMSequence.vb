Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMSequence
    Inherits IMSequenceMaster

#Region "PrivateVariables"

    Private lLastSequenceNo As Long
    Private dtLastEntryDate As Date

#End Region

#Region "Properties"

    Public ReadOnly Property LastSequenceNo() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lLastSequenceNo
        End Get

        'Set(ByVal Value As Long)
        '    lLastSequenceNo = Value
        'End Set

    End Property

    Public ReadOnly Property LastEntryDate() As Date

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return dtLastEntryDate
        End Get

        'Set(ByVal Value As Date)
        '    dtLastEntryDate = Value
        'End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Function SequenceSave(ByVal bDisplayErrorMessages As Boolean, _
            ByVal bDisplaySuccessMessages As Boolean, _
                ByVal bDisplayFailureMessages As Boolean) As Boolean
        'Saves a new sequence group details

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String
        Dim lMyMaxVal As Long

        Try

            If SequenceGroupID = 0 Then

                If bDisplayFailureMessages = True Then
                    MsgBox("Please provide an existing Sequence Group" & _
                                " of the sequence (Series)" & _
                                    " group whose details you want to save.", _
                                    MsgBoxStyle.Critical, "Save Action Failed")
                End If

                datSaved = Nothing
                objLogin = Nothing
                Exit Function
            End If


            '[Check if the SequenceGroupID exists
            'If the Sequence does not exist
            If SequenceFind("SELECT * FROM Sequence " & _
                " WHERE SequenceGroupID = " & SequenceGroupID & _
                " OR SequenceTitle  = '" & SequenceTitle & "'", _
                                    False, False) = False Then

                'If the sequncegroupID does not exist in sequencemaster then
                strInsertInto = "INSERT INTO Sequence (" & _
                    "SequenceGroupID," & _
                    "LastSequenceNo" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "(" & SequenceGroupID & _
                        "," & 1 & _
                                ")"

                objLogin.ConnectString = strAccessConnString
                objLogin.ConnectToDatabase()

                bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                strSaveQuery, _
                datSaved)

                objLogin.CloseDb()

                If bSaveSuccess = True Then
                    If bDisplaySuccessMessages = True Then
                        MsgBox("Sequence Saved Successfully", _
                            MsgBoxStyle.Information, _
                        "iManagement - Record Saved Successfully")

                    End If

                    Return True
                Else

                    If bDisplaySuccessMessages = True Then
                        MsgBox("'Save Sequence' action failed." & _
                            " Make sure all mandatory details are entered", _
                                MsgBoxStyle.Exclamation, _
                                    "iManagement - Sequence Group Addition Failed")

                    End If
                End If

                datSaved = Nothing
                objLogin = Nothing
                Exit Function

            End If


            '[Check if the SequenceGroupID exists
            'If the Sequence exists
            If SequenceFind("SELECT * FROM Sequence " & _
                " WHERE SequenceGroupID = " & SequenceGroupID & _
                " OR SequenceTitle  = '" & SequenceTitle & "'", _
                                    False, False) = True Then


                If ReturnNextSequenceNo(False, False, False) = 0 Then
                    datSaved = Nothing
                    objLogin = Nothing
                    Exit Function
                End If


                SequenceUpdate("UPDATE Sequence SET " & _
                " LastSequenceNo = " & _
                    ReturnNextSequenceNo(False, False, False) & _
                        " WHERE SequenceGroupID = " & SequenceGroupID, _
                            False, False, False)


                datSaved = Nothing
                objLogin = Nothing
                Exit Function

            End If


        Catch ex As Exception
            If bDisplayErrorMessages = False Then
                MsgBox(ex.Message.ToString, _
                    MsgBoxStyle.Exclamation, _
                        "iManagement System Error")

            End If
        End Try

    End Function

    Public Function ReturnNextSequenceNo(ByVal bDisplayErrorMessages As Boolean, _
           ByVal bDisplaySuccessMessages As Boolean, _
               ByVal bDisplayFailureMessages As Boolean) As Long
        'Saves a new sequence group details

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String
        Dim lMyMaxVal As Long

        Try

            If SequenceGroupID = 0 Then

                If bDisplayFailureMessages = True Then
                    MsgBox("Please provide an existing Sequence Group" & _
                                " of the sequence (Series)" & _
                                    " group whose details you want to save.", _
                                    MsgBoxStyle.Critical, "Save Action Failed")
                End If

                datSaved = Nothing
                objLogin = Nothing
                Exit Function

            End If


            'If the sequncegroupID does not exist in sequencemaster then
            lMyMaxVal = objLogin.ReturnMaxLongValue _
                (strOrgAccessConnString, _
                    "SELECT MAX(LastSequenceNo) As LastSeqNo FROM " & _
                    " Sequence WHERE SequenceGroupID = " & _
                    SequenceGroupID) + 1

            Return lMyMaxVal

            datSaved = Nothing
            objLogin = Nothing
            Exit Function


        Catch ex As Exception
            If bDisplayErrorMessages = False Then
                MsgBox(ex.Message.ToString, _
                    MsgBoxStyle.Exclamation, _
                        "iManagement System Error")

            End If
        End Try

    End Function

    Public Function SequenceFind(ByVal strQuery As String, _
                    ByVal bReturnDetails As Boolean, _
                            ByVal bReturnSequenceMasterDetails As Boolean) _
                                As Boolean

        Try

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
                        datRetData = Nothing
                        objLogin = Nothing
                        Return False
                        Exit Function

                    End If

                    If bReturnDetails = True Then
                        For Each myDataRows In myDataTables.Rows

                            lLastSequenceNo = myDataRows("LastSequenceNo")
                            dtLastEntryDate = myDataRows("LastEntryDate")

                            If bReturnSequenceMasterDetails = True Then

                                SequenceGroupID = myDataRows _
                                    ("Sequence.SequenceGroupID")
                                SequenceTitle = myDataRows _
                                    ("SequenceMaster.SequenceTitle").ToString()
                            Else
                                SequenceGroupID = myDataRows _
                                    ("SequenceGroupID")

                            End If
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
                    "iManagement System Error")
        End Try

    End Function

    Public Function SequenceDelete(ByVal strDelQuery As String, _
        ByVal bDisplayErrorMessages As Boolean, _
            ByVal bDisplaySuccessMessages As Boolean, _
                ByVal bDisplayFailureMessages As Boolean) As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            If SequenceGroupID = 0 _
                                Then
                If bDisplayFailureMessages = True Then
                    MsgBox("Some information is missing. Please select an " & _
                    "existing Sequence's Detail.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Cannot Delete")
                End If

                datDelete = Nothing
                objLogin = Nothing
                Exit Function

            End If

            strDeleteQuery = "DELETE * FROM Sequence WHERE " & _
                "SequenceGroupID = " & SequenceGroupID

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                If bDisplaySuccessMessages = True Then
                    MsgBox("Sequence Details Deleted", MsgBoxStyle.Information, _
                        "iManagement - Record Deleted Successfully")
                End If

                Return True

            Else
                If bDisplayFailureMessages = True Then
                    MsgBox("'Delete Sequence' action failed", _
                                        MsgBoxStyle.Exclamation, _
                                            "Deletion failed")
                End If

            End If

            datDelete = Nothing
            objLogin = Nothing

        Catch ex As Exception

        End Try

    End Function

    Public Function SequenceUpdate(ByVal strUpQuery As String, _
        ByVal bDisplayErrorMessages As Boolean, _
            ByVal bDisplaySuccessMessages As Boolean, _
                ByVal bDisplayFailureMessages As Boolean) As Boolean

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If SequenceGroupID = 0 _
                            Then

                If bDisplayFailureMessages = True Then
                    MsgBox("Missing information. Canot update", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Cannot update")
                End If

                objLogin = Nothing
                datUpdated = Nothing
                Exit Function

            End If

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                If bDisplaySuccessMessages = True Then
                    MsgBox("Record Updated Successfully", _
                        MsgBoxStyle.Information, _
                        "iManagement -  Sequence Details Updated")

                End If
                Return True

            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try

    End Function

#End Region

End Class
