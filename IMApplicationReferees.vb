Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMApplicationReferees


#Region "PrivateVariables"

    Private lApplicationID As Long
    Private lRefereesCustomerNo As Long

#End Region

#Region "Properties"

    Public Property ApplicationID() As Long

        Get
            Return lApplicationID
        End Get

        Set(ByVal Value As Long)
            lApplicationID = Value
        End Set

    End Property

    Public Property RefereesCustomerNo() As Long

        Get
            Return lRefereesCustomerNo
        End Get

        Set(ByVal Value As Long)
            lRefereesCustomerNo = Value
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


            If lApplicationID = 0 Or _
                    lRefereesCustomerNo = 0 _
                Then

                MsgBox("You must provide an appropriate Application and an Existing Customer who is the referee." _
                                , MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if there is an existing series with this name
            If Find("SELECT * FROM ApplicationReferees WHERE ApplicationID = '" _
                & lApplicationID & " RefereesCustomerNo  = " & lRefereesCustomerNo, False) = True Then

                If MsgBox("This Referee's Customer details already exist for this application." & _
                    Chr(10) & "Do you want to update the  details?", _
                            MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                    MsgBoxResult.No Then

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If

                Update("UPDATE ApplicationReferees SET " & _
                    "RefereesCustomerNo = " & lRefereesCustomerNo & _
                        " WHERE  ApplicationID = " & lApplicationID _
                            & " AND RefereesCustomerNo = " & lRefereesCustomerNo)

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If



            strInsertInto = "INSERT INTO ApplicationReferees (" & _
                    "RefereesCustomerNo," & _
                        "ApplicationID" & _
                            ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lRefereesCustomerNo & _
                        "," & lApplicationID & _
                            ")"


            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - Referee's Customer Saved")

                End If
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save Referee's Customer' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Referee's Customer Failed")
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

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.connectString = strOrgAccessConnString
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

                        lApplicationID = _
                                myDataRows("ApplicationID")
                        lRefereesCustomerNo = _
                                myDataRows("RefereesCustomerNo")

                    Next
                End If
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

    End Function

    'Delete data
    Public Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try


            strDeleteQuery = strDelQuery

            If lApplicationID = 0 Or lRefereesCustomerNo = 0 _
                Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - Referee's Customer Details Deleted")
                Else
                    MsgBox("'Referee's Customer delete' action failed", _
                        MsgBoxStyle.Exclamation, "Referee's Customer Deletion failed")
                End If
            Else

                MsgBox("Cannot Delete. Please select an existing Application and an Referee's Customer Details.", _
                        MsgBoxStyle.Exclamation, "iManagement - invalid or incomplete information")

            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Referee's Customer Details Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub


#End Region

End Class
