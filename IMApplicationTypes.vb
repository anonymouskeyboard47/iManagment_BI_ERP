
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMApplicationTypes


#Region "PrivateVariables"

    Private lApplicationTypeID As Long
    Private strApplicationType As String
    Private bApplicationTypeStatus As Boolean

#End Region

#Region "Properties"

    Public Property ApplicationTypeID() As Long

        Get
            Return lApplicationTypeID
        End Get

        Set(ByVal Value As Long)
            lApplicationTypeID = Value
        End Set

    End Property

    Public Property ApplicationType() As String

        Get
            Return strApplicationType
        End Get

        Set(ByVal Value As String)
            strApplicationType = Value
        End Set

    End Property

    Public Property ApplicationTypeStatus() As Boolean

        Get
            Return bApplicationTypeStatus
        End Get

        Set(ByVal Value As Boolean)
            bApplicationTypeStatus = Value
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


            If Trim(strApplicationType) = "" _
                Then

                MsgBox("You must provide an appropriate Application Type." _
                                , MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if there is an existing series with this name
            If Find("SELECT * FROM ApplicationTypes WHERE ApplicationType = '" _
                & Trim(strApplicationType) & "'", False) = True Then

                If MsgBox("This Application Type's details already exist." & _
                    Chr(10) & "Do you want to update the  details?", _
                            MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                    MsgBoxResult.No Then

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If

                Update("UPDATE ApplicationTypes SET " & _
                    " ApplicationTypeStatus = " & bApplicationTypeStatus & _
                        " WHERE  ApplicationType = '" & Trim(strApplicationType) & "'")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If



            strInsertInto = "INSERT INTO ApplicationTypes (" & _
                    "ApplicationType," & _
                    "ApplicationTypeStatus" & _
                            ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & Trim(strApplicationType) & _
                        "," & bApplicationTypeStatus & _
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
                    "iManagement - Application Type's details Saved")

                End If
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save Referee's Customer' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Application Type's details Failed")
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

                        lApplicationTypeID = _
                                myDataRows("ApplicationTypeID")
                        strApplicationType = _
                                myDataRows("ApplicationType")
                        bApplicationTypeStatus = _
                                myDataRows("ApplicationTypeStatus")

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

            If lApplicationTypeID = 0 _
                Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - Application Type's Details Deleted")
                Else
                    MsgBox("'Referee's Customer delete' action failed", _
                        MsgBoxStyle.Exclamation, "Application Type's Details Deletion failed")
                End If
            Else

                MsgBox("Cannot Delete. Please select an existing Application and an Application Type's Details.", _
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
                    "iManagement -  Application Type's Details Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub

   
#End Region

End Class
