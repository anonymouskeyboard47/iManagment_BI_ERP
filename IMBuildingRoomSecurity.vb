
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMBuildingRoomSecurity

#Region "PrivateVariables"

    Private lBuildingID As Long
    Private lRoomID As Long
    Private lSecurityID As Long

#End Region

#Region "Properties"

    Public Property BuildingID() As Long

        Get
            Return lBuildingID
        End Get

        Set(ByVal Value As Long)
            lBuildingID = Value
        End Set

    End Property

    Public Property RoomID() As Long

        Get
            Return lRoomID
        End Get

        Set(ByVal Value As Long)
            lRoomID = Value
        End Set

    End Property

    Public Property SecurityID() As Long

        Get
            Return lSecurityID
        End Get

        Set(ByVal Value As Long)
            lSecurityID = Value
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


            If lBuildingID = 0 _
                Or lRoomID = 0 _
                    Or lSecurityID = 0 _
                        Then

                MsgBox("You must provide appropriate Building, Room, and Security Scheme Details.", _
                                MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            'Check if there is an existing series with this name
            If Find("SELECT * FROM BuildingRoomSecurity WHERE BuildingID = " & _
                lBuildingID & " AND RoomID  = " & _
                    lRoomID, False) = True Then

                If MsgBox("The Building Room's Security Details already exists." & _
                Chr(10) & "Do you want to update the details?", _
                        MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                MsgBoxResult.Yes Then

                    Update("UPDATE BuildingRoomSecurity SET " & _
                        "lBuildingID = " & lBuildingID & _
                            " AND lRoomID = " & lRoomID & _
                                " AND lSecurityID = " & lSecurityID & _
                                    " WHERE  lBuildingID = " _
                                        & BuildingID & " AND RoomID  = " & _
                                            lRoomID)

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            strInsertInto = "INSERT INTO BuildingRoomSecurity (" & _
                "BuildingID," & _
                "RoomID," & _
                "SecurityID" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lBuildingID & _
                    "," & lRoomID & _
                    "," & lSecurityID & _
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
                    "iManagement - New Building Room's Security Details Saved")

                End If
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save New Building Room's Security Details' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Building Room's Security Details Failed")
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

                        lBuildingID = _
                                myDataRows("BuildingID")
                        lRoomID = _
                                myDataRows("RoomID")
                        lSecurityID = _
                                myDataRows("SecurityID")

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

            If lBuildingID = 0 Or lSecurityID = 0 Or lRoomID = 0 _
                Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - Building Room's Security Details Deleted")
                Else
                    MsgBox("'Building Room's Security Details delete' action failed", _
                            MsgBoxStyle.Exclamation, "Building Room's Security Details Deletion failed")
                End If
            Else

                MsgBox("Cannot Delete. Please select an existing Building Room's Security.", _
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
                    "iManagement - Building Room's Security Details Details Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub

  

#End Region

End Class
