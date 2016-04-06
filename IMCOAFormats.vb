Option Explicit On 
'Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCOAFormats

    Private lChartID As Long
    Private strChartOfAccountName As String
    Private bChartStatus As Boolean

#Region "Properties"

    Public Property ChartID() As Long

        Get
            Return lChartID
        End Get

        Set(ByVal Value As Long)
            lChartID = Value
        End Set

    End Property

    Public Property ChartOfAccountName() As String

        Get
            Return strChartOfAccountName
        End Get

        Set(ByVal Value As String)
            strChartOfAccountName = Value
        End Set

    End Property

    Public Property ChartStatus() As Boolean

        Get
            Return bChartStatus
        End Get

        Set(ByVal Value As Boolean)
            bChartStatus = Value
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

            If Trim(strOrganizationName) = "" Then
                MsgBox("Please open an existing company.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Select an existing company")
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            If Trim(strChartOfAccountName) = "" Then

                MsgBox("You must provide a the name of the Chart Of Account Format you want to  save." _
                , MsgBoxStyle.Critical, _
                "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if there is an existing format with this name
            If Find("SELECT * FROM AvailableCharts WHERE ChartOfAccountName = '" _
                & Trim(strChartOfAccountName) & "'", _
                    False) = True Then

                If MsgBox("The Account Format Name you have provided exists." & _
                        Chr(10) & " Do you want to enable it as the default format?", _
                            MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, _
                                "iManagement - Set this format as the default?") _
                                = MsgBoxResult.Yes Then

                    'Check if there is an existing format with this name
                    If Find("SELECT * FROM AvailableCharts WHERE ChartOfAccountName = '" _
                        & Trim(strChartOfAccountName) & "'", _
                            False) = True Then


                        If MsgBox("There is an existing enabled format for this company." & _
                    " Only one format can be enabled at one time per company." & _
                            Chr(10) & "Do you want to disable the currently" & _
                            " enabled format and set this new one to be the default" & _
                            " for this organization?", _
                                MsgBoxStyle.Exclamation, _
                                    "iManagement - Cannot alter status") = _
                                        MsgBoxResult.No Then

                            objLogin = Nothing
                            datSaved = Nothing

                            Exit Function

                        Else

                            Update("UPDATE AvailableCharts SET " & _
                                "ChartStatus = False " & _
                                    " WHERE ChartStatus = TRUE")
                        End If

                        Update("UPDATE AvailableCharts SET " & _
                                                 "ChartStatus = True " & _
                                                     " WHERE strChartOfAccountName = '" _
                                                         & strChartOfAccountName & "'")

                    End If
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO AvailableCharts (" & _
    "ChartOfAccountName," & _
    "ChartStatus" & _
    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & Trim(strChartOfAccountName) & _
                    "', FALSE " & _
                    ")"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Chart Of Account Format Saved Successfully.", _
                        MsgBoxStyle.Information, _
                            "iManagement - Record Saved")

                End If

                Return True

            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save New Chart Of Account Format' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Record Failed")
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

                        lChartID = _
                                myDataRows("ChartID")
                        strChartOfAccountName = _
                                Trim(myDataRows("ChartOfAccountName").ToString)
                        bChartStatus = _
                            myDataRows("ChartStatus")

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
    Public Function Delete() As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin


            If lChartID = 0 Then
                MsgBox("Cannot Delete. Please select an existing Chart Of Account Format.", _
                      MsgBoxStyle.Exclamation, _
                        "iManagement - invalid or incomplete information")

                objLogin = Nothing

                datDelete = Nothing
                Exit Function

            End If


            If Trim(strChartOfAccountName) = "Default" Then

                MsgBox("The 'Default' format cannot be deleted. Please consult with the Systems Administrator." _
                , MsgBoxStyle.Critical, _
                "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datDelete = Nothing

                Exit Function
            End If

            strDeleteQuery = "DELETE * From AvailableCharts WHERE " & _
            "ChartID = " & lChartID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Chart Of Account Format Deleted Successfully", _
                    MsgBoxStyle.Information, _
                        "iManagement - Record Deleted")
                Return False
            Else
                MsgBox("'Chart Of Account Format delete' action failed", _
                    MsgBoxStyle.Exclamation, " Deletion failed")
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
                MsgBox("Chart Of Account Format Details Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement - Record Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Function


#End Region

End Class
