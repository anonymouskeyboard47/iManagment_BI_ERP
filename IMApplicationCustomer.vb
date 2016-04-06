Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMApplicationCustomer



#Region "PrivateVariables"

    Private lApplicationID As Long
    Private lCustomerNo As Long

#End Region

#Region "Properties"

    Public Property ApplCustApplicationNumber() As Long

        Get
            Return lCustomerNo
        End Get

        Set(ByVal Value As Long)
            lCustomerNo = Value
        End Set

    End Property

    Public Property ApplCustCustomerNo() As Long

        Get
            Return lCustomerNo
        End Get

        Set(ByVal Value As Long)
            lCustomerNo = Value
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
    Public Function ApplCustSave(ByVal bDisplayErrorMessages As Boolean, _
            ByVal bDisplaySuccessMessages As Boolean, _
                ByVal bDisplayFailureMessages As Boolean, _
                    ByVal bDisplayConfirmMessages As Boolean) As Boolean

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
            lCustomerNo = 0 _
                Then

                MsgBox("You must provide an appropriate Application and an Existing Customer." _
                                , MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if there is an existing series with this name
            If ApplCustFind("SELECT * FROM ApplicationCustomer WHERE " & _
            "ApplicationID = " & lApplicationID & _
                    " AND CustomerNo  = " & lCustomerNo, False) = True Then

                If bDisplayConfirmMessages = True Then
                    MsgBox("This Application's Customer details already exist for this application." _
                                , MsgBoxStyle.YesNo, _
                                "iManagement - Record Exists")
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            strInsertInto = "INSERT INTO ApplicationCustomer (" & _
                    "CustomerNo," & _
                        "ApplicationID" & _
                            ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lCustomerNo & _
                        "," & lApplicationID & _
                            ")"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If bDisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - Application's Customer Saved")

                End If
            Else

                If bDisplayFailureMessages = True Then
                    MsgBox("'Save Application's Customer' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Application's Customer Failed")
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
    Public Function ApplCustFind(ByVal strQuery As String, _
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

                        lApplicationID = _
                                myDataRows("ApplicationID")
                        lCustomerNo = _
                                myDataRows("CustomerNo")

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
    Public Sub ApplCustDelete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try


            strDeleteQuery = strDelQuery

            If lApplicationID = 0 Or lCustomerNo = 0 _
                Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - Application's Customer Details Deleted")
                Else
                    MsgBox("'Application's Customer delete' action failed", _
                        MsgBoxStyle.Exclamation, "Application's Customer Deletion failed")
                End If
            Else

                MsgBox("Cannot Delete. Please select an existing Application and an Existing Customer.", _
                        MsgBoxStyle.Exclamation, "iManagement - invalid or incomplete information")

            End If

            objLogin = Nothing
            datDelete = Nothing

        Catch ex As Exception

        End Try

    End Sub

    Public Sub ApplCustUpdate(ByVal strUpQuery As String)

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
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Application Details Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub


#End Region

End Class
