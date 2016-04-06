Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCustomerPriorities

#Region "PrivateCustPriorityVariables"
    Private lCustPriorityID As String
    Private strPriorityDescription As String
#End Region

#Region "Properties"

    Public Property CustPriorityID() As String

        'USED TO SET AND RETRIEVE THE CustPriorityID (Long)
        Get
            Return lCustPriorityID
        End Get

        Set(ByVal Value As String)
            lCustPriorityID = Value
        End Set

    End Property

    Public Property PriorityDescription() As String

        'USED TO SET AND RETRIEVE THE PriorityDescription (STRING)
        Get
            Return strPriorityDescription
        End Get

        Set(ByVal Value As String)
            strPriorityDescription = Value
        End Set

    End Property


#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Sub NewRecord()
        lCustPriorityID = 0
        strPriorityDescription = ""
    End Sub

#End Region

#Region "DatabaseProcedures"

    Public Sub Save()
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String

        If Trim(strPriorityDescription) <> "" Then

            strInsertInto = "INSERT INTO CustomerPriority (" & _
                "CustomerPriorityID," & _
                "PriorityDescription) VALUES "

            strSaveQuery = strInsertInto & _
                        "('" & lCustPriorityID & _
                        "', '" & strPriorityDescription & _
                        "')"

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Customer Priority Saved")

            Else

                MsgBox("'Customer Priority' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Save Customer Priority Failed")

            End If

        Else

            MsgBox("'Customer Priority' action failed." & _
                " Make sure all mandatory details are entered", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Save Customer Priority Failed")

        End If

    End Sub

    Public Function Find(ByVal strQuery As String) As Boolean

        Dim datRetData As DataSet = New DataSet
        Dim bQuerySuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim objLogin As IMLogin = New IMLogin

        objLogin.connectString = strAccessConnString
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
                    Return False
                    datRetData = Nothing
                    objLogin = Nothing
                    Exit Function

                End If


                For Each myDataRows In myDataTables.Rows

                    lCustPriorityID = myDataRows("CustomerPriorityID")

                    strPriorityDescription = myDataRows _
                        ("PriorityDescription").ToString()

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Sub Delete(ByVal strDelQuery As String)

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strDeleteQuery = strDelQuery

        If Trim(lCustPriorityID) <> "" Or _
                          Trim(strPriorityDescription) <> "" Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Lookup Details Deleted")
            Else
                MsgBox("'Customer Priority delete' action failed", _
                    MsgBoxStyle.Exclamation, " Deletion failed")
            End If

        Else

            MsgBox("Cannot Delete. Please select an existing Priority Type", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")
        End If

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If lCustPriorityID <> "" Or _
                       Trim(strPriorityDescription) <> "" Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Lookup Details Updated")
            End If

        End If

    End Sub

   


#End Region

End Class
