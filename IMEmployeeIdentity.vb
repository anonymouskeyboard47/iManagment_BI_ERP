Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMEmployeeIdentity

#Region "PrivateVariables"

    Private lCustomerNo As Long
    Private lIdentityKeyField As Long
    Private strIdentityType As String
    Private strIdentityNo As String
    Private strIdentitySerialNo As String
    Private strPhone4 As String
    Private strPhone5 As String
    Private strPostAddress2 As String
    Private strPostCode2 As String
    Private strEmailAddress2 As String
    Private strPhysicalAddress2 As String

#End Region

#Region "Properties"

    Public Property CustomerNo() As String

        'USED TO SET AND RETRIEVE THE SALARYTYPE ID (STRING)
        Get
            Return lCustomerNo
        End Get

        Set(ByVal Value As String)
            lCustomerNo = Value
        End Set

    End Property

    Public Property IdentityKeyField() As Long

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return lIdentityKeyField
        End Get

        Set(ByVal Value As Long)
            lIdentityKeyField = Value
        End Set

    End Property

    Public Property IdentityType() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strIdentityType
        End Get

        Set(ByVal Value As String)
            strIdentityType = Value
        End Set

    End Property

    Public Property IdentityNo() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strIdentityNo
        End Get

        Set(ByVal Value As String)
            strIdentityNo = Value
        End Set

    End Property

    Public Property IdentitySerialNo() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strIdentitySerialNo
        End Get

        Set(ByVal Value As String)
            strIdentitySerialNo = Value
        End Set

    End Property

    Public Property Phone4() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhone4
        End Get

        Set(ByVal Value As String)
            strPhone4 = Value
        End Set

    End Property

    Public Property Phone5() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhone5
        End Get

        Set(ByVal Value As String)
            strPhone5 = Value
        End Set

    End Property

    Public Property PhysicalAddress2() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPhysicalAddress2
        End Get

        Set(ByVal Value As String)
            strPhysicalAddress2 = Value
        End Set

    End Property

    Public Property PostalAddress2() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostAddress2
        End Get

        Set(ByVal Value As String)
            strPostAddress2 = Value
        End Set

    End Property

    Public Property PostalCode() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strPostCode2
        End Get

        Set(ByVal Value As String)
            strPostCode2 = Value
        End Set

    End Property

    Public Property EmailAddress2() As String

        'USED TO SET AND RETRIEVE THE SalaryTypeName (STRING)
        Get
            Return strEmailAddress2
        End Get

        Set(ByVal Value As String)
            strEmailAddress2 = Value
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

        lCustomerNo = ""
        lIdentityKeyField = 0
        strIdentityType = ""
        strIdentityNo = ""
        strIdentitySerialNo = ""
        strPhone4 = ""
        strPhone5 = ""
        strPostAddress2 = ""
        strPostCode2 = ""
        strEmailAddress2 = ""
        strPhysicalAddress2 = ""


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

        If Trim(lCustomerNo) <> "" And _
            Trim(strIdentityType) <> "" And _
                strIdentityNo <> "" _
                            Then

            strInsertInto = "INSERT INTO CustomerIdentity (" & _
                "CustomerNo," & _
                "IdentityType," & _
                "IdentityNo," & _
                "IdentitySerialNo," & _
                "Phone3," & _
                "Phone4," & _
                "PostAddress2," & _
                "PostCode2," & _
                "emailaddress2," & _
                "PhysicalAddress2," & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "'" & lCustomerNo & _
                    "'" & strIdentityType & _
                    "'" & strIdentityNo & _
                    "'" & strIdentitySerialNo & _
                    "'" & strPhone4 & _
                    "'" & strPhone5 & _
                    "'" & strPostAddress2 & _
                    "'" & strPostCode2 & _
                    "'" & strEmailAddress2 & _
                    "'" & strPhysicalAddress2 & _
                    ")"

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                "iManagement - Customer's Identity Saved")

            Else

                MsgBox("'Save Customer' action failed." & _
                    " Make sure all mandatory details are entered", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement - Customer's Identity Addition Failed")

            End If

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
                    Exit Function

                End If

                For Each myDataRows In myDataTables.Rows

                    lCustomerNo = myDataRows("CustomerNo").ToString()
                    lIdentityKeyField = myDataRows("CustomerNo").ToString()
                    strIdentityType = myDataRows("CustomerNo").ToString()
                    strIdentityNo = myDataRows("CustomerNo").ToString()
                    strIdentitySerialNo = myDataRows("CustomerNo").ToString()
                    strPhone4 = myDataRows("CustomerNo").ToString()
                    strPhone5 = myDataRows("CustomerNo").ToString()
                    strPostAddress2 = myDataRows("CustomerNo").ToString()
                    strPostCode2 = myDataRows("CustomerNo").ToString()
                    strEmailAddress2 = myDataRows("CustomerNo").ToString()
                    strPhysicalAddress2 = myDataRows("CustomerNo").ToString()

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

        If Trim(lCustomerNo) <> "" Or _
                Trim(strIdentityNo) <> "" Or _
                    Trim(strIdentityType) <> "" _
                            Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Customer's Identity Details Deleted")
            Else
                MsgBox("'Delete Customer's Identity action failed", _
                    MsgBoxStyle.Exclamation, " Customer Deletion failed")
            End If
        Else
            MsgBox("Cannot Delete. Please select an existing Customer's Identity", _
                    MsgBoxStyle.Exclamation, "iManagement -Missing Information")

        End If

    End Sub

    Public Sub Update(ByVal strUpQuery As String)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If Trim(lCustomerNo) <> "" Or _
                 Trim(strIdentityType) <> "" Or _
                    Trim(strIdentityNo) <> "" _
                        Then

            objLogin.connectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                    "iManagement -  Customer's Identity Details Updated")
            End If

        End If

    End Sub

   

#End Region


End Class
