Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMBusinessProcessMaster

#Region "PrivateVariables"

    Private lBusinessProcessID As Long
    Private strBusinessProcessName As String
    Private strBusinessProcessDescription As String
    Private bBusinessProcessStatus As Boolean
    Private dtDateRegistered As Date
    Private lOrganizationID As Long

#End Region

#Region "Properties"

    Public Property DateRegistered() As Date

        Get
            Return dtDateRegistered
        End Get

        Set(ByVal Value As Date)
            dtDateRegistered = Value
        End Set

    End Property

    Public Property OrganizationID() As Long

        Get
            Return lOrganizationID
        End Get

        Set(ByVal Value As Long)
            lOrganizationID = Value
        End Set

    End Property


    Public Property BusinessProcessID() As Long

        Get
            Return lBusinessProcessID
        End Get

        Set(ByVal Value As Long)
            lBusinessProcessID = Value
        End Set

    End Property

    Public Property BusinessProcessName() As String

        Get
            Return strBusinessProcessName
        End Get

        Set(ByVal Value As String)
            strBusinessProcessName = Value
        End Set

    End Property

    Public Property BusinessProcessDescription() As String

        Get
            Return strBusinessProcessDescription
        End Get

        Set(ByVal Value As String)
            strBusinessProcessDescription = Value
        End Set

    End Property

    Public Property BusinessProcessStatus() As Boolean

        Get
            Return bBusinessProcessStatus
        End Get

        Set(ByVal Value As Boolean)
            bBusinessProcessStatus = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function ReturnWorkFlowMasterName _
       (ByVal strIMMappingName As String) As String

        Try

            If strIMMappingName = "Accounting" Then
                Return "Accounting"

            End If


            If strIMMappingName = "Money to pay" Then
                Return "Accounts payable"

            End If


            If strIMMappingName = "Money to receive" Then
                Return "Accounts reeivable"

            End If


            If strIMMappingName = "Bank or money deposit centre" Then
                Return "Bank"

            End If


            '--------------------
            If strIMMappingName = "Meeting	" Then
                Return "Board Of Directors"

            End If


            If strIMMappingName = "Many Documents" Then
                Return "Copy centre"

            End If


            If strIMMappingName = "Customer Service" Then
                Return "Customer Service"

            End If


            If strIMMappingName = "Shipping" Then
                Return "Distribution"

            End If


            If strIMMappingName = "Certificate" Then
                Return "Finance"

            End If


            '===================
            If strIMMappingName = "Information system" Then
                Return "Information system"

            End If


            If strIMMappingName = "International Division" Then
                Return "International Division"

            End If


            If strIMMappingName = "International Marketing" Then
                Return "International Marketing"

            End If


            If strIMMappingName = "Export and International Sales" Then
                Return "Export and International Sales"

            End If


            If strIMMappingName = "Inventory" Then
                Return "Inventory"

            End If


            '\\\\\\\\\\\\\\\\\\\\\
            If strIMMappingName = "Legal Department" Then
                Return "Legal Department"

            End If


            If strIMMappingName = "Mailroom 1" Then
                Return "Mailroom 1"

            End If


            If strIMMappingName = "Mailroom 2" Then
                Return "Mailroom 2"

            End If


            '\\\\\\\\\\\\\\\\\\\\\
            If strIMMappingName = "Hierachical Management" Then
                Return "Management"

            End If


            If strIMMappingName = "Manufacturing" Then
                Return "Manufacturing"

            End If


            If strIMMappingName = "Marketing" Then
                Return "Marketing"

            End If


            '==========
            If strIMMappingName = "Fleet" Then
                Return "Motorpool"

            End If


            If strIMMappingName = "Manufacturing" Then
                Return "Manufacturing"

            End If


            If strIMMappingName = "Marketing" Then
                Return "Marketing"

            End If


            '==========
            If strIMMappingName = "Packaging" Then
                Return "Packaging"

            End If


            If strIMMappingName = "Payroll" Then
                Return "Payroll"

            End If


            If strIMMappingName = "Person 1" Then
                Return "Person 1"

            End If


            '\\\\\\
            If strIMMappingName = "Person 2" Then
                Return "Person 2"

            End If


            If strIMMappingName = "Personnel and Staff" Then
                Return "Personnel and Staff"

            End If

            If strIMMappingName = "Publications" Then
                Return "Publications"

            End If


            '\\\\\\
            If strIMMappingName = "Purchasing" Then
                Return "Purchasing"

            End If


            If strIMMappingName = "Quality Control Approval" Or _
                strIMMappingName = "Approval" Or _
                    strIMMappingName = "Quality Control" Then

                Return "Quality Assurance"

            End If


            '\\\\\\
            If strIMMappingName = "Receiving Item" Or _
                strIMMappingName = "Receive Document from Archive" Then

                Return "Receiving"

            End If


            If strIMMappingName = "Reception" Then

                Return "Reception"

            End If


            If strIMMappingName = "Research and Development" Then
                Return "Research and Development"

            End If


            '\\\\\\
            If strIMMappingName = "Sales and PR" Then

                Return "Sales/PR"

            End If


            If strIMMappingName = "Outgoing Shipping" Then

                Return "Shipping"

            End If


            If strIMMappingName = "Incoming Shipping" Then
                Return "Suppliers"

            End If


            '\\\\\\
            If strIMMappingName = "Telecommunication" Then

                Return "Telecom"

            End If


            If strIMMappingName = "Treasurer" Then

                Return "Treasurer"

            End If


            If strIMMappingName = "Warehouse" Then
                Return "Warehouse"

            End If


            '\\\\\\
            If strIMMappingName = "Connector" Then

                Return "Dynamic Connector"

            End If


            If strIMMappingName = "Line-Curve connector" Then

                Return "Line-Curve connector"

            End If


            If strIMMappingName = "On-page reference" Then
                Return "On-page reference"

            End If


            If strIMMappingName = "Off-page reference" Then

                Return "Off-page reference"

            End If


            If strIMMappingName = "Start Terminator" Or _
                strIMMappingName = "End Terminator" Then

                Return "Terminator"

            End If



        Catch ex As Exception

        End Try

    End Function


    Public Function ReturnBusinessProcessesWorkFlows _
        (ByVal strValBusinessProcess As String, ByVal bUseWorkFlows As Boolean, _
            ByVal bAllBusinessProcesses As Boolean, _
                ByVal bIncludeBusinessProcessesWithNoWorkFlows As Boolean) _
                    As Object

        Try

            Dim arBusinessProcessFlow(,,) As Object

            Dim strQueryToUse As String
            Dim objLogin As IMLogin

            If bIncludeBusinessProcessesWithNoWorkFlows = True Then
                If bAllBusinessProcesses = True Then

                    strQueryToUse = "SELECT " & _
                    " BusinessProcessMaster.BusinessProcessName, " & _
                    " BusinessProcessWorkFlows.WorkFlowName, " & _
                    " BusinessProcessWorkFlows.WorkFlowType, " & _
                    " BusinessProcessWorkFlows.WorkFlowStatus, " & _
                    " BusinessProcessWorkFlows.WorkFlowPosition, " & _
                    " BusinessProcessWorkFlows.WorkFlowDescription, " & _
                    " BusinessProcessMaster.BusinessProcessDescription " & _
                    " FROM BusinessProcessMaster LEFT JOIN BusinessProcessWorkFlows " & _
                    " ON BusinessProcessMaster.BusinessProcessID = " & _
                    " BusinessProcessWorkFlows.BusinessProcessID " & _
                    " ORDER BY BusinessProcessWorkFlows.WorkFlowPosition ASC "

                Else

                    strQueryToUse = "SELECT " & _
                    " BusinessProcessMaster.BusinessProcessName, " & _
                    " BusinessProcessWorkFlows.WorkFlowName, " & _
                    " BusinessProcessWorkFlows.WorkFlowType, " & _
                    " BusinessProcessWorkFlows.WorkFlowStatus, " & _
                    " BusinessProcessWorkFlows.WorkFlowPosition, " & _
                    " BusinessProcessWorkFlows.WorkFlowDescription, " & _
                    " BusinessProcessMaster.BusinessProcessDescription " & _
                    " FROM BusinessProcessMaster LEFT JOIN BusinessProcessWorkFlows " & _
                    " ON BusinessProcessMaster.BusinessProcessID = " & _
                    " BusinessProcessWorkFlows.BusinessProcessID " & _
                    " WHERE BusinessProcessMaster.BusinessProcessName = '" & _
                    strValBusinessProcess & "'" & _
                    " ORDER BY BusinessProcessWorkFlows.WorkFlowPosition ASC"

                End If
            End If


            If bIncludeBusinessProcessesWithNoWorkFlows = False Then
                If bAllBusinessProcesses = True Then

                    strQueryToUse = "SELECT " & _
                    " BusinessProcessMaster.BusinessProcessName, " & _
                    " BusinessProcessWorkFlows.WorkFlowName, " & _
                    " BusinessProcessWorkFlows.WorkFlowType, " & _
                    " BusinessProcessWorkFlows.WorkFlowStatus, " & _
                    " BusinessProcessWorkFlows.WorkFlowPosition, " & _
                    " BusinessProcessWorkFlows.WorkFlowDescription, " & _
                    " BusinessProcessMaster.BusinessProcessDescription " & _
                    " FROM BusinessProcessMaster LEFT JOIN BusinessProcessWorkFlows " & _
                    " ON BusinessProcessMaster.BusinessProcessID = " & _
                    " BusinessProcessWorkFlows.BusinessProcessID " & _
                    " WHERE  (BusinessProcessWorkFlows.WorkFlowName <> ''" & _
                    " OR  BusinessProcessWorkFlows.WorkFlowName IS NOT NULL) " & _
                    " ORDER BY BusinessProcessWorkFlows.WorkFlowPosition ASC "

                Else

                    strQueryToUse = "SELECT " & _
                    " BusinessProcessMaster.BusinessProcessName, " & _
                    " BusinessProcessWorkFlows.WorkFlowName, " & _
                    " BusinessProcessWorkFlows.WorkFlowType, " & _
                    " BusinessProcessWorkFlows.WorkFlowStatus, " & _
                    " BusinessProcessWorkFlows.WorkFlowPosition, " & _
                    " BusinessProcessWorkFlows.WorkFlowDescription, " & _
                    " BusinessProcessMaster.BusinessProcessDescription " & _
                    " FROM BusinessProcessMaster LEFT JOIN BusinessProcessWorkFlows " & _
                    " ON BusinessProcessMaster.BusinessProcessID = " & _
                    " BusinessProcessWorkFlows.BusinessProcessID " & _
                    " WHERE BusinessProcessMaster.BusinessProcessName = '" & _
                    strValBusinessProcess & _
                    "' AND  (BusinessProcessWorkFlows.WorkFlowName <> ''" & _
                    " OR  BusinessProcessWorkFlows.WorkFlowName IS NOT NULL) " & _
                    " ORDER BY BusinessProcessWorkFlows.WorkFlowPosition ASC"

                End If
            End If


            objLogin = New IMLogin

            With objLogin

                arBusinessProcessFlow = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "", 3)

            End With

            objLogin = Nothing

            Return arBusinessProcessFlow

        Catch ex As Exception

        End Try

    End Function

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
            Dim objOverallSetup As IMOverallSetup


            If Trim(strOrganizationName) = "" Then

                returnerror += "Please open an existing company."
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Trim(strBusinessProcessName) = "" Or _
                    Trim(strBusinessProcessDescription) = "" Or _
                        lOrganizationID = 0 _
                            Then

                ReturnError += "You must provide an appropriate " & _
                "Business Process Name, an associated" & _
                Chr(10) & " Business Process Description, and " & _
                "Open an existing company."

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If




            'Check if there is an existing series with this name
            If Find("SELECT * FROM BusinessProcessMaster " & _
            "WHERE BusinessProcessName = '" & strBusinessProcessName & _
                    "' AND OrganizationID = " & lOrganizationID, _
                        False) = True Then

                ''confirm Update
                'If MsgBox("This Business Process Name already exists in" & _
                '" this Organization." & Chr(10) & "Do you want to update" & _
                '" the  details?", _
                '            MsgBoxStyle.YesNo, "iManagement - Record Exists. Update") = _
                '                    MsgBoxResult.No Then

                '    objLogin = Nothing
                '    datSaved = Nothing

                '    Exit Function

                'End If


                Update("UPDATE BusinessProcessMaster SET " & _
                    "BusinessProcessName = '" & Trim(strBusinessProcessName) & _
                    "', BusinessProcessDescription = '" & Trim(strBusinessProcessDescription) & _
                    "' , BusinessProcessStatus = " & bBusinessProcessStatus & _
                        " WHERE BusinessProcessName = '" _
                                & strBusinessProcessName & "' OR BusinessProcessID = " & _
                                        lBusinessProcessID)

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            ''Confirm Addition
            'If returnerror +="Do you want to add this new Business " & _
            '"Process Name?" _
            '    , MsgBoxStyle.YesNo, _
            '        "iManagement - Add the Business Process Name?") _
            '            = MsgBoxResult.No Then

            '    objLogin = Nothing
            '    datSaved = Nothing

            '    Exit Function
            'End If


            strInsertInto = "INSERT INTO BusinessProcessMaster (" & _
                    "BusinessProcessName," & _
                    "BusinessProcessDescription," & _
                    "BusinessProcessStatus," & _
                    "OrganizationID" & _
                            ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & Trim(strBusinessProcessName) & _
                        "','" & Trim(strBusinessProcessDescription) & _
                        "'," & bBusinessProcessStatus & _
                        "," & lOrganizationID & _
                            ")"


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    ReturnError += "New Business Process Record Saved " & _
                    "Successfully"

                End If
            Else

                If DisplayFailureMessages = True Then
                    returnerror += "'Save Business Process details' action failed." & _
                        " Make sure all mandatory details are entered."
                End If
            End If

            objLogin = Nothing
            datSaved = Nothing

        Catch ex As Exception

            If DisplayErrorMessages = True Then
                ReturnError += "iManagement - Database or system error " & _
                    ex.Message
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

        bQuerySuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strQuery, datRetData)

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


                For Each myDataRows In myDataTables.Rows

                    If bReturnValues = True Then

                        If IsDBNull(myDataRows("BusinessProcessID")) = False Then
                            lBusinessProcessID = _
                                    myDataRows("BusinessProcessID")
                        End If

                        If IsDBNull(myDataRows("BusinessProcessName")) = False Then
                            strBusinessProcessName = _
                                    myDataRows("BusinessProcessName").ToString
                        End If

                        If IsDBNull(myDataRows("BusinessProcessDescription")) = False Then
                            strBusinessProcessDescription = _
                                    myDataRows("BusinessProcessDescription").ToString
                        End If

                        If IsDBNull(myDataRows("BusinessProcessStatus")) = False Then
                            bBusinessProcessStatus = _
                                    myDataRows("BusinessProcessStatus")
                        End If

                        If IsDBNull(myDataRows("DateRegistered")) = False Then
                            dtDateRegistered = _
                                    myDataRows("DateRegistered")
                        End If

                        If IsDBNull(myDataRows("OrganizationID")) = False Then
                            lOrganizationID = _
                                    myDataRows("OrganizationID")
                        End If


                    End If
                Next

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
    Public Sub Delete()

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try

            If lBusinessProcessID = 0 _
                Then

                returnerror += "Cannot Delete. Please select an existing Business Process Details."

                objLogin = Nothing
                datDelete = Nothing

                Exit Sub

            End If

            'If MsgBox("Please be careful. Deleting a business process will also" & _
            'Chr(10) & "delete all its associated work flows and the system" & _
            'Chr(10) & "may fail to work properly. Are you are sure about the deletion.", _
            'MsgBoxStyle.YesNo, _
            '    "iManagement - Delete business process and its associated work flows?") = MsgBoxResult.No Then

            '    datDelete = Nothing
            '    objLogin = Nothing
            '    Exit Sub

            'End If


            strDeleteQuery = "DELETE * FROM BusinessProcessMaster " & _
            "WHERE BusinessProcessID = " & lBusinessProcessID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strDeleteQuery, datDelete)


            'Delete from business process work flows
            strDeleteQuery = "DELETE * FROM BusinessProcessWorkFlows " & _
            "WHERE BusinessProcessID = " & lBusinessProcessID


            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)


            objLogin.CloseDb()

            If bDelSuccess = True Then
                returnerror += "Business Process Record Deleted " & _
                "Successfully."
            Else
                returnerror += "'Delete Business Process' action failed."
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

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                strUpdateQuery, datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                ReturnError += "Business Process record updated " & _
                "successfully."
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub

#End Region



End Class
