Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections


Public Class IMMeetingMaster

#Region "PrivateVariables"

    Private strMeetingSerialNumber As String
    Private lMeetingID As Long
    Private strChairpersonNames As String
    Private strSecretariesName As String
    Private dtDateOfMeeting As Date
    Private dtStartTime As Date
    Private dtEndTime As Date
    Private strVenue As String
    Private strVenueDescription As String
    Private strMeetingType As String
    Private dtDateRegistered As String
    Private bMeetingStatus As Boolean
    Private bIsMeetingConfirmed As Boolean

#End Region

#Region "Properties"

    Public Property MeetingStatus() As Boolean

        Get
            Return bMeetingStatus
        End Get

        Set(ByVal Value As Boolean)
            bMeetingStatus = Value
        End Set

    End Property

    Public Property IsMeetingConfirmed() As Boolean

        Get
            Return bIsMeetingConfirmed
        End Get

        Set(ByVal Value As Boolean)
            bIsMeetingConfirmed = Value
        End Set

    End Property

    Public Property MeetingSerialNumber() As String

        Get
            Return strMeetingSerialNumber
        End Get

        Set(ByVal Value As String)
            strMeetingSerialNumber = Value
        End Set

    End Property

    Public Property MeetingID() As Long

        Get
            Return lMeetingID
        End Get

        Set(ByVal Value As Long)
            lMeetingID = Value
        End Set

    End Property

    Public Property ChairpersonNames() As String

        Get
            Return strChairpersonNames
        End Get

        Set(ByVal Value As String)
            strChairpersonNames = Value
        End Set

    End Property

    Public Property SecretariesName() As String

        Get
            Return strSecretariesName
        End Get

        Set(ByVal Value As String)
            strSecretariesName = Value
        End Set

    End Property

    Public Property DateOfMeeting() As Date

        Get
            Return dtDateOfMeeting
        End Get

        Set(ByVal Value As Date)
            dtDateOfMeeting = Value
        End Set

    End Property

    Public Property StartTime() As Date

        Get
            Return dtStartTime
        End Get

        Set(ByVal Value As Date)
            dtStartTime = Value
        End Set

    End Property

    Public Property EndTime() As Date

        Get
            Return dtEndTime
        End Get

        Set(ByVal Value As Date)
            dtEndTime = Value
        End Set

    End Property

    Public Property Venue() As String

        Get
            Return strVenue
        End Get

        Set(ByVal Value As String)
            strVenue = Value
        End Set

    End Property

    Public Property VenueDescription() As String

        Get
            Return strVenueDescription
        End Get

        Set(ByVal Value As String)
            strVenueDescription = Value
        End Set

    End Property

    Public Property MeetingType() As String

        Get
            Return strMeetingType
        End Get

        Set(ByVal Value As String)
            strMeetingType = Value
        End Set

    End Property

#End Region

#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function CalculateNextMeetingSerialNo() As String
        Try

            Dim MaxValue As Long
            Dim MyMaxValue() As String
            Dim strItem As String
            Dim strProposedApplNo As String

            MyMaxValue = FillControl(strOrgAccessConnString, _
                           "SELECT COUNT(*) AS TotalRecords FROM" & _
                               " MeetingMaster WHERE DateRegistered = Now()", "", "")

            If Not MyMaxValue Is Nothing Then
                For Each strItem In MyMaxValue
                    If Not strItem Is Nothing Then

                        MaxValue = CLng(Val(strItem))


                    End If
                Next
            End If

            MaxValue = MaxValue + 1

            strProposedApplNo = "Appl" & Now.Day.ToString _
                & Now.Month.ToString & _
                    Now.Year.ToString & _
                            MaxValue.ToString

            Return strProposedApplNo

        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, _
                "iManagement - System Error")
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

            If Trim(strOrganizationName) = "" Then

                MsgBox("Please open an existing company.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Select an existing company")
                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Trim(strChairpersonNames) = 0 _
                Or Trim(strSecretariesName) = 0 _
                    Or Trim(strVenue) = 0 _
                        Or Trim(strMeetingType) = 0 _
                        Then

                MsgBox("You must provide appropriate Chairperson's Names, SecretariesName, Meeting Venue," & _
                Chr(10) & "and the Type Of Meeting being conducted.", _
                                MsgBoxStyle.Critical, _
                                    "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            'Check if there is an existing series with this name
            If Find("SELECT * FROM MeetingMaster WHERE MeetingID = " & _
                lMeetingID & " OR MeetingSerialNumber = '" & _
                    Trim(strMeetingSerialNumber) & "'", False) = True Then

                If MsgBox("The Meeting's Details already exists." & _
                Chr(10) & "Do you want to update the details?", _
                        MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                MsgBoxResult.Yes Then

                    Update("UPDATE MeetingMaster SET " & _
                        "ChairpersonNames = '" & Trim(strChairpersonNames) & _
                                "' AND SecretariesName = '" & Trim(strSecretariesName) & _
                                 "' AND DateOfMeeting = '" & dtDateOfMeeting & _
                                  "' AND StartTime = '" & dtStartTime & _
                                   "' AND EndTime = '" & dtEndTime & _
                                    "' AND Venue = '" & Trim(strVenue) & _
                                     "' AND VenueDescription = '" & Trim(strVenueDescription) & _
                                     "' AND MeetingType = '" & Trim(strMeetingType) & _
                                     "' AND MeetingStatus = " & Trim(bMeetingStatus) & _
                                     " AND IsMeetingConfirmed = " & bIsMeetingConfirmed & _
                                    " WHERE  MeetingID = " _
                                        & lMeetingID & " OR MeetingSerialNumber  = '" & _
                                            strMeetingSerialNumber & "'")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO MeetingMaster (" & _
                "MeetingSerialNumber," & _
                "ChairpersonNames," & _
                "SecretariesName," & _
                "DateOfMeeting," & _
                "StartTime," & _
                "EndTime," & _
                "Venue," & _
                "VenueDescription," & _
                "MeetingType," & _
                "MeetingStatus," & _
                "IsMeetingConfirmed," & _
                "DateRegistered" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & Trim(strMeetingSerialNumber) & _
                    "','" & Trim(strChairpersonNames) & _
                    "','" & Trim(strSecretariesName) & _
                    "','" & dtDateOfMeeting & _
                    "','" & dtStartTime & _
                    "','" & dtEndTime & _
                    "','" & Trim(strVenue) & _
                    "','" & Trim(strVenueDescription) & _
                    "','" & Trim(strMeetingType) & _
                    "'," & bMeetingStatus & _
                    "," & bIsMeetingConfirmed & _
                    ",'" & Now() & _
                    "')"


            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
                    "iManagement - New Meeting Details Saved")

                End If
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Save New Meeting Details' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Meeting Details Failed")
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

                        lMeetingID = _
                                myDataRows("lMeetingID")
                        strMeetingSerialNumber = _
                                myDataRows("strMeetingSerialNumber").ToString
                        strChairpersonNames = _
                                myDataRows("strChairpersonNames").ToString
                        strSecretariesName = _
                                myDataRows("strSecretariesName").ToString
                        dtDateOfMeeting = _
                                myDataRows("dtDateOfMeeting")
                        dtStartTime = _
                                myDataRows("dtStartTime")
                        dtEndTime = _
                                myDataRows("dtEndTime")
                        strVenue = _
                                myDataRows("strVenue").ToString
                        strVenueDescription = _
                                myDataRows("strVenueDescription").ToString
                        strMeetingType = _
                                myDataRows("MeetingType").ToString
                        bMeetingStatus = _
                                myDataRows("MeetingStatus")
                        bIsMeetingConfirmed = _
                                myDataRows("IsMeetingConfirmed")

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

            If lMeetingID = 0 Or Trim(strMeetingSerialNumber) = "" _
                Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
                datDelete)

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                        "iManagement - Meeting Details Deleted")
                Else
                    MsgBox("'Meeting Details delete' action failed", _
                            MsgBoxStyle.Exclamation, "Meeting Details Deletion failed")
                End If
            Else

                MsgBox("Cannot Delete. Please select an existing Meeting.", _
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
                    "iManagement - Meeting Details Details Updated")
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Sub

    Public Function FillControl(ByVal strFillConnString As String, _
                ByVal strTSQL As String, ByVal strValueField As String, _
                    ByVal strTextField As String) As String()

        Dim datFillData As DataSet
        Dim bReturnedSuccess As Boolean
        Dim myDataTables As DataTable
        Dim myDataColumns As DataColumn
        Dim myDataRows As DataRow
        Dim strTextFieldData() As String
        Dim i As Integer
        Dim objLogin As IMLogin = New IMLogin

        Try

            datFillData = New DataSet

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            'The db is okay now get the recordset
            bReturnedSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                strTSQL, datFillData)

            objLogin.CloseDb()

            If datFillData Is Nothing Then
                Exit Function
            End If

            For Each myDataTables In datFillData.Tables

                'Check if there is any data. If not exit
                If myDataTables.Rows.Count = 0 Then
                    'Return an empty array
                    ReDim strTextFieldData(1)
                    strTextFieldData(0) = ""
                    Return strTextFieldData

                    Exit Function
                Else
                    'Resize the array
                    ReDim strTextFieldData(myDataTables.Rows.Count)

                End If

                i = 0
                For Each myDataRows In myDataTables.Rows
                    strTextFieldData(i) = myDataRows(0).ToString()
                    i = i + 1
                Next

            Next


            objLogin = Nothing
            datFillData = Nothing

            Return strTextFieldData
            datFillData.Dispose()

        Catch ex As Exception

        End Try

    End Function

#End Region

End Class
