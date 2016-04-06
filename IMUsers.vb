Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMUsers
    Inherits IMSIDMaster


#Region "PrivateVariables"

    Private lUserID As Long
    Private strUserName As String
    Private strNTUserID As String
    Private strSurname As String
    Private strFirstandMiddleName As String
    Private lNTSID As Long
    Private lDigitalSignature As Long
    Private strPhoneNumber As String
    Private strMobileNumber As String
    Private strOrganization As String
    Private strIDTypes As String
    Private strIDNumber As String
    Private dtCreationDate As Date
    Private strOldUserPassword As String
    Private strNewUserPassword As String
    Private strConfirmUserPassword As String
    Private bUserStatus As Boolean
    Private bUsePasswords As Boolean

#End Region


#Region "Properties"

    Public Property UsePasswords() As Boolean

        Get
            Return bUsePasswords
        End Get

        Set(ByVal Value As Boolean)
            bUsePasswords = Value
        End Set

    End Property

    Public Property OldUserPassword() As String

        Get
            Return strOldUserPassword
        End Get

        Set(ByVal Value As String)
            strOldUserPassword = Value
        End Set

    End Property

    Public Property NewUserPassword() As String

        Get
            Return strNewUserPassword
        End Get

        Set(ByVal Value As String)
            strNewUserPassword = Value
        End Set

    End Property

    Public Property ConfirmUserPassword() As String

        Get
            Return strConfirmUserPassword
        End Get

        Set(ByVal Value As String)
            strConfirmUserPassword = Value
        End Set

    End Property

    Public Property UserID() As Long

        Get
            Return lUserID
        End Get

        Set(ByVal Value As Long)
            lUserID = Value
        End Set

    End Property

    Public Property UserName() As String

        Get
            Return Trim(strUserName)
        End Get

        Set(ByVal Value As String)
            strUserName = Value
        End Set

    End Property

    Public Property NTUserID() As String

        Get
            Return Trim(strNTUserID)
        End Get

        Set(ByVal Value As String)
            strNTUserID = Value
        End Set

    End Property

    Public Property Surname() As String

        Get
            Return Trim(strSurname)
        End Get

        Set(ByVal Value As String)
            strSurname = Value
        End Set

    End Property

    Public Property FirstandMiddleName() As String

        Get
            Return Trim(strFirstandMiddleName)
        End Get

        Set(ByVal Value As String)
            strFirstandMiddleName = Value
        End Set

    End Property

    Public Property NTSID() As Long

        Get
            Return lNTSID
        End Get

        Set(ByVal Value As Long)
            lNTSID = Value
        End Set

    End Property

    Public Property DigitalSignature() As Long

        Get
            Return lDigitalSignature
        End Get

        Set(ByVal Value As Long)
            lDigitalSignature = Value
        End Set

    End Property

    Public Property PhoneNumber() As String

        Get
            Return Trim(strPhoneNumber)
        End Get

        Set(ByVal Value As String)
            strPhoneNumber = Value
        End Set

    End Property

    Public Property MobileNumber() As String

        Get
            Return Trim(strMobileNumber)
        End Get

        Set(ByVal Value As String)
            strMobileNumber = Value
        End Set

    End Property

    Public Property Organization() As String

        Get
            Return Trim(strOrganization)
        End Get

        Set(ByVal Value As String)
            strOrganization = Value
        End Set

    End Property

    Public Property IDTypes() As String

        Get
            Return Trim(strIDTypes)
        End Get

        Set(ByVal Value As String)
            strIDTypes = Value
        End Set

    End Property

    Public Property IDNumber() As String

        Get
            Return Trim(strIDNumber)
        End Get

        Set(ByVal Value As String)
            strIDNumber = Value
        End Set

    End Property

    Public Property CreationDate() As Date

        Get
            Return dtCreationDate
        End Get

        Set(ByVal Value As Date)
            dtCreationDate = Value
        End Set

    End Property

    Public Property UserStatus() As Boolean

        Get
            Return bUserStatus
        End Get

        Set(ByVal Value As Boolean)
            bUserStatus = Value
        End Set

    End Property


#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "GeneralProcedures"

    'UserID From SID
    Public Function ReturnUserIDFromSID _
        (ByVal strValUserSID As String) As Long

        Try

            Dim lUserId As Long
            Dim arFillControl() As String
            Dim strItem As String

            'Get the employer ID
            Dim objLogin As IMLogin = New IMLogin

            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT UserID FROM Users WHERE SystemSID = '" & _
                        Trim(strValUserSID) & "'", "", "")
            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        lUserId = CLng(Val(strItem))

                    End If
                Next
            End If

            objLogin = Nothing

            Return lUserId

        Catch ex As Exception

        End Try

    End Function

    'SID from User ID
    Public Function ReturnUserSIDFromUserID _
        (ByVal lValUserID As Long) As String

        Try

            Dim lUserSID As String
            Dim arFillControl() As String
            Dim strItem As String

            'Get the employer ID
            Dim objLogin As IMLogin = New IMLogin

            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT SystemSID FROM Users WHERE UserSID = " & _
                        lValUserID, "", "")
            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        lUserSID = CLng(Val(strItem))

                    End If
                Next
            End If

            objLogin = Nothing

            Return lUserSID

        Catch ex As Exception

        End Try

    End Function

    'User Name from User ID
    Public Function ReturnUserNameFromUserID _
        (ByVal lValUserID As Long) As String

        Try

            Dim strUserName As String
            Dim arFillControl() As String
            Dim strItem As String

            'Get the employer ID
            Dim objLogin As IMLogin = New IMLogin

            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT UserName FROM Users WHERE UserID = " & _
                        lValUserID, "", "")
            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        strUserName = strItem

                    End If
                Next
            End If

            objLogin = Nothing

            Return strUserName

        Catch ex As Exception

        End Try
    End Function

    'User Name from User ID
    Public Function ReturnUserIDFromUserName _
        (ByVal strValUserName As String) As Long

        Try

            Dim strUserID As Long
            Dim arFillControl() As String
            Dim strItem As String

            'Get the employer ID
            Dim objLogin As IMLogin = New IMLogin

            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT UserID FROM Users WHERE UserName = '" & _
                        strValUserName & "'", "", "")

            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        strUserID = CLng(Val(strItem))

                    End If
                Next
            End If

            objLogin = Nothing

            Return strUserID

        Catch ex As Exception

        End Try
    End Function

    Public Function ReturnUserPersonalDetailsFromUserID _
        (ByVal lValUserID As Long) As String

        Try

            Dim strUserName As Long
            Dim strUserFirstAndMiddleName As Long
            Dim strUserSurname As Long
            Dim strUserTelephone As Long
            Dim strUserMobile As Long
            Dim strUserIDType As Long
            Dim strUserIDNumber As Long
            Dim arFillControl() As String
            Dim strItem As String
            Dim strReturn As String

            'Get the UserName
            Dim objLogin As IMLogin = New IMLogin

            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT UserName FROM Users WHERE UserID = " & _
                        lValUserID, "", "")
            End With

            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        strUserName = strItem

                    End If
                Next
            End If


            '--strUserFirstAndMiddleName
            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT FirstandMiddleName FROM Users WHERE UserID = " & _
                        lValUserID, "", "")
            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        strUserFirstAndMiddleName = strItem

                    End If
                Next
            End If


            '--strUserSurname
            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT Surname FROM Users WHERE UserID = " & _
                        lValUserID, "", "")
            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        strUserSurname = strItem

                    End If
                Next
            End If


            '--strUserTelephone
            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT PhoneNumber FROM Users WHERE UserID = " & _
                        lValUserID, "", "")
            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        strUserTelephone = strItem

                    End If
                Next
            End If


            '--strUserMobile
            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT MobileNumber FROM Users WHERE UserID = " & _
                        lValUserID, "", "")
            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        strUserMobile = strItem

                    End If
                Next
            End If


            '--strUserIDType
            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT IDTypes FROM Users WHERE UserID = " & _
                        lValUserID, "", "")
            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        strUserIDType = strItem

                    End If
                Next
            End If


            '--strUserIDNumber
            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT IDNumber FROM Users WHERE UserID = " & _
                        lValUserID, "", "")
            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        strUserIDNumber = strItem

                    End If
                Next
            End If

            objLogin = Nothing

            strReturn = "User Name = " & strUserName & Chr(10) & _
            "User's First And MiddleName = " & strUserFirstAndMiddleName & Chr(10) & _
            "User Telephone = " & strUserTelephone & Chr(10) & _
            "User Mobile = " & strUserMobile & Chr(10) & _
            "User ID Type = " & strUserIDType & Chr(10) & _
            "User ID Number = " & strUserIDNumber

            Return strReturn

        Catch ex As Exception

        End Try
    End Function

    Public Function ReturnUserStatusFromUserID _
        (ByVal lValUserID As Long) As String

        Try

            Dim strUserName As Boolean
            Dim arFillControl() As String
            Dim strItem As String

            'Get the employer ID
            Dim objLogin As IMLogin = New IMLogin

            With objLogin
                arFillControl = .FillArray(strAccessConnString, _
                    "SELECT UserStatus FROM Users WHERE UserID = " & _
                        lValUserID, "", "")
            End With


            If Not arFillControl Is Nothing Then
                For Each strItem In arFillControl
                    If Not strItem Is Nothing Then
                        strUserName = strItem

                    End If
                Next
            End If

            objLogin = Nothing

            Return strUserName

        Catch ex As Exception

        End Try
    End Function

#End Region


#Region "DatabaseProcedures"

    Public Sub UserSave()
        'Saves a new country name

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String
        Dim objOrgSID As IMOrganizationSID

        If Trim(strUserName) = "" Or _
            Trim(strFirstandMiddleName) = "" Or _
            Trim(strSurname) = "" Or _
                Trim(strIDTypes) = "" Or _
                    Trim(strIDNumber) = "" _
                            Then

            MsgBox("Please provide an existing" & _
            Chr(10) & "1. User Name" & _
            Chr(10) & "2. User's Surname and First Name" & _
            Chr(10) & "3. ID Type" & _
            Chr(10) & "3. ID Number" _
                           , MsgBoxStyle.Exclamation, _
                           "iManagement - invalid or incomplete information")

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If

        If bUsePasswords = True Then
            If (strOldUserPassword <> strConfirmUserPassword) Or _
                 (strNewUserPassword <> strConfirmUserPassword) Then
                MsgBox("The passwords do not match with the confirmation password.", _
                    MsgBoxStyle.Critical, "iManagement - Passwords do not match")

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub
            End If

            If strOldUserPassword = strNewUserPassword Then
                MsgBox("The Old and New passwords are an identical match (they are similar). This is not allowed.", _
                            MsgBoxStyle.Critical, _
                                "iManagement - Passwords do not match")

                objLogin = Nothing
                datSaved = Nothing

                Exit Sub

            End If
        End If


        'Check if there is an existing series with this name
        If Find("SELECT * FROM Users WHERE  UserName = '" _
                    & strUserName & "'", False) = True Then

            If MsgBox("The User Name already exists." & _
            Chr(10) & "Do you want to update the details?", _
                    MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                            MsgBoxResult.Yes Then

                If bUsePasswords = True Then
                    If MsgBox("Are you sure you want to change the user's passwords?." & _
                    Chr(10) & "This will affect all the user's password in all" & _
                    " the organizations the user is linked to") = MsgBoxResult.No Then

                        objLogin = Nothing
                        datSaved = Nothing

                        Exit Sub
                    End If
                End If



                Update("UPDATE Users SET " & _
                    "Surname = '" & Trim(strSurname) & _
                            "' , NTUserID = '" & Trim(strNTUserID) & _
                            "' , FirstandMiddleName = '" & Trim(strFirstandMiddleName) & _
                            "' , NTSID = " & lNTSID & _
                            " , DigitalSignature = " & lDigitalSignature & _
                            " , PhoneNumber = '" & Trim(strPhoneNumber) & _
                            "' , MobileNumber = '" & Trim(strMobileNumber) & _
                            "' , Organization = '" & Trim(strOrganization) & _
                            "' , IDTypes = '" & Trim(strIDTypes) & _
                            "' , IDNumber = '" & Trim(strIDNumber) & _
                                "' WHERE  strUserName = '" _
                                    & strUserName & "'", True)

                If bUsePasswords = True Then
                    'Changes the password in all organizations related to the user
                    AlterDBPasswordInAllUserOrganizations()

                End If

            End If

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If

        If MsgBox("Are you sure you want to this new user?" _
            , MsgBoxStyle.YesNo, _
                "iManagment - Add new user record?") _
                    = MsgBoxResult.No Then

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If


        Type = "User"
        TypeUserOrGroupID = "User"
        SIDStatus = bUserStatus


        If Save(False, False, False, False) = False Then
            MsgBox("Cannot save Security Identity details. Cannot Save the Group Details." _
                       , MsgBoxStyle.Exclamation, "iManagement - Cannot save the Group's Details")

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If


        If SystemSID = "" Then
            MsgBox("Cannot save Security Identity details. Cannot Save the Group Details." _
                      , MsgBoxStyle.Exclamation, "iManagement - Cannot save the Group's Details")

            objLogin = Nothing
            datSaved = Nothing

            Exit Sub
        End If

        strInsertInto = "INSERT INTO Users (" & _
            "UserName," & _
            "NTUserID," & _
            "Surname," & _
            "FirstandMiddleName," & _
            "NTSID," & _
            "DigitalSignature," & _
            "PhoneNumber," & _
            "MobileNumber," & _
            "Organization," & _
            "IDTypes," & _
            "IDNumber," & _
            "SystemSID" & _
                ") VALUES "

        strSaveQuery = strInsertInto & _
                "('" & Trim(strUserName) & _
                "','" & Trim(strNTUserID) & _
                "','" & Trim(strSurname) & _
                "','" & Trim(strFirstandMiddleName) & _
                "'," & lNTSID & _
                "," & lDigitalSignature & _
                ",'" & Trim(strPhoneNumber) & _
                "','" & Trim(strMobileNumber) & _
                "','" & Trim(strOrganization) & _
                "','" & Trim(strIDTypes) & _
                "','" & Trim(strIDNumber) & _
                "','" & SystemSID & _
                        "')"

        objLogin.ConnectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bSaveSuccess = objLogin.ExecuteQuery(strAccessConnString, _
        strSaveQuery, _
        datSaved)

        objLogin.CloseDb()

        If bSaveSuccess = True Then

            With objOrgSID

                Dim objOvSetup As IMOverallSetup = New IMOverallSetup

                If objOvSetup.Find("SELECT * FROM CompanyMaster WHERE OrganizationName = '" & _
                strOrganizationName & "'", True, False, True) = False Then

                    MsgBox("Please open an existing organization in order to register Business Processes", _
                    MsgBoxStyle.Exclamation, "iManagement - Please Open an existing company")

                    objOvSetup = Nothing
                    Exit Sub

                End If

                .OrganizationID = objOvSetup.OrganizationID
                .SystemSID = SystemSID

                .Save(False, False, False, False)

                objOvSetup = Nothing

            End With

            'AddUserToAllUserOrganizations()

            MsgBox("Record Saved Successfully", MsgBoxStyle.Information, _
            "iManagement - User Details Saved")

        Else

            MsgBox("'Save User' action failed." & _
                " Make sure all mandatory details are entered", _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - User Addition Failed")

        End If


    End Sub

    Private Function AddUserToAllUserOrganizations()

        'Add to system DB
        AddDBUser(strUserName, strConfirmUserPassword, strAccessConnStringADOX, strAccessMdw)

        'Add to org database
        AddDBUser(strUserName, strConfirmUserPassword, strOrgAccessConnString, strAccessMdw)

    End Function

    Private Function AlterDBPasswordInAllUserOrganizations()

        Try

            'AlterDBUserPassword(strUserName, strOldUserPassword, strNewUserPassword, _
            '    strAccessConnString, strAccessMdw)

            'AlterDBUserPassword(strUserName, strOldUserPassword, strNewUserPassword, _
            '    strOrgAccessConnString, strAccessMdw)

        Catch ex As Exception

        End Try

    End Function

    Private Function DeleteDBUserInAllOrganization()

        Try

            'DeleteDBUser(strUserName, strConfirmUserPassword, strAccessConnString, strAccessMdw)

            'DeleteDBUser(strUserName, strConfirmUserPassword, strOrgAccessConnString, strAccessMdw)

        Catch ex As Exception

        End Try
    End Function

    Public Function UserFind(ByVal strQuery As String, _
            ByVal bReturnValues As Boolean, _
                ByVal bUseSIDMasterDetails As Boolean) As Boolean

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
                    Return False
                    Exit Function

                End If

                For Each myDataRows In myDataTables.Rows
                    If bReturnValues = True Then

                        lUserID = myDataRows("UserID")
                        strUserName = myDataRows("UserName")
                        strNTUserID = myDataRows("NTUserID")
                        strSurname = myDataRows("Surname")
                        strFirstandMiddleName = myDataRows("FirstandMiddleName").ToString()
                        lNTSID = myDataRows("NTSID")
                        lDigitalSignature = myDataRows("DigitalSignature")
                        strPhoneNumber = myDataRows("PhoneNumber").ToString()
                        strMobileNumber = myDataRows("MobileNumber").ToString()
                        strOrganization = myDataRows("Organization").ToString()
                        strIDTypes = myDataRows("IDTypes").ToString()
                        strIDNumber = myDataRows("IDNumber").ToString()
                        dtCreationDate = myDataRows("CreationDate")

                        If bUseSIDMasterDetails = True Then
                            SystemSID = myDataRows("SIDMaster.SystemSID")
                        Else
                            SystemSID = myDataRows("SystemSID")

                        End If

                        Find("SELECT * FROM SIDMaster WHERE " & _
                        "SystemSID = '" & SystemSID & "'", True)

                        bUserStatus = SIDStatus


                    End If

                Next

            Next

            Return True
        Else
            Return False
        End If


    End Function

    Public Sub UserDelete()

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        If lUserID = 0 Or Trim(SystemSID) = "" Then
            MsgBox("Cannot Delete due to missing information. Please provide an existing" & _
            " User's Details." _
                           , MsgBoxStyle.Exclamation, _
                           "iManagement - invalid or incomplete information")
            objLogin = Nothing
            datDelete = Nothing

            Exit Sub

        End If

        If MsgBox("Are you sure you want to delete this user's detaisls?" _
        , MsgBoxStyle.YesNo, "iManagement - Delete the user's details?") = MsgBoxResult.No Then

            objLogin = Nothing
            datDelete = Nothing

            Exit Sub
        End If


        If Delete(False, False, False, False) = False Then
            MsgBox("Security ID Deletion Failed. 'Delete User' action failed", _
                           MsgBoxStyle.Exclamation, " User Deletion failed")

            objLogin = Nothing
            datDelete = Nothing

            Exit Sub
        End If


        'Deletion of UserID from Users
        strDeleteQuery = "DELETE * FROM Users WHERE UserID = " & _
                    lUserID

        objLogin.ConnectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
        datDelete)

        'Deletion of UserID from UserGroup
        strDeleteQuery = "DELETE * FROM UserGroup WHERE UserID = " & _
                    lUserID

        objLogin.ConnectString = strAccessConnString
        objLogin.ConnectToDatabase()

        bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
        datDelete)


        objLogin.CloseDb()

        If bDelSuccess = True Then

            DeleteDBUserInAllOrganization()

            MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                "iManagement - User Details Deleted")
        Else
            MsgBox("'Delete User' action failed", _
                MsgBoxStyle.Exclamation, " User Deletion failed")
        End If

    End Sub

    Public Sub UserUpdate(ByVal strUpQuery As String, ByVal DisplaySuccess As Boolean)

        Dim strUpdateQuery As String
        Dim datUpdated As DataSet = New DataSet
        Dim bUpdateSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        strUpdateQuery = strUpQuery

        If (lUserID) <> 0 Then

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            If SystemSID = "" Then
                MsgBox("Cannot Update Security Identity details. Cannot update the User Details." _
                          , MsgBoxStyle.Exclamation, "iManagement - Cannot update the User's Details")

                objLogin = Nothing
                datUpdated = Nothing

                Exit Sub
            End If

            If bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                               "UPDATE SIDMaster SET SIDStatus  = " & bUserStatus & _
                                " WHERE SystemSID = '" & SystemSID & "'", _
                                        datUpdated) = False Then

                MsgBox("Cannot Update Security Identity details. Cannot update the User Details." _
                                        , MsgBoxStyle.Exclamation, "iManagement - Cannot update the User's Details")

                objLogin = Nothing
                datUpdated = Nothing

                Exit Sub
            End If

            bUpdateSuccess = False

            bUpdateSuccess = objLogin.ExecuteQuery(strAccessConnString, _
                                strUpdateQuery, _
                                        datUpdated)

            objLogin.CloseDb()

            If bUpdateSuccess = True Then
                If DisplaySuccess = True Then
                    MsgBox("Record Updated Successfully", MsgBoxStyle.Information, _
                        "iManagement -  User Details Updated")
                End If

            End If

        End If

    End Sub

    'Return all database user names within the workgroup and connection
    'Public Function ReturnDBUSerNames _
    '    (ByVal strValConnString As String, _
    '        ByVal strValWorkGroupFile As String) As String()

    '    Try

    '        Dim cat As ADOX.Catalog
    '        Dim usrNew As ADOX.User
    '        Dim usrLoop As ADOX.User
    '        Dim grpLoop As ADOX.Group
    '        Dim arUsers() As String
    '        Dim i As Long


    '        If Trim(strValConnString) = "" Or Trim(strValWorkGroupFile) = "" Then
    '            Exit Function
    '        End If

    '        cat = New ADOX.Catalog

    '        'Add to access table
    '        cat.ActiveConnection = strValConnString & _
    '                "jet oledb:system database=" & _
    '                strValWorkGroupFile

    '        'Check if the 
    '        If cat Is Nothing Then
    '            Exit Function
    '        End If

    '        With cat
    '            ReDim arUsers(.Users.Count())

    '            i = 0
    '            For Each usrLoop In .Users
    '                arUsers(i) = usrLoop.Name()
    '                i = i + 1
    '            Next

    '        End With

    '        Return arUsers
    '        cat = Nothing

    '    Catch ex As Exception


    '    End Try

    'End Function

    'Check if the user is in the database
    'Public Function CheckIfUserIsInDB _
    '    (ByVal strValUserName As String, _
    '        ByVal strValConnString As String, _
    '            ByVal strValWorkGroupFile As String)

    '    Try

    '        Dim cat As ADOX.Catalog
    '        Dim usrNew As ADOX.User
    '        Dim usrLoop As ADOX.User
    '        Dim grpLoop As ADOX.Group

    '        'Add to access table
    '        cat.ActiveConnection = strValConnString & _
    '                "jet oledb:system database=" & _
    '                strValWorkGroupFile

    '        With cat

    '            ' Create and append new user with an object.
    '            usrNew = New ADOX.User

    '            usrNew.Name = "Pat Smith"
    '            usrNew.ChangePassword("", "Password1")
    '            .Users.Append(usrNew)



    '        End With


    '    Catch ex As Exception

    '    End Try
    'End Function

    'Add user to the Microsoft Access database
    Private Function AddDBUser _
        (ByVal strValUserName As String, _
            ByVal strValUserPassword As String, _
                ByVal strValConnString As String, _
                    ByVal strValWorkGroupFile As String)
        'Try

        '    Dim adoCn As ADODB.Connection = New ADODB.Connection
        '    Dim cat As ADOX.Catalog = New ADOX.Catalog
        '    Dim adoRcs As ADODB.Recordset = New ADODB.Recordset

        '    Dim usrNew As ADOX.User
        '    Dim daoCN As DAO.Connection
        '    Dim DaoWsp As DAO.Workspace
        '    Dim daoDB As DAO.Database

        '    Dim daoDBEng As DAO.DBEngine = New DAO.DBEngine
        '    Dim usrNew2 As DAO.User = New DAO.User
        '    daoDBEng.OpenDatabase("iMSysManager.mdb")

        '    daoDBEng.SystemDB = "H:\Systems\iManagement Systems\iManagementWorkGroupFile.mdw"
        '    DaoWsp = daoDBEng.CreateWorkspace("", strDBUserName, strDBPassword)

        '    With DaoWsp

        '        usrNew2 = .CreateUser(strValUserName)
        '        usrNew2.PID = SystemSID
        '        usrNew2.Password = strValUserPassword
        '        MsgBox(.Users.Count)
        '        .Users.Append(usrNew2)


        '    End With

        '    usrNew2 = Nothing
        '    DaoWsp = Nothing
        '    daoCN = Nothing





        '    adoCn.Open(strAccessConnString)
        '    adoCn.Execute("CREATE USER " & strValUserName & _
        '                " " & _
        '                    strValUserPassword & " " & SystemSID, Nothing, 1)



        '    adoCn = Nothing

        '    cat.ActiveConnection = adoCn

        '    Add to access table
        '    cat.ActiveConnection = strValConnString '& _
        '    "jet oledb:system database=" & _
        '    strValWorkGroupFile & ""


        '    With cat

        '         Create and append new user with an object.
        '        usrNew = New ADOX.User
        '        usrNew.Name = strUserName
        '        usrNew.ChangePassword("", strValUserPassword)
        '        .Users.Append(usrNew)

        '    End With

        '    usrNew = Nothing

        'Catch ex As Exception

        'End Try

    End Function

    'Delete user to the Microsoft Access database
    'Private Function DeleteDBUser _
    '    (ByVal strUserName As String, _
    '            ByVal strValUserPassword As String, _
    '                ByVal strValConnString As String, _
    '                    ByVal strValWorkGroupFile As String)

    '    Try

    '        Dim cat As ADOX.Catalog

    '        With cat
    '            .Users.Delete(strUserName)
    '        End With


    '    Catch ex As Exception

    '    End Try

    'End Function

    'Alter user to the Microsoft Access database
    'Private Function AlterDBUserPassword(ByVal strValUserName As String, _
    '        ByVal strValOldUserPassword As String, _
    '            ByVal strValNewUserPassword As String, _
    '                ByVal strValConnString As String, _
    '                    ByVal strValWorkGroupFile As String)

    '    Try

    '        Dim cat As ADOX.Catalog
    '        Dim usrNew As ADOX.User

    '        cat = New ADOX.Catalog

    '        'Add to access table
    '        cat.ActiveConnection = strValConnString & _
    '                "jet oledb:system database=" & _
    '                strValWorkGroupFile

    '        With cat

    '            ' Create and append new user with an object.
    '            usrNew = New ADOX.User
    '            usrNew.Name = strUserName
    '            usrNew.ChangePassword(strValOldUserPassword, strValNewUserPassword)

    '        End With

    '        usrNew = Nothing

    '    Catch ex As Exception

    '    End Try

    'End Function



#End Region


End Class
