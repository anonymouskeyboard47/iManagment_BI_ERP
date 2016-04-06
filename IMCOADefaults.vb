Option Explicit On 
'Option Strict On

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCOADefaults
    Private lCOAAccountNr As Long
    Private strDefaultText As String
    Private lCOADefaultID As Long


#Region "Properties"

    Public Property COAAccountNr() As Long

        Get
            Return lCOAAccountNr
        End Get

        Set(ByVal Value As Long)
            lCOAAccountNr = Value
        End Set

    End Property

    Public Property DefaultText() As String

        Get
            Return strDefaultText
        End Get

        Set(ByVal Value As String)
            strDefaultText = Value
        End Set

    End Property

    Public Property COADefaultID() As Long

        Get
            Return lCOADefaultID
        End Get

        Set(ByVal Value As Long)
            lCOADefaultID = Value
        End Set

    End Property


#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region

#Region "GeneralProcedures"

    Public Function SaveAccountingADefaultCOA _
     (ByVal lValCOAAccountNr As Long, _
         ByVal strValDefaultText As String) As Boolean

        Try


            Dim objCOADef As IMCOADefaults = New IMCOADefaults

            With objCOADef
                .DefaultText = strValDefaultText
                .COAAccountNr = lValCOAAccountNr

                Return .Save(True, True, True)

            End With

            objCOADef = Nothing

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
            Dim objCOA As IMChartOfAccount = New IMChartOfAccount

            If Trim(strOrganizationName) = "" Then
                MsgBox("Please open an existing company.", _
                    MsgBoxStyle.Critical, _
                        "iManagement - Select an existing company")
                objLogin = Nothing
                datSaved = Nothing
                objCOA = Nothing

                Exit Function
            End If


            If Trim(strDefaultText) = "" Or _
               lCOAAccountNr = 0 Then

                MsgBox("You must provide an account number and its Default accompanying text." _
                , MsgBoxStyle.Critical, _
                "iManagement - Invalid or incomplete data")

                objLogin = Nothing
                datSaved = Nothing
                objCOA = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM COADefaults WHERE DefaultText = '" _
                & Trim(strDefaultText) & "'", False) = True Then

                If Find("SELECT * FROM COADefaults WHERE DefaultText = '" _
                    & Trim(strDefaultText) & "' AND COAAccountNr = " & _
                        lCOAAccountNr, False) = True Then

                    objLogin = Nothing
                    datSaved = Nothing
                    objCOA = Nothing

                    Exit Function
                End If


                If MsgBox("Do you want to update the existing '" & _
                    strDefaultText & "' Default Account details?", _
                    MsgBoxStyle.YesNo, "iManagement - Update  details?") _
                        = MsgBoxResult.Yes Then

                    Dim strItem As String
                    Dim arItem() As String
                    Dim lValCOAAccountNr As Long

                    arItem = objLogin.FillArray(strOrgAccessConnString, _
                        "SELECT COAAccountNr FROM COADefaults " & _
                    " WHERE DefaultText = '" & strDefaultText & _
                    "'", "", "")

                    If Not arItem Is Nothing Then
                        For Each strItem In arItem
                            If Not strItem Is Nothing Then
                                lValCOAAccountNr = CLng(Val(strItem))

                            End If
                        Next
                    End If

                    bSaveSuccess = Update("UPDATE COADefaults SET" & _
                        " COAAccountNr =  " & lCOAAccountNr & _
                        " WHERE DefaultText = '" & Trim(strDefaultText) & "'")

                    If bSaveSuccess = True Then


                        objCOA.COAAccountNr = lValCOAAccountNr
                        objCOA.UnReserveAccount()

                        objCOA.COAAccountNr = lCOAAccountNr
                        objCOA.ReservedBy = Trim(strDefaultText)
                        objCOA.ReserveAccount(False)

                    End If

                End If

                objLogin = Nothing
                datSaved = Nothing
                objCOA = Nothing

                Exit Function
            End If


            'Reserve a new Default Account
            With objCOA
                .COAAccountNr = lCOAAccountNr
                .ReservedBy = strDefaultText

                bSaveSuccess = .ReserveAccount(True)

            End With


            If bSaveSuccess = False Then
                objLogin = Nothing
                datSaved = Nothing
                objCOA = Nothing
                Exit Function
            End If

            bSaveSuccess = False


            strInsertInto = "INSERT INTO COADefaults (" & _
                "DefaultText," & _
                "COAAccountNr" & _
                ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & Trim(strDefaultText) & _
                    "', " & lCOAAccountNr & _
                    ")"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)

            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccessMessages = True Then
                    MsgBox("Default Chart Account Number for '" & _
                        strDefaultText & "' Successfully", _
                            MsgBoxStyle.Information, _
                                "iManagement - Record Saved")

                End If

                Return True
            Else

                If DisplayFailureMessages = True Then
                    MsgBox("'Saving default '" & strDefaultText & "' Chart Account Number' action failed." & _
                        " Make sure all mandatory details are entered.", _
                            MsgBoxStyle.Exclamation, _
                                "iManagement - Save Record Failed")
                End If
            End If

            objLogin = Nothing
            datSaved = Nothing
            objCOA = Nothing

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
                        lCOAAccountNr = _
                                myDataRows("COAAccountNr")
                        strDefaultText = _
                                myDataRows("DefaultText").ToString
                        lCOADefaultID = _
                            myDataRows("COADefaultID")

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
    Public Function Delete(ByVal strDelQuery As String) As Boolean

        Dim strDeleteQuery As String
        Dim datDelete As DataSet = New DataSet
        Dim bDelSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin

        Try

            If Trim(strDefaultText) = "" Then
                MsgBox("Cannot delete the Chart Of Account Default " & _
                "Details. Please select an existing Chart Of Account Default", _
                        MsgBoxStyle.Exclamation, _
                            "iManagement -Missing Information")

                objLogin = Nothing
                datDelete = Nothing
                Exit Function
            End If

            strDeleteQuery = "DELETE * FROM COADefaults WHERE " & _
               "DefaultText = '" & strDefaultText & "'"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, strDeleteQuery, _
            datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                MsgBox("Record Deleted Successfully", MsgBoxStyle.Information, _
                    "iManagement - Chart Of Account Default Deleted")
            Else
                MsgBox("'Chart Of Account Default delete' action failed", _
                    MsgBoxStyle.Exclamation, _
                        "Chart Of Account Default Deletion failed")
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
                MsgBox("Chart Of Account Default Record Updated Successfully", _
                    MsgBoxStyle.Information, _
                    "iManagement - Record Updated")

                Return True
            End If

            objLogin = Nothing
            datUpdated = Nothing

        Catch ex As Exception

        End Try


    End Function


#End Region

End Class
