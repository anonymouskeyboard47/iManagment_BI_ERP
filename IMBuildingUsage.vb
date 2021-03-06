
Public Class IMBuildingUsage


#Region "PrivateVariables"

    Private lBuildingID As Long
    Private lUsageID As Long

#End Region


#Region "Properties"

    Public Property UsageID() As Long

        Get
            Return lUsageID
        End Get

        Set(ByVal Value As Long)
            lUsageID = Value
        End Set

    End Property

    Public Property BuildingID() As Long

        Get
            Return lBuildingID
        End Get

        Set(ByVal Value As Long)
            lBuildingID = Value
        End Set

    End Property

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "DatabaseProcedures"

    Public Function Save(ByVal DisplayErrorMessages As Boolean, _
        ByVal DisplayConfirmation As Boolean, _
            ByVal DisplayFailure As Boolean, _
                ByVal DisplaySuccess As Boolean) As Boolean

        Dim strSaveQuery As String
        Dim datSaved As DataSet = New DataSet
        Dim bSaveSuccess As Boolean
        Dim objLogin As IMLogin = New IMLogin
        Dim strInsertInto As String


        Try

            If lBuildingID = 0 Or _
                lUsageID = 0 Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please provide the Building and the Usage that you want to save"

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM BuildingUsage WHERE " & _
            "BuildingID = " & lBuildingID, False) = True Then

                Update("UPDATE BuildingUsage SET " & _
                            "UsageID = " & lUsageID)

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'If MsgBox(" Are you sure you want to add this Cost Centre?", _
            '            MsgBoxStyle.YesNo, _
            '            "iManagement - Add New Record?") = _
            '            MsgBoxResult.No Then

            '    datSaved = Nothing
            '    objLogin = Nothing
            '    Exit Function
            'End If


            strInsertInto = "INSERT INTO BuildingUsage (" & _
                "BuildingID," & _
                "UsageID" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lBuildingID & _
                    "," & lUsageID & _
                            ")"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, _
            strSaveQuery, datSaved)


            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccess = True Then

                    ReturnError += "Building's Usage Saved Successfully."

                End If

                Return True

            Else

                If DisplayFailure = True Then

                    ReturnError += "'Save Building Usage' action failed." & _
            " Make sure all mandatory details are entered."

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If

            objLogin = Nothing
            datSaved = Nothing


        Catch ex As Exception
            If DisplayErrorMessages = True Then
                ReturnError += ex.Message.ToString

            End If
        End Try

    End Function

    Public Function Find(ByVal strQuery As String, _
                        ByVal ReturnStatus As Boolean) As Boolean
        'Query must contain at least rows from Sequence

        Try

            Dim datRetData As DataSet = New DataSet
            Dim bQuerySuccess As Boolean
            Dim myDataTables As DataTable
            Dim myDataColumns As DataColumn
            Dim myDataRows As DataRow
            Dim objLogin As IMLogin = New IMLogin

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bQuerySuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strQuery, datRetData)

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
                        datRetData = Nothing
                        objLogin = Nothing

                        Return False
                        Exit Function

                    End If

                    'Whether to fill properties with values or not
                    If ReturnStatus = True Then

                        For Each myDataRows In myDataTables.Rows

                            lBuildingID = _
                                myDataRows("BuildingID")
                            lUsageID = _
                                myDataRows("UsageID")

                        Next
                    End If
                Next

                Return True

            End If

        Catch ex As Exception
            ReturnError += ex.Message.ToString

        End Try

    End Function

    Public Function Delete(Optional ByVal bDisplayError As Boolean = False, _
    Optional ByVal bDisplayConfirm As Boolean = False, _
    Optional ByVal bDisplaySuccess As Boolean = False) As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            If lBuildingID = 0 Then
                ReturnError += "Cannot Delete. Please select an " & _
                    "existing Building Usage Detail"

                datDelete = Nothing
                objLogin = Nothing

                Exit Function
            End If

            'If bDisplayConfirm = True Then
            '    If MsgBox(" Are you sure you want to Delete this Cost Centre?", _
            '                             MsgBoxStyle.YesNo, _
            '                             "iManagement - Delete Record?") = _
            '                              MsgBoxResult.No Then

            '        datDelete = Nothing
            '        objLogin = Nothing
            '        Exit Function
            '    End If
            'End If


            strDeleteQuery = "DELETE * FROM BuildingUsage WHERE " & _
            "BuildingID = " & lBuildingID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            datDelete = Nothing
            objLogin = Nothing

            If bDelSuccess = True Then
                ReturnError += "Building's Usage Details Deleted"
                Return True
            Else

                ReturnError += "'Delete Building's Usage action failed"

            End If

        Catch ex As Exception

        End Try

    End Function

    Public Sub Update(ByVal strUpQuery As String)

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If lBuildingID <> 0 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery _
                                (strOrgAccessConnString, _
                                    strUpdateQuery, datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    ReturnError += "Building's Usage details updated Successfully"
                End If

            End If

            datUpdated = Nothing
            objLogin = Nothing


        Catch ex As Exception

        End Try

    End Sub

#End Region


End Class
