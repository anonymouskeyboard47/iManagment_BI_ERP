Public Class IMOrderDocuments

#Region "PrivateVariables"

    Private lOrderDocumentsID As Long
    Private lOrderID As Long
    Private lDocumentID As Long

#End Region

#Region "Properties"

    Public Property OrderDocumentsID() As Long

        Get
            Return lOrderDocumentsID
        End Get

        Set(ByVal Value As Long)
            lOrderDocumentsID = Value
        End Set

    End Property

    Public Property OrderID() As Long

        Get
            Return lOrderID
        End Get

        Set(ByVal Value As Long)
            lOrderID = Value
        End Set

    End Property

    Public Property DocumentID() As Long

        Get
            Return lDocumentID
        End Get

        Set(ByVal Value As Long)
            lDocumentID = Value
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

            If lOrderID = 0 Or _
                lDocumentID = 0 Then

                If DisplayErrorMessages = True Then

                    MsgBox("Please provide the following details in" & _
                " order to link a Order's Document " & _
                Chr(10) & "to iManagement Document Manager:" & _
                Chr(10) & "1. Existing Order" & _
                Chr(10) & "2. Existing Document" & _
                MsgBoxStyle.Critical, _
            "iManagement - Save Action Failed")

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            Else

                If Find("SELECT * FROM OrderDocuments WHERE lDocumentID = " & _
                    lDocumentID & " AND OrderID = " & lOrderID _
                        , False) = False Then

                    If MsgBox("The Order Document Details already exists." & _
                        Chr(10) & "Do you want to update the details?", _
                                MsgBoxStyle.YesNo, "iManagement - Record Exists") = _
                                        MsgBoxResult.Yes Then

                        Update("UPDATE OrderDocuments SET " & _
                                    " AND DocumentID = " & lDocumentID & _
                                    " AND OrderID = " & lOrderID & _
                                        " WHERE  OrderDocumentsID = " _
                                            & lOrderDocumentsID)

                    End If

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If

                strInsertInto = "INSERT INTO OrderDocuments (" & _
                    "DocumentID," & _
                    "OrderID" & _
                        ") VALUES "

                strSaveQuery = strInsertInto & _
                        "(" & lDocumentID & _
                        "," & lOrderID & _
                                ")"

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bSaveSuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, _
                strSaveQuery, _
                datSaved)


                objLogin.CloseDb()

                If bSaveSuccess = True Then
                    If DisplaySuccess = True Then
                        MsgBox("Order Documents Saved Successfully.", _
                            MsgBoxStyle.Information, _
                                "iManagement - Record Saved Successfully")

                    End If

                Else

                    If DisplayFailure = True Then
                        MsgBox("'Save Order Documents action failed." & _
                " Make sure all mandatory details are entered.", _
                MsgBoxStyle.Exclamation, _
                "iManagement -  Addition Failed")

                    End If

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function

                End If
            End If

            objLogin = Nothing
            datSaved = Nothing

            Return True

        Catch ex As Exception
            If DisplayErrorMessages = True Then
                MsgBox(ex.Message.ToString, _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Critical System Error")
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

            objLogin.connectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bQuerySuccess = objLogin.ExecuteQuery _
                    (strOrgAccessConnString, strQuery, _
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

                    'Whether to fill properties with values or not
                    If ReturnStatus = True Then

                        For Each myDataRows In myDataTables.Rows

                            'lProductID = _
                            '    myDataRows("ProductID")
                            'lSupplierOrganizationID = _
                            '    myDataRows("SupplierOrganizationID")
                            'lRangeID = _
                            '    myDataRows("RangeID")
                            'dbPricePerUnit = _
                            '    myDataRows("PricePerUnit")
                            'dtDateEntered = _
                            '    myDataRows("DateEntered")
                            'dtExpiryDate = _
                            '    myDataRows("ExpiryDate")
                            'dbMinNumOfUnits = _
                            '   myDataRows("MinNumOfUnits")
                            'bPriceRangeStatus = _
                            '   myDataRows("PriceRangeStatus")

                        Next

                    End If

                Next
                Return True

            Else
                Return False

            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString, _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Critical System Error")

        End Try

    End Function

    Public Sub Delete(ByVal strDelQuery As String)

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = strDelQuery

            If lDocumentID <> 0 Or lOrderID <> 0 _
                              Then

                objLogin.connectString = strAccessConnString
                objLogin.ConnectToDatabase()

                bDelSuccess = objLogin.ExecuteQuery(strAccessConnString, strDeleteQuery, _
                datDelete)

               

                objLogin.CloseDb()

                If bDelSuccess = True Then
                    MsgBox("Order Documents Details Deleted", MsgBoxStyle.Information, _
                        "iManagement - Record Deleted Successfully")

                Else

                    MsgBox("'Delete Order Documents' action failed", _
                        MsgBoxStyle.Exclamation, "Order Documents Deletion failed")

                    objLogin.RollbackTheTrans()

                End If

            Else

                MsgBox("Cannot Delete. Please select an existing Order Document Detail", _
                        MsgBoxStyle.Exclamation, "iManagement -Missing Information")

                objLogin.RollbackTheTrans()

            End If
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

            If lDocumentID <> 0 Or lOrderID <> 0 _
                            Then

                objLogin.connectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    MsgBox("Record Updated Successfully", _
                        MsgBoxStyle.Information, _
                            "iManagement -  Order Document Updated")
                End If

            End If

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

            Return strTextFieldData
            datFillData.Dispose()

        Catch ex As Exception

        End Try

    End Function


#End Region
End Class
