
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMPlots

#Region "PrivateVariables"

    Private lPlotID As Long
    Private strTitleDeedFileNo As String
    Private strLandReferenceNo As String
    Private dbPlotArea As Double
    Private dtDateCreated As Date
    Private strRegistryMapSheetNo As String
    Private strParcelNo As String

#End Region


#Region "Properties"

    Public Property DateCreated() As Date

        Get
            Return dtDateCreated
        End Get

        Set(ByVal Value As Date)
            dtDateCreated = Value
        End Set

    End Property

    Public Property PlotID() As Long

        Get
            Return lPlotID
        End Get

        Set(ByVal Value As Long)
            lPlotID = Value
        End Set

    End Property

    Public Property TitleDeedFileNumber() As String

        Get
            Return strTitleDeedFileNo
        End Get

        Set(ByVal Value As String)
            strTitleDeedFileNo = Value
        End Set

    End Property

    Public Property LandReferenceNo() As String

        Get
            Return strLandReferenceNo
        End Get

        Set(ByVal Value As String)
            strLandReferenceNo = Value
        End Set

    End Property

    Public Property PlotArea() As Double

        Get
            Return dbPlotArea
        End Get

        Set(ByVal Value As Double)
            dbPlotArea = Value
        End Set

    End Property

    Public Property RegistryMapSheetNo() As String

        Get
            Return strRegistryMapSheetNo
        End Get

        Set(ByVal Value As String)
            strRegistryMapSheetNo = Value
        End Set

    End Property

    Public Property ParcelNo() As String

        Get
            Return strParcelNo
        End Get

        Set(ByVal Value As String)
            strParcelNo = Value
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

            If Trim(strTitleDeedFileNo) = "" Or _
                Trim(strLandReferenceNo) = "" Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please provide the plot's Title\Deed " & _
                    "File Number and its Land Reference Number"

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If Find("SELECT * FROM PlotMaster WHERE " & _
            "TitleDeedFileNumber = '" & _
            Trim(strTitleDeedFileNo) & _
            "' OR LandReferenceNo = '" & strLandReferenceNo & "'", False) _
            = True Then

                Update("UPDATE PlotMaster SET " & _
                            "ParcelNo = '" & _
                            Trim(strParcelNo) & _
                            "', PlotArea = " & _
                            dbPlotArea & _
                            "', RegistryMapSheetNo = " & _
                            strRegistryMapSheetNo & _
                            " WHERE TitleDeedFileNumber = '" & _
                            strTitleDeedFileNo & _
                            "' OR LandReferenceNo = '" & _
                            strLandReferenceNo & "'")

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


            strInsertInto = "INSERT INTO PlotMaster (" & _
                "LandReferenceNo," & _
                "TitleDeedFileNumber," & _
                "PlotArea," & _
                "RegisryMapSheetNo," & _
                "ParcelNo" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & Trim(strLandReferenceNo) & _
                    "','" & Trim(strTitleDeedFileNo) & _
                    "'," & dbPlotArea & _
                    ",'" & strRegistryMapSheetNo & _
                    "','" & strParcelNo & _
                            "')"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)


            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccess = True Then

                    ReturnError += "Plot Saved Successfully."

                End If

                Return True

            Else

                If DisplayFailure = True Then

                    ReturnError += "'Save Plot' action failed." & _
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

                            lPlotID = _
                                myDataRows("PlotID")
                            strLandReferenceNo = _
                                myDataRows("LandReferenceNo").ToString
                            strTitleDeedFileNo = _
                                myDataRows("TitleDeedFileNo").ToString
                            dbPlotArea = _
                                myDataRows("PlotArea")
                            strRegistryMapSheetNo = _
                                myDataRows("RegistryMapSheetNo")
                            strParcelNo = _
                                myDataRows("ParcelNo").ToString
                            dtDateCreated = _
                               myDataRows("DateCreated")

                        Next
                    End If
                Next

                Return True

            End If



        Catch ex As Exception
            returnerror +=ex.Message.ToString

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

            If lPlotID = 0 Then
                ReturnError += "Cannot Delete. Please select an " & _
                    "existing Plot Detail"

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


            strDeleteQuery = "DELETE * FROM PlotMaster WHERE " & _
            "PlotID = " & lPlotID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            datDelete = Nothing
            objLogin = Nothing

            If bDelSuccess = True Then
                ReturnError += "Plot Details Deleted"
                Return True
            Else

                ReturnError += "'Delete Plot action failed"

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

            If lPlotID <> 0 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery _
                                (strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    ReturnError += "Plot's details updated Successfully"
                End If

            End If

            datUpdated = Nothing
            objLogin = Nothing


        Catch ex As Exception

        End Try

    End Sub

#End Region


End Class
