Public Class IMSystemConfiguration

#Region "GeneralVariables"

    Private strRegistryPath As String

#End Region

#Region "Properties"

    Public Property RegistryKeyPath() As String

        Get
            Return strRegistryPath
        End Get

        Set(ByVal Value As String)
            strRegistryPath = Value
        End Set

    End Property


#End Region


#Region "RegistryKeyProcedures"

    Public Function DoesKeyExist() As Boolean
        '*****************************************************
        '** USED TO CHECK IF A PARTICULAR REGISTRY KEY EXISTS
        '*****************************************************

        Dim regVersion As Microsoft.Win32.RegistryKey

        regVersion = _
       Microsoft.Win32.Registry.CurrentUser.OpenSubKey _
       (strRegistryPath, True)


        If regVersion Is Nothing Then
            'If the registry value is inexistent
            Return False
        Else
            'If the registry value exists
            Return True
        End If
    End Function


    Public Function ReadRegistryKey() As String
        '***********************************************************
        '*** USED TO RETURN THE VALUE OF A PARTICULAR REGISTRY KEY
        '***********************************************************

        Dim key As Microsoft.Win32.RegistryKey
        key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(strRegistryPath)
        Dim name As String = CType(key.GetValue(strRegistryPath), String)

    End Function

#End Region

    
End Class
