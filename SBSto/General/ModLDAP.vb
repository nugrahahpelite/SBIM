Module ModLDAP
    Public Function fbValidateADUser(ByVal vriUsername As String, ByVal vriPassword As String) As Boolean
        'Find valid user in Active Directory
        Dim Success As Boolean = False
        Dim Entry As New System.DirectoryServices.DirectoryEntry("LDAP://SUMBERBERKAT.LOCAL", vriUsername, vriPassword, DirectoryServices.AuthenticationTypes.Secure)
        Dim Searcher As New System.DirectoryServices.DirectorySearcher(Entry)

        Searcher.SearchScope = DirectoryServices.SearchScope.OneLevel
        Try
            Dim Results As System.DirectoryServices.SearchResult = Searcher.FindOne
            Success = Not (Results Is Nothing)
        Catch ex As Exception
            Success = False
        End Try
        Return Success
    End Function
End Module
