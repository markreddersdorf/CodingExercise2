Public Class clsBank

    Private m_strBankName As String

    Public ReadOnly Property strBankName As String
        Get
            Return m_strBankName
        End Get
    End Property

    Public Sub New(ByVal strName As String)

        m_strBankName = "Second National Bank"

    End Sub

    Public Function ListAccounts() As String(,)

        Dim strAccountNames(5, 5) As String

        strAccountNames(0, 0) = "John Jones"
        strAccountNames(0, 1) = "Checking"
        strAccountNames(1, 0) = "Sarah Smith"
        strAccountNames(1, 1) = "Checking"
        strAccountNames(2, 0) = "MegaCorpInvestments"
        strAccountNames(2, 1) = "CorporateInvestment"
        strAccountNames(3, 0) = "SmallCompany"
        strAccountNames(3, 1) = "CorporateInvestment"
        strAccountNames(4, 0) = "Sarah Smith"
        strAccountNames(4, 1) = "IndividualInvestment"

        Return strAccountNames

    End Function

    Public Function GetAccounts() As String(,)

        Dim strAccounts As String(,)
        Dim i As Integer = 0
        Dim j As Integer = 0

        strAccounts = ListAccounts()

        Return strAccounts

    End Function

End Class
