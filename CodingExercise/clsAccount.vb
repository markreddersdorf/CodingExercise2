Imports System

Public Class clsAccount

    Private m_strAccountOwner As String
    Private m_strAccountType As String
    Private m_decBalance As Decimal

    Public Property strAccountOwner As String
        Get
            Return m_strAccountOwner
        End Get
        Set(value As String)
            m_strAccountOwner = value
        End Set
    End Property

    Public Property strAccountType As String
        Get
            Return m_strAccountType
        End Get
        Set(value As String)
            m_strAccountType = value
        End Set
    End Property

    Public Property decBalance As Decimal
        Get
            Return m_decBalance
        End Get
        Set(value As Decimal)
            m_decBalance = value
        End Set
    End Property

    Public Sub New(ByVal strAcctType As String, ByVal strAccountOwner As String)

        m_strAccountOwner = strAccountOwner
        m_strAccountType = strAccountType
        m_decBalance = GetBalance()

    End Sub

    Public Function Deposit(ByVal decAmount As Decimal, ByVal strAcctOwner As String, ByVal strAcctType As String) As String

        Dim objChecking As clsCheckingAccount = Nothing
        Dim objCorpInv As clsCorporateInvestmentAccount = Nothing
        Dim objIndvInv As clsIndividualInvestmentAccount = Nothing
        Dim strReturn As String = ""

        Try
            Select Case strAcctType
                Case "Checking"
                    objChecking = New clsCheckingAccount(strAcctOwner, strAcctType)
                    objChecking.decBalance = GetBalance()

                    objChecking.decBalance += decAmount
                    strReturn = "Successfully deposited."
                Case "CorporateInvestment"
                    objCorpInv = New clsCorporateInvestmentAccount(strAcctOwner, strAcctType)
                    objCorpInv.decBalance = GetBalance()

                    objCorpInv.decBalance += decAmount
                    strReturn = "Successfully deposited."
                Case "IndividualInvestment"
                    objIndvInv = New clsIndividualInvestmentAccount(strAcctOwner, strAcctType)
                    objIndvInv.decBalance = GetBalance()

                    objIndvInv.decBalance += decAmount
                    strReturn = "Successfully deposited."
            End Select
        Catch ex As Exception
            strReturn = "Deposit failed: " & ex.Message
        Finally
            objChecking = Nothing
            objCorpInv = Nothing
            objIndvInv = Nothing
        End Try

        Return strReturn

    End Function

    Public Function Withdrawl(ByVal decAmount As Decimal, ByVal strAcctOwner As String, ByVal strAcctType As String) As String

        Dim objChecking As clsCheckingAccount = Nothing
        Dim objCorpInv As clsCorporateInvestmentAccount = Nothing
        Dim objIndvInv As clsIndividualInvestmentAccount = Nothing
        Dim strReturn As String = ""

        Try
            Select Case strAcctType
                Case "Checking"
                    objChecking = New clsCheckingAccount(strAcctOwner, strAcctType)
                    objChecking.decBalance = GetBalance()

                    If decAmount >= objChecking.decBalance Then
                        strReturn = "Insufficient Funds."
                    Else
                        objChecking.decBalance -= decAmount
                        strReturn = "Successfully withdrawn."
                    End If
                Case "CorporateInvestment"
                    objCorpInv = New clsCorporateInvestmentAccount(strAcctOwner, strAcctType)
                    objCorpInv.decBalance = GetBalance()

                    If decAmount >= objCorpInv.decBalance Then
                        strReturn = "Insufficient Funds."
                    Else
                        objCorpInv.decBalance -= decAmount
                        strReturn = "Successfully withdrawn."
                    End If
                Case "IndividualInvestment"
                    If decAmount > 1000.0 Then
                        strReturn = "Withdrawl amount exceeds the daily maximum."
                    Else
                        objIndvInv = New clsIndividualInvestmentAccount(strAcctOwner, strAcctType)
                        objIndvInv.decBalance = GetBalance()

                        If decAmount >= objIndvInv.decBalance Then
                            strReturn = "Insufficient Funds."
                        Else
                            objIndvInv.decBalance -= decAmount
                            strReturn = "Successfully withdrawn."
                        End If
                    End If
            End Select
        Catch ex As Exception
            strReturn = "Withdrawl failed: " & ex.Message
        Finally
            objChecking = Nothing
            objCorpInv = Nothing
            objIndvInv = Nothing
        End Try

        Return strReturn

    End Function

    Public Function Transfer(ByVal decAmount As Decimal, ByVal strFromAcctOwner As String, ByVal strFromAcctType As String, ByVal strToAcctOwner As String, ByVal strToAcctType As String) As String

        Dim objChecking1 As clsCheckingAccount = Nothing
        Dim objChecking2 As clsCheckingAccount = Nothing
        Dim objCorpInv1 As clsCorporateInvestmentAccount = Nothing
        Dim objCorpInv2 As clsCorporateInvestmentAccount = Nothing
        Dim objIndvInv1 As clsIndividualInvestmentAccount = Nothing
        Dim objIndvInv2 As clsIndividualInvestmentAccount = Nothing

        Dim decBalanceTo As Decimal = 0.0
        Dim decBalanceFrom As Decimal = 0.0
        Dim strReturn As String = ""

        Try
            Select Case strFromAcctType
                Case "Checking"
                    objChecking1 = New clsCheckingAccount(strFromAcctOwner, strFromAcctType)
                    m_strAccountType = strFromAcctType
                    objChecking1.decBalance = GetBalance()

                    Select Case strToAcctType
                        Case "Checking"
                            objChecking2 = New clsCheckingAccount(strToAcctOwner, strToAcctType)
                            m_strAccountType = strToAcctType
                            objChecking2.decBalance = GetBalance()

                            If decAmount > objChecking1.decBalance Then
                                strReturn = "Insufficient Funds."
                                Return strReturn
                                Exit Function
                            End If

                            objChecking1.decBalance -= decAmount
                            objChecking2.decBalance += decAmount

                            strReturn = "Transfer Successful."
                        Case "CorporateInvestment"
                            objCorpInv2 = New clsCorporateInvestmentAccount(strToAcctOwner, strToAcctType)
                            m_strAccountType = strToAcctType
                            objCorpInv2.decBalance = GetBalance()

                            If decAmount >= objChecking1.decBalance Then
                                strReturn = "Insufficient Funds."
                                Return strReturn
                                Exit Function
                            End If

                            objChecking1.decBalance -= decAmount
                            objCorpInv2.decBalance += decAmount

                            strReturn = "Transfer Successful."
                        Case "IndividualInvestment"
                            objIndvInv2 = New clsIndividualInvestmentAccount(strToAcctOwner, strToAcctType)
                            m_strAccountType = strToAcctType
                            objIndvInv2.decBalance = GetBalance()

                            If decAmount >= objChecking1.decBalance Then
                                strReturn = "Insufficient Funds."
                                Return strReturn
                                Exit Function
                            End If

                            objChecking1.decBalance -= decAmount
                            objIndvInv2.decBalance += decAmount

                            strReturn = "Transfer Successful."
                    End Select
                Case "CorporateInvestment"
                    objCorpInv1 = New clsCorporateInvestmentAccount(strFromAcctOwner, strFromAcctType)
                    m_strAccountType = strFromAcctType
                    objCorpInv1.decBalance = GetBalance()

                    Select Case strToAcctType
                        Case "Checking"
                            objChecking2 = New clsCheckingAccount(strToAcctOwner, strToAcctType)
                            m_strAccountType = strToAcctType
                            objChecking2.decBalance = GetBalance()

                            If decAmount > objCorpInv1.decBalance Then
                                strReturn = "Insufficient Funds."
                                Return strReturn
                                Exit Function
                            End If

                            objCorpInv1.decBalance -= decAmount
                            objChecking2.decBalance += decAmount

                            strReturn = "Transfer Successful."
                        Case "CorporateInvestment"
                            objCorpInv2 = New clsCorporateInvestmentAccount(strToAcctOwner, strToAcctType)
                            m_strAccountType = strToAcctType
                            objCorpInv2.decBalance = GetBalance()

                            If decAmount > objCorpInv1.decBalance Then
                                strReturn = "Insufficient Funds."
                                Return strReturn
                                Exit Function
                            End If

                            objCorpInv1.decBalance -= decAmount
                            objCorpInv2.decBalance += decAmount

                            strReturn = "Transfer Successful."
                        Case "IndividualInvestment"
                            objIndvInv2 = New clsIndividualInvestmentAccount(strToAcctOwner, strToAcctType)
                            m_strAccountType = strToAcctType
                            objIndvInv2.decBalance = GetBalance()

                            If decAmount > objCorpInv1.decBalance Then
                                strReturn = "Insufficient Funds."
                                Return strReturn
                                Exit Function
                            End If

                            objCorpInv1.decBalance -= decAmount
                            objIndvInv2.decBalance += decAmount

                            strReturn = "Transfer Successful."
                    End Select
                Case "IndividualInvestment"
                    objIndvInv1 = New clsIndividualInvestmentAccount(strFromAcctOwner, strFromAcctType)
                    m_strAccountType = strFromAcctType
                    objIndvInv1.decBalance = GetBalance()

                    Select Case strToAcctType
                        Case "Checking"
                            objChecking2 = New clsCheckingAccount(strToAcctOwner, strToAcctType)
                            m_strAccountType = strToAcctType
                            objChecking2.decBalance = GetBalance()

                            If decAmount > objIndvInv1.decBalance Then
                                strReturn = "Insufficient Funds."
                                Return strReturn
                                Exit Function
                            End If

                            objIndvInv1.decBalance -= decAmount
                            objChecking2.decBalance += decAmount

                            strReturn = "Transfer Successful."
                        Case "CorporateInvestment"
                            objCorpInv2 = New clsCorporateInvestmentAccount(strToAcctOwner, strToAcctType)
                            m_strAccountType = strToAcctType
                            objCorpInv2.decBalance = GetBalance()

                            If decAmount > objIndvInv1.decBalance Then
                                strReturn = "Insufficient Funds."
                                Return strReturn
                                Exit Function
                            End If

                            objIndvInv1.decBalance -= decAmount
                            objCorpInv2.decBalance += decAmount

                            strReturn = "Transfer Successful."
                        Case "IndividualInvestment"
                            objIndvInv2 = New clsIndividualInvestmentAccount(strToAcctOwner, strToAcctType)
                            m_strAccountType = strToAcctType
                            objIndvInv2.decBalance = GetBalance()

                            If decAmount > objIndvInv1.decBalance Then
                                strReturn = "Insufficient Funds."
                                Return strReturn
                                Exit Function
                            End If

                            objIndvInv1.decBalance -= decAmount
                            objIndvInv2.decBalance += decAmount

                            strReturn = "Transfer Successful."
                    End Select
            End Select
        Catch ex As Exception
            strReturn = "Error during Transfer. Transfer unsuccessful."
        Finally
            objChecking1 = Nothing
            objChecking2 = Nothing
            objCorpInv1 = Nothing
            objCorpInv2 = Nothing
            objIndvInv1 = Nothing
            objIndvInv2 = Nothing
        End Try

        Return strReturn

    End Function

    Private Function GetBalance() As Decimal

        Dim decReturnBalance As Decimal = 0.0

        Select Case m_strAccountType
            Case "Checking"
                decReturnBalance = 5000.0
            Case "CorporateInvestment"
                decReturnBalance = 500000.0
            Case "IndividualInvestment"
                decReturnBalance = 50000.0
            Case Else
                decReturnBalance = 0.0
        End Select

        Return decReturnBalance

    End Function

End Class
