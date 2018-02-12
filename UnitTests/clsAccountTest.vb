Imports System

Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports CodingExercise



'''<summary>
'''This is a test class for clsAccountTest and is intended
'''to contain all clsAccountTest Unit Tests
'''</summary>
<TestClass()> _
Public Class clsAccountTest


    Private testContextInstance As TestContext

    '''<summary>
    '''Gets or sets the test context which provides
    '''information about and functionality for the current test run.
    '''</summary>
    Public Property TestContext() As TestContext
        Get
            Return testContextInstance
        End Get
        Set(value As TestContext)
            testContextInstance = Value
        End Set
    End Property

#Region "Additional test attributes"
    '
    'You can use the following additional attributes as you write your tests:
    '
    'Use ClassInitialize to run code before running the first test in the class
    '<ClassInitialize()>  _
    'Public Shared Sub MyClassInitialize(ByVal testContext As TestContext)
    'End Sub
    '
    'Use ClassCleanup to run code after all tests in a class have run
    '<ClassCleanup()>  _
    'Public Shared Sub MyClassCleanup()
    'End Sub
    '
    'Use TestInitialize to run code before running each test
    '<TestInitialize()>  _
    'Public Sub MyTestInitialize()
    'End Sub
    '
    'Use TestCleanup to run code after each test has run
    '<TestCleanup()>  _
    'Public Sub MyTestCleanup()
    'End Sub
    '
#End Region


    '''<summary>
    '''A test for Deposit
    '''</summary>
    <TestMethod()> _
    Public Sub DepositTest()
        Dim intAcctNum As Integer = 0 ' TODO: Initialize to an appropriate value
        Dim target As clsAccount = New clsAccount("John Jones", "Checking")
        Dim decAmount As [Decimal] = New [Decimal](500.0)
        Dim expected As String = "Successfully deposited."
        Dim actual As String
        actual = target.Deposit(500.0, "John Jones", "Checking")
        Assert.AreEqual(expected, actual)
        Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub

    '''<summary>
    '''A test for Transfer
    '''</summary>
    <TestMethod()> _
    Public Sub TransferTest()
        Dim intAcctNum As Integer = 0 ' TODO: Initialize to an appropriate value
        Dim target As clsAccount = New clsAccount("Sarah Smith", "Checking")
        Dim decAmount As [Decimal] = New [Decimal](1000.0)
        Dim intToAcctNum As Integer = 0 ' TODO: Initialize to an appropriate value
        Dim expected As String = "Transfer Successful."
        Dim actual As String
        actual = target.Transfer(100.0, "Sarah Smith", "Checking", "Sarah Smith", "IndividualInvestment")
        Assert.AreEqual(expected, actual)
        Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub

    '''<summary>
    '''A test for Withdrawl
    '''</summary>
    <TestMethod()> _
    Public Sub WithdrawlTest()
        Dim intAcctNum As Integer = 0 ' TODO: Initialize to an appropriate value
        Dim target As clsAccount = New clsAccount("MegaCorpInvestments", "CorporateInvestment")
        Dim decAmount As [Decimal] = New [Decimal]() ' TODO: Initialize to an appropriate value
        Dim expected As String = "Successfully withdrawn."
        Dim actual As String
        actual = target.Withdrawl(2000.0, "MegaCorpInvestments", "CorporateInvestment")
        Assert.AreEqual(expected, actual)
        Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub
End Class
