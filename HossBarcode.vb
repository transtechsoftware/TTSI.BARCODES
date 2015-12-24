'This class represents a barcode formatted using the HOSS format.

Public Class HossBarcode
    Inherits Barcode

    ' Individual Barcode Elements
    Private _strDestBranchCityCode As String ' three alpha-numeric characters
    Private _strServiceType As String ' 1 alpha-numeric character
    Private _strDestLocId As String ' Integer between 000000 and 999999, no commas or signs
    Private _strSourceLocCode As String ' 6 alpha-numeric characters

    ' The barcode itself
    ''Private _strBarcode As String

    ' Other Private Members
    ''Private _bErrorCondition As Boolean = False
    ''Private _strErrorMessage As String = ""

    ''Public ReadOnly Property HasError() As Boolean
    ''Get
    ''Return _bErrorCondition
    ''End Get
    ''End Property

    ''Public ReadOnly Property ErrorMessage() As String
    ''Get
    ''Return _strErrorMessage
    ''End Get
    ''End Property

    'Declare the Properties that Access the Individual Barcode Elements
    Public Property ToCityCode() As String
        Get
            Return _strDestBranchCityCode
        End Get
        Set(ByVal Value As String)
            If ValidateCityCode(Value) = True Then
                _strDestBranchCityCode = Value
            Else
                _strDestBranchCityCode = "000"
                SetError("Invalid City Code: " & Value)
            End If
        End Set
    End Property

    Public Property ServiceType() As String
        Get
            Return _strServiceType
        End Get
        Set(ByVal Value As String)
            If ValidateServiceType(Value) Then
                _strServiceType = Value
            Else
                _strServiceType = "0"
                SetError("Invalid Service Type: " & Value)
            End If
        End Set
    End Property

    Public Property DestinationCode() As String
        Get
            Return _strDestLocId
        End Get
        Set(ByVal Value As String)
            If ValidateDestinationLocation(Value) Then
                _strDestLocId = Value
            Else
                _strDestLocId = "000000"
                SetError("Invalid Destination Location Id: " & Value)
            End If
        End Set
    End Property

    Public Property SourceCode() As String
        Get
            Return _strSourceLocCode
        End Get
        Set(ByVal Value As String)
            If ValidateSourceLocationCode(Value) Then
                _strSourceLocCode = Value
            Else
                _strSourceLocCode = "000000"
                SetError("Invalid Source Location Code")
            End If
        End Set
    End Property

    ''Public Property Barcode() As String
    ''Get
    ''Return CompleteBarcode()
    ''End Get
    ''Set(ByVal Value As String)
    ''DecomposeBarcode(Value)
    ''End Set
    ''End Property

    'Declare the Functions that Validate the Individual Barcode Elements
    Private Function ValidateCityCode(ByVal p_strCityCode As String) As Boolean

        Dim bReturnValue As Boolean = True

        'TO DO:  Add validation code here
        _strDestBranchCityCode = p_strCityCode

        Return bReturnValue

    End Function

    Private Function ValidateServiceType(ByVal p_strServiceType As String) As Boolean

        Dim bReturnValue As Boolean = True

        'TO DO:  Add validation code here
        _strServiceType = p_strServiceType

        Return bReturnValue

    End Function

    Private Function ValidateDestinationLocation(ByVal p_strDestLocId As String) As Boolean

        Dim bReturnValue As Boolean = True

        'TO DO:  Add validation code here
        _strDestLocId = p_strDestLocId

        Return bReturnValue

    End Function

    Private Function ValidateSourceLocationCode(ByVal p_strSourceCode As String) As Boolean

        Dim bReturnValue As Boolean = True

        'TO DO:  Add validation code here
        _strSourceLocCode = p_strSourceCode

        Return bReturnValue

    End Function

    Private Function ValidLength(ByVal p_iLength As Integer) As Boolean
        If p_iLength = 16 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''Private Sub SetError(ByVal p_strErrMsg As String)

    ''_bErrorCondition = True
    ''_strErrorMessage = p_strErrMsg

    ''End Sub

    'Implement the Overridable Methods
    Protected Overrides Function CompleteBarcode() As String
        Dim strBarcode As String
        strBarcode = _strDestBranchCityCode & _strServiceType & _strDestLocId & _strSourceLocCode
        Return strBarcode
    End Function

    Protected Overrides Sub DecomposeBarcode(ByVal p_strBarcode As String)

        If ValidLength(p_strBarcode.Length) Then

            If Not ValidateCityCode(p_strBarcode.Substring(0, 3)) Then Exit Sub

            If Not ValidateServiceType(p_strBarcode.Substring(3, 1)) Then Exit Sub

            If Not Me.ValidateDestinationLocation(p_strBarcode.Substring(4, 6)) Then Exit Sub

            If Not Me.ValidateSourceLocationCode(p_strBarcode.Substring(10, 6)) Then Exit Sub

        End If

    End Sub

    ''Public Overrides Function ToString() As String
    ''Return Me.Barcode
    ''End Function

    Public Sub New()

        MyBase.New()

        BarcodeFormat = BarcodeFactory.BarcodeFormat.Unity_HOSS

        _strDestBranchCityCode = "000"
        _strServiceType = "0"
        _strDestLocId = "000000"
        _strSourceLocCode = "000000"

    End Sub

    Public Sub New(ByVal p_strBarcode As String)

        MyBase.New(p_strBarcode)

        BarcodeFormat = BarcodeFactory.BarcodeFormat.Unity_HOSS

        DecomposeBarcode(p_strBarcode)

    End Sub

End Class
