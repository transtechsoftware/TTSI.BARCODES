Public Class Barcode

    ' Individual Barcode Components will be defined in each Child class

    ' The barcode string itself
    Private _strBarcode As String

    ' Other Private Members
    Private _bErrorCondition As Boolean
    Private _strErrorMessage As String
    Private _eFormat As BarcodeFormat

    Public ReadOnly Property BarcodeName() As String

        Get
            Select Case _eFormat
                Case BarcodeFactory.BarcodeFormat.Unity_HOSS
                    Return "HOSS"
                Case BarcodeFactory.BarcodeFormat.TPC_Tracking
                    Return "TPCTRK"
                Case BarcodeFactory.BarcodeFormat.TPC_Point
                    Return "TPCPNT"
                Case BarcodeFactory.BarcodeFormat.TPC_Operator
                    Return "TPCOPR"
                Case Else
                    Return "UNKNOWN"
            End Select
        End Get

    End Property

    Public Property BarcodeFormat() As BarcodeFormat
        Get
            Return _eFormat
        End Get
        Set(ByVal Value As BarcodeFormat)
            _eFormat = Value
        End Set
    End Property

    Public ReadOnly Property HasError() As Boolean
        Get
            Return _bErrorCondition
        End Get
    End Property

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return _strErrorMessage
        End Get
    End Property

    Protected Sub SetError(ByVal p_strErrMsg As String)
        _bErrorCondition = True
        _strErrorMessage = p_strErrMsg
    End Sub

    Protected Overridable Function CompleteBarcode() As String
        Return _strBarcode
    End Function

    Protected Overridable Sub DecomposeBarcode(ByVal p_strBarcode As String)
        _strBarcode = p_strBarcode
    End Sub

    Public Property Barcode() As String
        Get
            Return CompleteBarcode()
        End Get
        Set(ByVal Value As String)
            DecomposeBarcode(Value)
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return Me.Barcode
    End Function

    Public Sub New()

        _bErrorCondition = False
        _strErrorMessage = ""

        _eFormat = BarcodeFactory.BarcodeFormat.Unknown

        _strBarcode = ""

    End Sub

    Public Sub New(ByVal p_strBarcode As String)

        _bErrorCondition = False
        _strErrorMessage = ""

        _eFormat = BarcodeFactory.BarcodeFormat.Unknown

        MyClass.DecomposeBarcode(p_strBarcode)

    End Sub

End Class
