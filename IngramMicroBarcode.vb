Public Class IngramMicroBarcode
    Inherits Barcode

    ' Individual Barcode Elements
    Private _sFormatCode As String ' IM
    Private _sTrNum As String ' Integer between 0000000000 and 9999999999 which refers to the unique, system-wide address id

    'Declare the Properties that Access the Individual Barcode Elements
    Public Property FormatCode() As String
        Get
            Return _sFormatCode
        End Get
        Set(ByVal Value As String)
            ValidateFormatCode(Value)
        End Set
    End Property

    Public Property TrackingValue() As String
        Get
            Return _sTrNum
        End Get
        Set(ByVal Value As String)
            ValidateTrNum(Value)
        End Set
    End Property

    'Declare the Functions that Validate the Individual Barcode Elements
    Private Function ValidateFormatCode(ByVal p_sFormatCode As String) As Boolean
        Dim bRetVal As Boolean
        If p_sFormatCode.Equals("IM") Then
            _sFormatCode = p_sFormatCode
            bRetVal = True
        Else
            _sFormatCode = "00"
            SetError("Invalid Format Code: " & p_sFormatCode)
            bRetVal = False
        End If
        Return bRetVal
    End Function

    Private Function ValidateTrNum(ByVal p_strTrNum As String) As Boolean
        _sTrNum = p_strTrNum
        Return True
    End Function

    Private Function ValidLength(ByVal p_iLength As Integer) As Boolean
        If p_iLength > 2 Then
            Return True
        Else
            Return False
        End If
    End Function

    'Implement the Overridable Methods
    Protected Overrides Function CompleteBarcode() As String

        Dim strBarcode As String

        strBarcode = _sFormatCode & _sTrNum

        Return strBarcode

    End Function

    Protected Overrides Sub DecomposeBarcode(ByVal p_strBarcode As String)

        If ValidLength(p_strBarcode.Length) Then

            If Not ValidateFormatCode(p_strBarcode.Substring(0, 2)) Then Exit Sub

            If Not ValidateTrNum(p_strBarcode.Substring(2, p_strBarcode.Length - 2)) Then Exit Sub

        End If

    End Sub

    Public Sub New()

        MyBase.New()

        BarcodeFormat = BarcodeFactory.BarcodeFormat.TPC_Point

        _sFormatCode = "00"
        _sTrNum = "0000000000"

    End Sub

    Public Sub New(ByVal p_strBarcode As String)

        MyBase.New(p_strBarcode)

        BarcodeFormat = BarcodeFactory.BarcodeFormat.IngramMicro

        MyClass.DecomposeBarcode(p_strBarcode)

    End Sub

End Class
