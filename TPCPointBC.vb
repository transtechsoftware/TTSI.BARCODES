Public Class TPCPointBC
    Inherits Barcode

    ' Individual Barcode Elements
    Private _strFormatCode As String ' one upper-case letters
    Private _strAddressId As String ' Integer between 0000000 and 9999999 which refers to the unique, system-wide address id

    'Declare the Properties that Access the Individual Barcode Elements
    Public Property FormatCode() As String
        Get
            Return _strFormatCode
        End Get
        Set(ByVal Value As String)
            ValidateFormatCode(Value)
        End Set
    End Property

    Public Property AddressId() As String
        Get
            Return _strAddressId
        End Get
        Set(ByVal Value As String)
            ValidateAddressId(Value)
        End Set
    End Property

    'Declare the Functions that Validate the Individual Barcode Elements
    Private Function ValidateFormatCode(ByVal p_strFormatCode As String) As Boolean
        Dim bRetVal As Boolean
        If p_strFormatCode.Equals("P") Then
            _strFormatCode = p_strFormatCode
            bRetVal = True
        Else
            _strFormatCode = "0"
            SetError("Invalid Format Code: " & p_strFormatCode)
            bRetVal = False
        End If
        Return bRetVal
    End Function

    Private Function ValidateAddressId(ByVal p_strAddressId As String) As Boolean
        Dim bRetVal As Boolean
        Try
            'Make sure it is all numeric
            Dim iEmployeeId As Integer
            iEmployeeId = CInt(p_strAddressId)

            'Make sure the number is within acceptable bounds
            If (iEmployeeId < 0) Or (iEmployeeId > 9999999) Then
                _strAddressId = "0000000"
                bRetVal = False
                SetError("Address Id is Out-of-Range: " & p_strAddressId)
            End If
            'TO DO: Add code to check if employee is valid and active.  Either by querying the Unison database or by checking a local file or db
            _strAddressId = p_strAddressId
            bRetVal = True
        Catch ex As Exception
            bRetVal = False
            _strAddressId = "0000000"
            SetError(ex.Message)
        End Try
        Return bRetVal
    End Function

    Private Function ValidLength(ByVal p_iLength As Integer) As Boolean
        If p_iLength = 8 Then
            Return True
        Else
            Return False
        End If
    End Function

    'Implement the Overridable Methods
    Protected Overrides Function CompleteBarcode() As String

        Dim strBarcode As String

        strBarcode = _strFormatCode & _strAddressId

        Return strBarcode

    End Function

    Protected Overrides Sub DecomposeBarcode(ByVal p_strBarcode As String)

        If ValidLength(p_strBarcode.Length) Then

            If Not ValidateFormatCode(p_strBarcode.Substring(0, 1)) Then Exit Sub

            If Not ValidateAddressId(p_strBarcode.Substring(1, 7)) Then Exit Sub

        End If

    End Sub

    Public Sub New()

        MyBase.New()

        BarcodeFormat = BarcodeFactory.BarcodeFormat.TPC_Point

        _strFormatCode = "0"
        _strAddressId = "0000000"

    End Sub

    Public Sub New(ByVal p_strBarcode As String)

        MyBase.New(p_strBarcode)

        BarcodeFormat = BarcodeFactory.BarcodeFormat.TPC_Point

        MyClass.DecomposeBarcode(p_strBarcode)

    End Sub

End Class
