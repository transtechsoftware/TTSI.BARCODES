Public Class TPCOperatorBC
    Inherits Barcode

    ' Individual Barcode Elements
    Private _strFormatCode As String ' one upper-case letters
    Private _strEmployeeId As String ' Integer between 0000000 and 9999999 which refers to the unique, system-wide employee id

    'Declare the Properties that Access the Individual Barcode Elements
    Public Property FormatCode() As String
        Get
            Return _strFormatCode
        End Get
        Set(ByVal Value As String)
            ValidateFormatCode(Value)
        End Set
    End Property

    Public Property EmployeeId() As String
        Get
            Return _strEmployeeID
        End Get
        Set(ByVal Value As String)
            ValidateEmployeeId(Value)
        End Set
    End Property

    'Declare the Functions that Validate the Individual Barcode Elements
    Private Function ValidateFormatCode(ByVal p_strFormatCode As String) As Boolean
        Dim bRetVal As Boolean
        If p_strFormatCode.Equals("E") Then
            _strFormatCode = p_strFormatCode
            bRetVal = True
        Else
            _strFormatCode = "0"
            SetError("Invalid Format Code: " & p_strFormatCode)
            bRetVal = False
        End If
        Return bRetVal
    End Function

    Private Function ValidateEmployeeId(ByVal p_strEmployeeId As String) As Boolean
        Dim bRetVal As Boolean
        Try
            'Make sure it is all numeric
            Dim iEmployeeId As Integer
            iEmployeeId = CInt(p_strEmployeeId)

            'Make sure the number is within acceptable bounds
            If (iEmployeeId < 0) Or (iEmployeeId > 9999999) Then
                _strEmployeeId = "0000000"
                bRetVal = False
                SetError("Employee Id is Out-of-Range: " & p_strEmployeeId)
            End If
            'TO DO: Add code to check if employee is valid and active.  Either by querying the Unison database or by checking a local file or db
            _strEmployeeId = p_strEmployeeId
            bRetVal = True
        Catch ex As Exception
            bRetVal = False
            _strEmployeeId = "0000000"
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

        strBarcode = _strFormatCode & _strEmployeeId

        Return strBarcode

    End Function

    Protected Overrides Sub DecomposeBarcode(ByVal p_strBarcode As String)

        If ValidLength(p_strBarcode.Length) Then

            If Not ValidateFormatCode(p_strBarcode.Substring(0, 1)) Then Exit Sub

            If Not ValidateEmployeeId(p_strBarcode.Substring(1, 7)) Then Exit Sub

        End If

    End Sub

    Public Sub New()

        MyBase.New()

        BarcodeFormat = BarcodeFactory.BarcodeFormat.TPC_Operator

        _strFormatCode = "0"
        _strEmployeeId = "0000000"

    End Sub

    Public Sub New(ByVal p_strBarcode As String)

        MyBase.New(p_strBarcode)

        BarcodeFormat = BarcodeFactory.BarcodeFormat.TPC_Operator

        MyClass.DecomposeBarcode(p_strBarcode)

    End Sub


End Class
