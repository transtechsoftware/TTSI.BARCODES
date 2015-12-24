Public Class TPCBarcode
    Inherits Barcode

    ' Individual Barcode Elements
    Private _strCompanyCode As String ' three upper-case letters
    Private _strPackageType As String ' 1 upper-case letter
    Private _strFromAddressId As String ' Integer between 0000000 and 9999999, no commas or signs
    Private _strSequenceNo As String ' 4 numeric characters that uniquely identify the movement


    'Declare the Properties that Access the Individual Barcode Elements
    Public Property CompanyCode() As String
        Get
            Return _strCompanyCode
        End Get
        Set(ByVal Value As String)
            ValidateCompanyCode(Value)
        End Set
    End Property

    Public Property PackageType() As String
        Get
            Return _strPackageType
        End Get
        Set(ByVal Value As String)
            ValidatePackageType(Value)
        End Set
    End Property

    Public Property FromAddressId() As String
        Get
            Return _strFromAddressId
        End Get
        Set(ByVal Value As String)
            ValidateFromAddressId(Value)
        End Set
    End Property

    Public Property SequenceNo() As String
        Get
            Return _strSequenceNo
        End Get
        Set(ByVal Value As String)
            ValidateSequenceNo(Value)
        End Set
    End Property


    'Declare the Functions that Validate the Individual Barcode Elements
    Private Function ValidateCompanyCode(ByVal p_strCompanyCode As String) As Boolean

        Dim bRetVal As Boolean

        If p_strCompanyCode.Equals("TPC") Or p_strCompanyCode.Equals("CFC") Or p_strCompanyCode.Equals("TTI") Then
            _strCompanyCode = p_strCompanyCode
            bRetVal = True
        Else
            _strCompanyCode = "000"
            SetError("Invalid Company Code: " & p_strCompanyCode)
            bRetVal = False
        End If

        Return bRetVal

    End Function

    Private Function ValidatePackageType(ByVal p_strPackageType As String) As Boolean

        Dim bRetVal As Boolean

        ' Currently supported Service Codes (based on "Simplified Logic - Barcode Labels.doc")
        '   S - Single Package Item
        '   F - Pouch with No Traceable Items
        '   C - Container with Traceable Items
        '   X - Pouch with Traceable Items
        '   G - Generic Barcode.  Unique System Wide.

        If p_strPackageType.Equals("S") Or p_strPackageType.Equals("F") Or p_strPackageType.Equals("C") Or p_strPackageType.Equals("X") Or p_strPackageType.Equals("G") Then
            _strPackageType = p_strPackageType
            bRetVal = True
        Else
            _strPackageType = "0"
            SetError("Invalid Package Type: " & p_strPackageType)
            bRetVal = False
        End If

        Return bRetVal

    End Function

    Private Function ValidateFromAddressId(ByVal p_strFromAddressId As String) As Boolean

        Dim bRetVal As Boolean

        Try

            ' Make sure it is all numeric
            Dim iFromAddressId As Integer
            iFromAddressId = CInt(p_strFromAddressId)

            ' Make sure the number is within acceptable bounds
            If (iFromAddressId < 0) Or (iFromAddressId > 9999999) Then
                _strFromAddressId = "0000000"
                bRetVal = False
                SetError("Address Id is Out-of-Range: " & p_strFromAddressId)
            End If

            ' TO DO:    Add code to check if address id is valid and active.  Either by querying the Unison
            '           database or by checking a local file or db.
            _strFromAddressId = p_strFromAddressId
            bRetVal = True

        Catch ex As Exception

            bRetVal = False
            _strFromAddressId = "0000000"
            SetError(ex.Message)

        End Try

        Return bRetVal

    End Function

    Private Function ValidateSequenceNo(ByVal p_strSeqNo As String) As Boolean

        Dim bRetVal As Boolean

        Try

            ' Make sure it is all numeric
            Dim iSeqNo As Integer
            iSeqNo = CInt(p_strSeqNo)

            ' Make sure the number is within acceptable bounds
            If (iSeqNo < 0) Or (iSeqNo > 9999) Then

                bRetVal = False
                _strSequenceNo = "0000"
                SetError("Sequence Number is Out-of-Range: " & p_strSeqNo)

            End If

            _strSequenceNo = p_strSeqNo
            bRetVal = True

        Catch ex As Exception

            bRetVal = False
            _strSequenceNo = "0000"
            SetError(ex.Message)

        End Try

        Return bRetVal

    End Function


    Private Function ValidLength(ByVal p_iLength As Integer) As Boolean
        If p_iLength = 15 Then
            Return True
        Else
            Return False
        End If
    End Function

    'Implement the Overridable Methods
    Protected Overrides Function CompleteBarcode() As String

        Dim strBarcode As String

        strBarcode = _strCompanyCode & _strPackageType & _strFromAddressId & _strSequenceNo

        Return strBarcode

    End Function

    Protected Overrides Sub DecomposeBarcode(ByVal p_strBarcode As String)

        If ValidLength(p_strBarcode.Length) Then

            If Not ValidateCompanyCode(p_strBarcode.Substring(0, 3)) Then Exit Sub

            If Not ValidatePackageType(p_strBarcode.Substring(3, 1)) Then Exit Sub

            If Not ValidateFromAddressId(p_strBarcode.Substring(4, 7)) Then Exit Sub

            If Not ValidateSequenceNo(p_strBarcode.Substring(11, 4)) Then Exit Sub

        End If

    End Sub

    Public Sub New()

        MyBase.New()

        BarcodeFormat = BarcodeFactory.BarcodeFormat.TPC_Tracking

        _strCompanyCode = "000"
        _strPackageType = "0"
        _strFromAddressId = "0000000"
        _strSequenceNo = "0000"

    End Sub

    Public Sub New(ByVal p_strBarcode As String)

        MyBase.New(p_strBarcode)

        BarcodeFormat = BarcodeFactory.BarcodeFormat.TPC_Tracking

        MyClass.DecomposeBarcode(p_strBarcode)

    End Sub

End Class
