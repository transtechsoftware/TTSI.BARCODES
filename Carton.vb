Imports System.Text

Public Class Carton

    Private _sCode As String
    Private _dLength As Decimal
    Private _dWidth As Decimal
    Private _dHeight As Decimal

    'Properties
    Public Property Code() As String
        Get
            Return _sCode
        End Get
        Set(ByVal Value As String)
            _sCode = Value
        End Set
    End Property

    Public Property Length() As Decimal
        Get
            Return _dLength
        End Get
        Set(ByVal Value As Decimal)
            If Value < 0 Then Throw New ArgumentException("Length", "Length cannot be negative") Else _dLength = Value
        End Set
    End Property

    Public Property Width() As Decimal
        Get
            Return _dWidth
        End Get
        Set(ByVal Value As Decimal)
            If Value < 0 Then Throw New ArgumentException("Width", "Width cannot be negative") Else _dWidth = Value
        End Set
    End Property

    Public Property Height() As Decimal
        Get
            Return _dWidth
        End Get
        Set(ByVal Value As Decimal)
            If Value < 0 Then Throw New ArgumentException("Height", "Height cannot be negative") Else _dHeight = Value
        End Set
    End Property

    Public ReadOnly Property CubicValue() As Decimal
        Get
            Return _dLength * _dWidth * _dHeight
        End Get
    End Property

    Public ReadOnly Property DimensionString() As String
        Get
            Dim s As New StringBuilder

            s.Append(_dLength.ToString())
            s.Append("*")
            s.Append(_dWidth.ToString())
            s.Append("*")
            s.Append(_dHeight.ToString())

            Return s.ToString()
        End Get
    End Property

    Public Sub New(ByVal p_dLength As Decimal, ByVal p_dWidth As Decimal, ByVal p_dHeight As Decimal, ByVal p_sCartonCode As String)
        Length = p_dLength
        Width = p_dWidth
        Height = p_dHeight
        Code = p_sCartonCode
    End Sub

    Public Sub New(ByVal p_sStringOrCode As String)
        Dim sValues As String() = p_sStringOrCode.Split(New Char() {"*"c})
        If sValues.Length = 3 Then
            _dLength = CDec(sValues(0))
            _dWidth = CDec(sValues(1))
            _dHeight = CDec(sValues(2))
            _sCode = "CU"
        Else
            Dim c As Carton = IngramMicroCartonCodes.GetCartonByCode(p_sStringOrCode)
            If Not c Is Nothing Then
                _dLength = c.Length
                _dWidth = c.Width
                _dHeight = c.Height
                _sCode = c.Code
            End If
        End If
    End Sub

End Class
