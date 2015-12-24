'   LOG:
'
'   11-04-2010 (Sammy Nava)
'   
'   This class was modified in order to handle an expanded number of record columns; HHid
'
'   08-20-2010 (Sammy Nava)
'   
'   This class was modifified in order to handle an expanded number of record columns; BatchDate & ErrorLog
'
'   08-08-2008 (Sammy Nava)
'
'   This class implements a scan record that tries to combine as many of the elements of the current
'   files that Unison deals with.  We currently import 4 different types of files, 1) Files generated by 
'   our handheld scanners, 2) Files generated by our TopPriorityClient application, 3) Files generated by
'   IPI and 4) EDI based files generated by IBC.  This ScanRecord is going to gather and/or generate as
'   many fields as it can using only a scanned barcode and weight.  The goal is to make this a general
'   purpose record that can scan any barcoded package for later processing by Unison.
'
'   COMMENTS AND NOTES:
'
'   These are the most commonly used Date/Time format characters.  Move this to a more appropriate module
'
'   d - Numeric day of the month without a leading zero.
'   dd - Numeric day of the month with a leading zero.
'   ddd - Abbreviated name of the day of the week.
'   dddd - Full name of the day of the week.

'   f,ff,fff,ffff,fffff,ffffff,fffffff - Fraction of a second. The more F's the higher the presision.

'   h - 12 Hour clock, no leading zero.
'   hh - 12 Hour clock with leading zero.
'   H - 24 Hour clock, no leading zero.
'   HH - 24 Hour clock with leading zero.

'   m - Minutes with no leading zero.
'   mm - Minutes with leading zero.

'   M - Numeric month with no leading zero.
'   MM - Numeric month with a leading zero.
'   MMM - Abbreviated name of month.
'   MMMM - Full month name.

'   s - Seconds with no leading zero.
'   ss - Seconds with leading zero.

'   t - AM/PM but only the first letter.
'   tt - AM/PM

'   y - Year with out century and leading zero.
'   yy - Year with out century, with leading zero.
'   yyyy - Year with century.

'   zz - Time zone off set with +/-.
'

Public Class ScanRecord

    Public Enum Fields      ' Index Order of Fields:
        EventCode = 0       '   0   Maps to EVENT.EventCode 
        OperatorId = 1      '   1   Maps to EVENT.OperatorID
        PointId = 2         '   2   Maps to EVENT.PointID
        Barcode = 3         '   3   Maps to EVENT.TrackingNum or EVENT.ThirdPartyBarcode
        Weight = 4          '   4   Maps to EVENT.Weight
        X = 5               '   5   Maps to EVENT.Pieces with modification (1/1, 1/2 2/2 etc)
        ScanError = 6       '   6   0 = No Error, 1 = Duplicate, 2 = Tag Not Turned, 3 = Test Scan, 4 = No Weight, 5 = Scale Error, 6 = Unrecognized Barcode
        '                           EVENT.Void = 'T', EVENT.VoidMsg = 'SCAN ERROR: ...'
        BatchId = 7         '   7   Maps to EVENT.BatchNum
        '                           Format = BBXX: BB = 2 hex-digit digit Branch ID, unique system-wide; 2 hex-digit counter, unique within file.
        DateStamp = 8       '   8   Maps to EVENT.ScanDate
        BatchDate = 9       '   9   Maps to EVENT.BatchDate
        ErrorLog = 10       '   10  Maps to EVENT.ErrorLog
        HHid = 11           '   11  Maps to EVENT.HHid
        ProcessDate = 12    '   90  Is not currently part of a scan record
        Processed = 13      '   91  Is not currently part of a scan record
        'Add Support for Route Name in Delivery Comments for BD.  Route name currently carrie in X field
        DeliveryOption = 14 '   14  Maps to EVENT.DeliveryOption
        DeliveryComments = 15 '  15  Maps to EVENT.DelieryComments
    End Enum

    ' Members necessary for record parsing.  These variables will hold the current record being written or read.
    Private _chrDelimiter As String
    Private _dTimeStamp As Date

    Private _strEventCode As String
    Private _oOperatorId As Barcode
    Private _oPointId As Barcode
    Private _clsBarcode As Barcode
    Private _fWeight As Decimal
    Private _iX As Integer '1 by default.  Increments by 1 for each identical barcode that is not a duplicate within the same file.
    Private _iScanError As Integer 'Marks scan record void because of numerically defined error.  0 is no-error condition.
    Private _iBatchId As Integer 'Tracking Numbers are scanned in batches.  A new batch is started each time an Operator logs in.
    Private _dProcessDate As Date
    Private _iProcessed As Integer ' 0 if not read by unison, 1 if it has
    Private _dBatchDate As Date
    Private _lErrorLog As Long
    Private _sHHid As String
    Private _iDeliveryOption As Short
    Private _sDeliveryComments As String 'nvarchar(50)

    ' Members that hold statistical data about the currently loaded record
    Private _strRecordString As String
    Private _iExpectedFieldCount As Integer
    Private _strErrorMessage As String
    Private _bHasError As Boolean

    Public Sub New()
        _chrDelimiter = "|"
        _strEventCode = "WC"
        _fWeight = 0D
        _iX = 1
        _iScanError = 0
        _iBatchId = 0
        _iProcessed = 0
        _iExpectedFieldCount = 12
        _strErrorMessage = ""
        _sHHid = "01W0"
        _bHasError = False
        _iDeliveryOption = 0
        _sDeliveryComments = String.Empty
    End Sub

    Protected Sub SetRecordString(ByVal p_sValue As String)
        _strRecordString = p_sValue
    End Sub

    Protected Property ExpectedFieldCount() As Integer
        Get
            Return _iExpectedFieldCount
        End Get
        Set(ByVal Value As Integer)
            _iExpectedFieldCount = Value
        End Set
    End Property

    ' Collection of ScanList Records
    Public ReadOnly Property HasError() As Boolean
        Get
            Return _bHasError
        End Get
    End Property

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return _strErrorMessage
        End Get
    End Property

    Protected Sub SetError(ByVal p_errMsg As String)

        _bHasError = True
        _strErrorMessage = p_errMsg

    End Sub

    Public Property Delimiter() As String
        Get
            Return _chrDelimiter
        End Get
        Set(ByVal Value As String)
            _chrDelimiter = Value
        End Set
    End Property

    Public Property EventCode() As String
        Get
            Return _strEventCode
        End Get
        Set(ByVal Value As String)
            AssignEventCode(Value)
        End Set
    End Property

    Public Property HHid() As String
        Get
            Return _sHHid
        End Get
        Set(ByVal Value As String)
            AssignHHid(Value)
        End Set
    End Property

    Public Property OperatorId() As String
        Get
            Return _oOperatorId.ToString()
        End Get
        Set(ByVal Value As String)
            AssignOperatorId(Value)
        End Set
    End Property

    Public Property PointId() As String
        Get
            Return _oPointId.ToString()
        End Get
        Set(ByVal Value As String)
            AssignPointId(Value)
        End Set
    End Property

    Public Property Barcode() As String
        Get
            If Not IsNothing(_clsBarcode) Then
                Return _clsBarcode.ToString()
            Else
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            AssignBarcode(Value)
        End Set
    End Property

    Public ReadOnly Property BarcodeName() As String
        Get
            If Not IsNothing(_clsBarcode) Then
                Return _clsBarcode.BarcodeName
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property BarcodeFormat() As BarcodeFormat
        Get
            If Not IsNothing(_clsBarcode) Then
                Return _clsBarcode.BarcodeFormat
            Else
                Return BarcodeFactory.BarcodeFormat.Unknown
            End If
        End Get
    End Property

    Public Property Weight() As Decimal
        Get
            Return _fWeight
        End Get
        Set(ByVal Value As Decimal)
            AssignWeight(Value)
        End Set
    End Property

    Public ReadOnly Property TimeStamp() As Date
        Get
            Return _dTimeStamp
        End Get
    End Property

    Public Property ProcessDate() As Date
        Get
            Return _dProcessDate
        End Get
        Set(ByVal Value As Date)
            _dProcessDate = Value
        End Set
    End Property

    Public Property Processed() As Integer
        Get
            Return _iProcessed
        End Get
        Set(ByVal Value As Integer)
            AssignProcessed(Value)
        End Set
    End Property

    Public Property X() As Integer
        Get
            Return _iX
        End Get
        Set(ByVal Value As Integer)
            _iX = Value
        End Set
    End Property

    Public Property ScanError() As Integer
        Get
            Return _iScanError
        End Get
        Set(ByVal Value As Integer)
            _iScanError = Value
        End Set
    End Property

    Public Property BatchID() As Integer
        Get
            Return _iBatchId
        End Get
        Set(ByVal Value As Integer)
            _iBatchId = Value
        End Set
    End Property

    Public Property BatchDate() As Date
        Get
            Return _dBatchDate
        End Get
        Set(ByVal Value As Date)
            _dBatchDate = Value
        End Set
    End Property

    Public Property DeliveryOption() As Short
        Get
            Return _iDeliveryOption
        End Get
        Set(ByVal Value As Short)
            _iDeliveryOption = Value
        End Set
    End Property

    Public Property DeliveryComments() As String
        Get
            Return _sDeliveryComments
        End Get
        Set(ByVal Value As String)
            _sDeliveryComments = Value
        End Set
    End Property

    Public Property ErrorLog() As Long
        Get
            Return _lErrorLog
        End Get
        Set(ByVal Value As Long)
            _lErrorLog = Value
        End Set
    End Property

    Public ReadOnly Property RecordString() As String

        Get

            Dim strBarcode As String = ""
            Dim strOperatorId As String = ""
            Dim strPointId As String = ""

            If Not IsNothing(_clsBarcode) Then strBarcode = _clsBarcode.ToString()
            If Not IsNothing(_oOperatorId) Then strOperatorId = _oOperatorId.ToString()
            If Not IsNothing(_oPointId) Then strPointId = _oPointId.ToString()

            _strRecordString = _
                _strEventCode & _chrDelimiter & _
                strOperatorId & _chrDelimiter & _
                strPointId & _chrDelimiter & _
                strBarcode & _chrDelimiter & _
                _fWeight.ToString() & _chrDelimiter & _
                _iX & _chrDelimiter & _
                _iScanError & _chrDelimiter & _
                _iBatchId & _chrDelimiter & _
                Format(_dTimeStamp, "MMddyyyyHHmmss") & _chrDelimiter & _
                Format(_dBatchDate, "MMddyyyyHHmmss") & _chrDelimiter & _
                _lErrorLog & _chrDelimiter & _
                _sHHid
            'Format(_dProcessDate, "MMddyyyy") & _chrDelimiter & _
            '_iProcessed

            Return _strRecordString

        End Get

    End Property

    Protected Function DecomposeRecordString(ByVal p_strSubStrings() As String) As Boolean

        Try

            ' Validate and Assign EventCode
            If Not AssignEventCode(p_strSubStrings(Fields.EventCode)) Then Return False

            ' Validate and Assign OperatorId
            If Not AssignOperatorId(p_strSubStrings(Fields.OperatorId)) Then Return False

            ' Validate and Assign PointId
            If Not AssignPointId(p_strSubStrings(Fields.PointId)) Then Return False

            ' Validate and Assign Barcode
            If Not AssignBarcode(p_strSubStrings(Fields.Barcode)) Then Return False

            ' Validate and Assign Weight
            Dim fWeight As Decimal
            fWeight = CDec(p_strSubStrings(Fields.Weight))
            If Not AssignWeight(fWeight) Then Return False

            ' Validate and Assign X
            If p_strSubStrings(Fields.EventCode).CompareTo("BD") = 0 Then
                DeliveryOption = 5
                DeliveryComments = p_strSubStrings(Fields.X)
                AssignX(1)
            Else
                If Not AssignX(p_strSubStrings(Fields.X)) Then Return False
            End If

            ' Validate and Assign Y
            If Not AssignScanError(p_strSubStrings(Fields.ScanError)) Then Return False

            ' Validate and Assign BatchId
            If Not AssignBatchId(p_strSubStrings(Fields.BatchId)) Then Return False

            ' Assign Time Stamp
            _dTimeStamp = Date.ParseExact(p_strSubStrings(Fields.DateStamp), "MMddyyyyHHmmss", New Globalization.CultureInfo("en-US"))

            ' Assign Batch Date
            _dBatchDate = Date.ParseExact(p_strSubStrings(Fields.BatchDate), "MMddyyyyHHmmss", New Globalization.CultureInfo("en-US"))

            ' Assign Error Log
            If Not AssignErrorLog(p_strSubStrings(Fields.ErrorLog)) Then Return False

            ' Assign HHid
            If Not AssignHHid(p_strSubStrings(Fields.HHid)) Then Return False

            ' Assign Process Date
            '_dProcessDate = Date.ParseExact(p_strSubStrings(Fields.ProcessDate), "MMddyyyy", New Globalization.CultureInfo("en_US"))

            ' Assign Processed Flag's Value
            'If Not AssignProcessed(p_strSubStrings(Fields.Processed)) Then Return False

        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges

            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function

    Private Function AssignEventCode(ByVal p_sEventCode As String) As Boolean

        Try

            Dim s As String

            s = p_sEventCode

            If s.Equals("DD") Then
                _strEventCode = s
                Return True 'Driver Delivery
            End If

            If s.Equals("BD") Then
                _strEventCode = s
                Return True 'Branch Deliver
            End If

            If s.Equals("TR") Then
                _strEventCode = s
                Return True 'Transit
            End If

            If s.Equals("PK") Then
                _strEventCode = s
                Return True 'Pack
            End If

            If s.Equals("UP") Then
                _strEventCode = s
                Return True 'Un-pack
            End If

            If s.Equals("WC") Then
                _strEventCode = s
                Return True 'Weight Capture
            End If

            If s.Equals("MT") Then
                _strEventCode = s
                Return True 'Manifest
            End If

            If s.Equals("EZ") Then
                _strEventCode = s
                Return True 'EZ-Scan Delivery Event
            End If

            If s.Equals("AF") Then
                _strEventCode = s
                Return True 'EZ-Scan Pickup Event
            End If

            SetError("Unknown Event Code")

            Return False

        Catch ex As Exception

            SetError(ex.Message & " [ScanRecord::AssignEventCode]")
            Return False

        End Try

    End Function

    Private Function AssignBarcode(ByVal p_strBarcode As String) As Boolean

        Try

            Dim bRetVal As Boolean

            '_clsBarcode.Barcode = p_strBarcode
            _clsBarcode = Nothing
            _clsBarcode = NewBarcodeObject(p_strBarcode)

            If Not IsNothing(_clsBarcode) Then
                If _clsBarcode.HasError Then
                    SetError(_clsBarcode.ErrorMessage)
                    bRetVal = False
                Else
                    bRetVal = True
                End If
            Else
                SetError("Unknown Barcode Type")
                bRetVal = False
            End If

            Return bRetVal

        Catch ex As Exception

            SetError(ex.Message & " [ScanRecord::AssignBarcode]")

        End Try

    End Function

    Private Function AssignOperatorId(ByVal p_strOperatorId As String) As Boolean

        Try


            Dim bRetVal As Boolean

            _oOperatorId = Nothing
            _oOperatorId = NewBarcodeObject(p_strOperatorId)

            If Not IsNothing(_oOperatorId) Then
                If _oOperatorId.HasError Then
                    SetError(_oOperatorId.ErrorMessage)
                    bRetVal = False
                Else
                    bRetVal = True
                End If
            Else
                SetError("Unknown Barcode Type")
                bRetVal = False
            End If

            Return bRetVal

        Catch ex As Exception

            SetError(ex.Message & " [ScanRecord::AssignOperatorId]")
            Return False

        End Try


    End Function

    Private Function AssignPointId(ByVal p_strPointId As String) As Boolean

        Try

            Dim bRetVal As Boolean

            _oPointId = Nothing
            _oPointId = NewBarcodeObject(p_strPointId)

            If Not IsNothing(_oPointId) Then
                If _oPointId.HasError Then
                    SetError(_oPointId.ErrorMessage)
                    bRetVal = False
                Else
                    bRetVal = True
                End If
            Else
                SetError("Unknown Barcode Type")
                bRetVal = False
            End If

            Return bRetVal

        Catch ex As Exception

            SetError(ex.Message & " [ScanRecord::AssignPointId]")
            Return False

        End Try

    End Function

    Private Function AssignWeight(ByVal Value As Decimal) As Boolean

        Try

            Dim bRetVal As Boolean = True

            'Temporarily Allow negative weight
            _fWeight = Value
            'If Value >= 0 Then
            '    _fWeight = Value
            'Else
            '    SetError("Invalid Weight")
            '    _fWeight = 0.0
            '    bRetVal = False
            'End If

            Return bRetVal

        Catch ex As Exception

            SetError(ex.Message & " [ScanRecord::AssignWeight]")
            Return False

        End Try


    End Function

    Private Function AssignX(ByVal Value As Integer) As Boolean

        Try

            Dim bRetVal As Boolean = True

            If Value > 0 Then
                _iX = Value
            Else
                SetError("Invalid X")
                _iX = 0
                bRetVal = False
            End If

            Return bRetVal

        Catch ex As Exception

            SetError(ex.Message & " [ScanRecord::AssignX]")
            Return False

        End Try


    End Function

    Private Function AssignProcessed(ByVal Value As Integer) As Boolean

        Try

            Dim bRetVal As Boolean = True

            If Value = 0 Or Value = 1 Then

                _iProcessed = Value

            Else

                SetError("Invalid Processed Flag")
                _iProcessed = 0
                bRetVal = False

            End If

            Return bRetVal

        Catch ex As Exception

            SetError(ex.Message & " [ScanRecord::Processed]")
            Return False

        End Try

    End Function

    Private Function AssignScanError(ByVal Value As Integer) As Boolean

        Try

            Dim bRetVal As Boolean = True

            If Value >= 0 Then
                _iScanError = Value
            Else
                SetError("Invalid Scan Error")
                _iScanError = 0
                bRetVal = False
            End If

            Return bRetVal

        Catch ex As Exception

            SetError(ex.Message & "ScanRecord::AssignScanError")
            Return False

        End Try

    End Function

    Private Function AssignBatchId(ByVal Value As Integer) As Boolean
        Try

            Dim bRetVal As Boolean = True

            If Value >= 0 Then

                _iBatchId = Value

            Else

                SetError("Invalid BatchId")
                _iBatchId = 0
                bRetVal = False

            End If

            Return bRetVal

        Catch ex As Exception

            SetError(ex.Message & "ScanRecord::AssignBatchId")
            Return False

        End Try

    End Function

    Private Function AssignErrorLog(ByVal Value As Long) As Boolean

        Dim bReturn As Boolean = True

        Try

            If Value >= 0L Then

                _lErrorLog = Value

            Else

                SetError("Invalid Error Log")
                _lErrorLog = 3L
                bReturn = False

            End If

        Catch ex As Exception

            SetError(ex.Message & "ScanRecord::AssignErrorLog")
            bReturn = False

        End Try

        Return bReturn

    End Function

    Private Function AssignHHid(ByVal p_sHHid As String) As Boolean

        Try

            If p_sHHid.Length = 4 Then
                _sHHid = p_sHHid
                Return True
            End If

            SetError("Improperly Formatted Station ID")

            Return False

        Catch ex As Exception

            SetError(ex.Message & " [ScanRecord::AssignHHid]")
            Return False

        End Try

    End Function

    Protected Sub Reset()

        _clsBarcode = Nothing
        _fWeight = 0D
        _iProcessed = 0
        _dTimeStamp = Date.MinValue
        _lErrorLog = 3L
        _sHHid = Nothing

    End Sub

    Public Overridable Function AssignRecord(ByVal p_strLine As String) As Boolean

        Try

            _strRecordString = p_strLine

            ' Check #1 - Count number of fields in record
            Dim str() As String
            Dim iFieldCount As Integer

            str = p_strLine.Split(_chrDelimiter)
            iFieldCount = str.Length

            If iFieldCount = _iExpectedFieldCount Then
                Return DecomposeRecordString(str)
            Else
                SetError("Unexpected Number of Fields")
                Reset()
                Return False
            End If

        Catch ex As Exception

            SetError(ex.Message & " [ScanRecord::AssignRecord]")
            Return False

        End Try


    End Function

    Public Sub Stamp()

        _dTimeStamp = Date.Now

    End Sub

End Class
