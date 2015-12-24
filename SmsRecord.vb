' 2013-06-17 Created
'   The Primary Purpose of the SmsRecord is to provide a placeholder for information that is going to be input into 
'   Unison's [EVENT] table based on SMS text messages received.  This class also introduces a rule that should be 
'   recognized by all future developers; when toupper(HHID) = "CELL" then BatchNum is equal to the RowID of the
'   SMS_In record from which this data was derived.
Public Class SmsRecord
    Inherits ScanRecord

    Shadows Enum Fields
        ' Common to ScanRecord
        EventCode = 0           '   0   Maps to EVENT.EventCode 
        OperatorId = 1          '   1   Maps to EVENT.OperatorID
        PointId = 2             '   2   Maps to EVENT.PointID
        Barcode = 3             '   3   Maps to EVENT.TrackingNum or EVENT.ThirdPartyBarcode
        Weight = 4              '   4   Maps to EVENT.Weight
        X = 5                   '   5   Maps to EVENT.Pieces with modification (1/1, 1/2 2/2 etc)
        ScanError = 6           '   6   0 = No Error, 1 = Duplicate, 2 = Tag Not Turned, 3 = Test Scan, 4 = No Weight, 5 = Scale Error, 6 = Unrecognized Barcode
        '                               EVENT.Void = 'T', EVENT.VoidMsg = 'SCAN ERROR: ...'
        BatchId = 7             '   7   Maps to EVENT.BatchNum
        '                               Format = BBXX: BB = 2 hex-digit digit Branch ID, unique system-wide; 2 hex-digit counter, unique within file.
        DateStamp = 8           '   8   Maps to EVENT.ScanDate
        BatchDate = 9           '   9   Maps to EVENT.BatchDate
        ErrorLog = 10           '   10  Maps to EVENT.ErrorLog
        HHid = 11               '   11  Maps to EVENT.HHid
        ProcessDate = 90        '   90  Is not currently part of a scan record
        Processed = 91          '   91  Is not currently part of a scan record
        ' New to SmsRecord
        ToAddressId = 12        '   14  Maps to EVENT.ToAdd

        'ToLocId = 15            '   15  Maps to EVENT.ToLocID
        'ToLocName = 16          '   16  Maps to EVENT.ToLocName
        'Depot = 14              '   14  delete
        'DriverFirstName = 15    '   15  delete
        'DriverLastName = 16     '   16  delete
        'DriverCell = 17         '   17  delete
        'RecipientName = 18      '   18  delete
        'ServiceLevel = 19       '   19  delete
        'ToStreet = 22           '   22  delete
        'ToCity = 23             '   23  delete
        'ToZip = 24              '   24  delete
        'ToState = 25            '   25  delete
    End Enum

    ' Members necessary for record parsing.  These variables will hold the current record being written or read.
    Private _iToAddId As Integer

    Public Sub New()

        MyBase.New()
        ExpectedFieldCount = 13
        _iToAddId = 0

    End Sub

    Public Property ToAddressId() As Integer
        Get
            Return _iToAddId
        End Get
        Set(ByVal Value As Integer)
            _iToAddId = Value
        End Set
    End Property

    Public Shadows ReadOnly Property RecordString() As String
        Get
            Return MyBase.RecordString & Delimiter & _iToAddId.ToString()
        End Get
    End Property

    Protected Shadows Function DecomposeRecordString(ByVal p_strSubStrings() As String) As Boolean

        Dim bReturn As Boolean = True

        Try

            If (bReturn = MyBase.DecomposeRecordString(p_strSubStrings)) Then
                ' Validate and Assign ToAddressId
                bReturn = AssignToAddressId(p_strSubStrings(Fields.ToAddressId))
            End If

        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges

            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            bReturn = False

        End Try

        Return bReturn

    End Function

    Private Function AssignToAddressId(ByVal Value As Integer) As Boolean

        Try

            Dim bRetVal As Boolean = True

            If Value > 0 Then
                _iToAddId = Value
            Else
                SetError("Invalid ToAddressId")
                _iToAddId = 0
                bRetVal = False
            End If

            Return bRetVal

        Catch ex As Exception

            SetError(ex.Message & " [ScanRecord::AssignToAddressId]")
            Return False

        End Try

    End Function

    Protected Shadows Sub Reset()

        MyBase.Reset()
        _iToAddId = 0

    End Sub


    Public Overrides Function AssignRecord(ByVal p_strLine As String) As Boolean

        Try

            ' Check #1 - Count number of fields in record
            Dim str() As String
            Dim iFieldCount As Integer

            str = p_strLine.Split(Delimiter)
            iFieldCount = str.Length

            If iFieldCount = ExpectedFieldCount Then
                SetRecordString(p_strLine)
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

End Class
