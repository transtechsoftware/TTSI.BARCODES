Imports System.IO
Imports System.Data

Public Class PickupManifest

    ' Members necessary for stream operation
    Protected m_bOpenForRead As Boolean = True 'True = Open for Read, False = Open for Write.  In write mode it will either create a new file or append to an existing one.
    Protected m_strFileName As String
    Protected m_sFileVersion As String
    Protected m_strErrorMessage As String
    Protected m_oReader As StreamReader = Nothing
    Protected m_oWriter As StreamWriter = Nothing

    ' Members necessary for record parsing.  These variables will hold the current record being written or read.
    Protected m_CurrentRecord As PickupManifestRecord
    Protected m_Records As New PickupManifestRecords
    Protected m_iRecordCount As Integer

    ' Collection of ScanList Records

    Public ReadOnly Property FileVersion() As String
        Get
            Return m_sFileVersion
        End Get
    End Property

    Public ReadOnly Property Records() As PickupManifestRecords
        Get
            Return m_Records
        End Get
    End Property

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return m_strErrorMessage
        End Get
    End Property

    Public ReadOnly Property RecordCount() As Integer
        Get
            Return m_iRecordCount
        End Get
    End Property

    Public ReadOnly Property OpenForRead() As Boolean
        Get
            Return m_bOpenForRead
        End Get
    End Property

    Public ReadOnly Property FileName() As String
        Get
            Return m_strFileName
        End Get
    End Property

    'Propery Wrappers
    Public Property Delimiter() As String
        Get
            Return m_CurrentRecord.Delimiter
        End Get
        Set(ByVal Value As String)
            m_CurrentRecord.Delimiter = Value
        End Set
    End Property

    Public ReadOnly Property BarcodeName() As String
        Get
            Return m_CurrentRecord.BarcodeName
        End Get
    End Property

    Public Property Barcode() As String
        Get
            Return m_CurrentRecord.TrackingNumber
        End Get
        Set(ByVal Value As String)
            m_CurrentRecord.TrackingNumber = Value
        End Set
    End Property

    Public Property TrackingNumber() As String
        Get
            Return m_CurrentRecord.TrackingNumber
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.TrackingNumber = value
        End Set
    End Property

    Public Property OrderId() As Integer
        Get
            Return m_CurrentRecord.OrderId
        End Get
        Set(ByVal value As Integer)
            m_CurrentRecord.OrderId = value
        End Set
    End Property
    Public Property FromCustId() As String
        Get
            Return m_CurrentRecord.FromCustId
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromCustId = value
        End Set
    End Property
    Public Property FromCustName() As String
        Get
            Return m_CurrentRecord.FromCustName
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromCustName = value
        End Set
    End Property
    Public Property FromAddRowId() As Integer
        Get
            Return m_CurrentRecord.FromAddRowId
        End Get
        Set(ByVal value As Integer)
            m_CurrentRecord.FromAddRowId = value
        End Set
    End Property
    Public Property FromLocId() As String   '10
        Get
            Return m_CurrentRecord.FromLocId
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromLocId = value
        End Set
    End Property
    Public Property FromLocName() As String '70
        Get
            Return m_CurrentRecord.FromLocName
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromLocName = value
        End Set
    End Property
    Public Property FromLocStreet() As String   '40
        Get
            Return m_CurrentRecord.FromLocStreet
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromLocStreet = value
        End Set
    End Property
    Public Property FromLocAddress2() As String '30
        Get
            Return m_CurrentRecord.FromLocAddress2
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromLocAddress2 = value
        End Set
    End Property
    Public Property FromLocCity() As String '32
        Get
            Return m_CurrentRecord.FromLocCity
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromLocCity = value
        End Set
    End Property
    Public Property FromLocState() As String    '2
        Get
            Return m_CurrentRecord.FromLocState
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromLocState = value
        End Set
    End Property
    Public Property FromLocZip() As String '10
        Get
            Return m_CurrentRecord.FromLocZip
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromLocZip = value
        End Set
    End Property
    Public Property FromLocPhone() As String '20
        Get
            Return m_CurrentRecord.FromLocPhone
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromLocPhone = value
        End Set
    End Property
    Public Property FromLocContact() As String '40
        Get
            Return m_CurrentRecord.FromLocContact
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromLocContact = value
        End Set
    End Property
    Public Property FromLocEmail() As String '30
        Get
            Return m_CurrentRecord.FromLocEmail
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.FromLocEmail = value
        End Set
    End Property
    Public Property ToAddRowId() As Integer
        Get
            Return m_CurrentRecord.ToAddRowId
        End Get
        Set(ByVal value As Integer)
            m_CurrentRecord.ToAddRowId = value
        End Set
    End Property
    Public Property ToCustId() As String '10
        Get
            Return m_CurrentRecord.ToCustId
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToCustId = value
        End Set
    End Property
    Public Property ToCustName() As String '70
        Get
            Return m_CurrentRecord.ToCustName
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToCustName = value
        End Set
    End Property
    Public Property ToLocId() As String '10
        Get
            Return m_CurrentRecord.ToLocId
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToLocId = value
        End Set
    End Property
    Public Property ToLocName() As String '70
        Get
            Return m_CurrentRecord.ToLocName
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToLocName = value
        End Set
    End Property
    Public Property ToLocStreet() As String '40
        Get
            Return m_CurrentRecord.ToLocStreet
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToLocStreet = value
        End Set
    End Property
    Public Property ToLocAddress2() As String '30
        Get
            Return m_CurrentRecord.ToLocAddress2
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToLocAddress2 = value
        End Set
    End Property
    Public Property ToLocCity() As String '32
        Get
            Return m_CurrentRecord.ToLocCity
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToLocCity = value
        End Set
    End Property
    Public Property ToLocState() As String '2
        Get
            Return m_CurrentRecord.ToLocState
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToLocState = value
        End Set
    End Property
    Public Property ToLocZip() As String '10
        Get
            Return m_CurrentRecord.ToLocZip
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToLocZip = value
        End Set
    End Property
    Public Property ToLocContact() As String '40
        Get
            Return m_CurrentRecord.ToLocContact
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToLocContact = value
        End Set
    End Property
    Public Property ToLocPhone() As String '20
        Get
            Return m_CurrentRecord.ToLocPhone
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToLocPhone = value
        End Set
    End Property
    Public Property ToLocEmail() As String '30
        Get
            Return m_CurrentRecord.ToLocEmail
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ToLocEmail = value
        End Set
    End Property
    Public Property Weight() As Decimal 'dec(6,2)
        Get
            Return m_CurrentRecord.Weight
        End Get
        Set(ByVal value As Decimal)
            m_CurrentRecord.Weight = value
        End Set
    End Property
    Public Property Pieces() As String '10
        Get
            Return m_CurrentRecord.Pieces
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.Pieces = value
        End Set
    End Property
    Public Property SentByName() As String '30
        Get
            Return m_CurrentRecord.SentByName
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.SentByName = value
        End Set
    End Property
    Public Property CartonCode() As String '20
        Get
            Return m_CurrentRecord.CartonCode
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.CartonCode = value
        End Set
    End Property
    Public Property Dimensions() As String '26 (shared)
        Get
            Return m_CurrentRecord.Dimensions
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.Dimensions = value
        End Set
    End Property
    Public Property ServiceLevel() As String '20
        Get
            Return m_CurrentRecord.ServiceLevel
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ServiceLevel = value
        End Set
    End Property
    Public Property BillType() As String '20
        Get
            Return m_CurrentRecord.BillType
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.BillType = value
        End Set
    End Property
    Public Property BillNum() As String '50
        Get
            Return m_CurrentRecord.BillNum
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.BillNum = value
        End Set
    End Property
    Public Property TranDate() As Date
        Get
            Return m_CurrentRecord.TranDate
        End Get
        Set(ByVal value As Date)
            m_CurrentRecord.TranDate = value
        End Set
    End Property
    Public ReadOnly Property UniqueRecordId() As String '29
        Get
            Return m_CurrentRecord.UniqueRecordId
        End Get
    End Property
    Public Property Void() As String '1 (Y|N)
        Get
            Return m_CurrentRecord.Void
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.Void = value
        End Set
    End Property
    Public Property ReferenceNumber() As String '40
        Get
            Return m_CurrentRecord.ReferenceNumber
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ReferenceNumber = value
        End Set
    End Property
    Public Property PONumber() As String '40
        Get
            Return m_CurrentRecord.PONumber
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.PONumber = value
        End Set
    End Property
    Public Property ThirdPartyBillNum() As String '40
        Get
            Return m_CurrentRecord.ThirdPartyBillNum
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.ThirdPartyBillNum = value
        End Set
    End Property
    Public Property Modifiers() As String '113 (shared)
        Get
            Return m_CurrentRecord.Modifiers
        End Get
        Set(ByVal value As String)
            m_CurrentRecord.Modifiers = value
        End Set
    End Property
    Public Property DeclaredValue() As Decimal '150 (shared)
        Get
            Return m_CurrentRecord.DeclaredValue
        End Get
        Set(ByVal value As Decimal)
            m_CurrentRecord.DeclaredValue = value
        End Set
    End Property

    Public Sub New()

    End Sub

    ' This constructor will populate it _Records collection with Scan List Records delivered via a DataSet
    ' The DataSet is assumed to contain 1 table at index 0 with 1 column named 'RecordString'.  The RecordString row
    ' has the exact same format and values as a single line of a text-based ScanList file would have
    Sub New(ByRef p_dsDataSet As DataSet)

        m_strErrorMessage = String.Empty

        Try

            m_CurrentRecord = New PickupManifestRecord(m_strFileName) 'Make sure there is a valid record for immediate use

            If PopulateFromDataSet(p_dsDataSet) = False Then

                m_Records = Nothing

            End If

        Catch ex As Exception

            m_strErrorMessage = ex.Message
            m_Records = Nothing

        End Try


    End Sub

    ' This constructor will attempt to open the file in Read mode
    Sub New(ByVal p_strFileName As String)

        m_strFileName = p_strFileName

        m_CurrentRecord = New PickupManifestRecord(m_strFileName) 'Make sure there is a valid record for immediate use

        OpenFileForReading() 'Will open file and read into _Records
    End Sub

    ' This constructor will allow the caller to specify whether it should open for Read (True) or Write (False)
    Sub New(ByVal p_strFileName As String, ByVal p_bOpenForRead As Boolean)

        m_strFileName = p_strFileName

        m_CurrentRecord = New PickupManifestRecord(p_strFileName) 'Make sure there is a valid record for immediate use

        If p_bOpenForRead = True Then
            OpenFileForReading() 'Will open file and read into _Records
        Else
            OpenFileForWriting()
        End If

    End Sub

    Protected Overridable Sub OpenFileForReading()

        Try

            ' Open the file for reading
            m_oReader = New StreamReader(m_strFileName)
            m_bOpenForRead = True
            m_Records = New PickupManifestRecords
            m_iRecordCount = 0

            ' Read the contents of the file into the Records collection
            Dim strCurrentLine As String
            Do
                strCurrentLine = m_oReader.ReadLine()
                If strCurrentLine Is Nothing Then
                    m_oReader.Close()
                    Exit Do
                Else
                    Dim objRecord As New PickupManifestRecord(m_strFileName)
                    If objRecord.AssignRecord(strCurrentLine) Then
                        m_Records.Add(objRecord)
                        m_iRecordCount += 1
                    Else
                        If objRecord.HasError Then
                            m_strErrorMessage = "Error in Record at Line " & m_iRecordCount + 1 & ":  " & objRecord.ErrorMessage
                            m_Records = Nothing
                            m_strFileName = ""
                            m_oReader.Close()
                            Exit Do
                        End If
                    End If
                End If
            Loop

        Catch ex As Exception
            m_strErrorMessage = ex.Message
            m_strFileName = ""
            If Not IsNothing(m_oReader) Then m_oReader.Close()
        End Try

    End Sub

    Private Sub OpenFileForWriting()

        Try

            'TO DO: If file exists, read current contents into _Records.  
            '       _Records should be kept in-sync with File Contents to allow for faster searching and decision-making.

            ' Open file for writing
            m_oWriter = New StreamWriter(m_strFileName, True)
            m_bOpenForRead = False

        Catch ex As Exception

            m_strErrorMessage = ex.Message
            m_strFileName = ""
            m_oWriter.Close()

        End Try

    End Sub

    Private Function ComposeRecord() As String

        ''Not Currently Implemented
        Return String.Empty

        '_CurrentRecord.Stamp()
        'Return _CurrentRecord.RecordString()

    End Function

    Public Function WriteRecord() As Boolean

        Try

            If Not (m_oWriter Is Nothing) Then

                m_oWriter.WriteLine(ComposeRecord())
                m_oWriter.Flush()
                m_Records.Add(m_CurrentRecord)
                m_CurrentRecord = New PickupManifestRecord("") ' When ready to implement, a valid file name must be supplied
                Return True

            Else

                m_strErrorMessage = "Unable to Write Record to File : [WriteRecord()]"
                m_strFileName = ""
                m_oWriter.Close()
                Return False

            End If

        Catch ex As Exception

            m_strErrorMessage = ex.Message & " : [WriteRecord()]"
            m_strFileName = ""
            m_oWriter.Close()

        End Try

    End Function

    Public Sub CloseFile()

        ' Execute code that will close the open files here
        If Not (m_oReader Is Nothing) Then
            m_oReader.Close()
            m_oReader = Nothing
        End If

        If Not (m_oWriter Is Nothing) Then
            m_oWriter.Close()
            m_oWriter = Nothing
        End If

        ' Clear the _Records collection
        m_Records = Nothing

    End Sub

    Private Function PopulateFromDataSet(ByRef p_dsDataSet As DataSet) As Boolean

        Dim bReturnValue As Boolean = True

        Try

            'This Function is Not Implemented at This Time
            bReturnValue = False

            '' Test for Valid DataSet
            'If Not p_dsDataSet Is Nothing Then

            '    ' Test for Valid Table
            '    If p_dsDataSet.Tables.Count > 0 Then

            '        'Test for expected number of columns
            '        If p_dsDataSet.Tables(0).Columns.Count > 0 Then

            '            ' Prepare Required Objects
            '            _Records = New PickupManifestRecords
            '            _iRecordCount = 0

            '            ' Read the contents of the file into the Records collection
            '            Dim strCurrentLine As String
            '            For Each dr As DataRow In p_dsDataSet.Tables(0).Rows

            '                strCurrentLine = CStr(dr.Item(1))

            '                If strCurrentLine Is Nothing Then

            '                    Exit For

            '                Else

            '                    Dim objRecord As New PickupManifestRecord

            '                    If objRecord.AssignRecord(strCurrentLine) Then

            '                        _Records.Add(objRecord)
            '                        _iRecordCount += 1

            '                    Else

            '                        _strErrorMessage = "Error in Record at RowId =  " & dr.Item("RowId") & ":  " & objRecord.ErrorMessage
            '                        _Records = Nothing
            '                        bReturnValue = False

            '                        Exit For

            '                    End If

            '                End If

            '            Next

            '        Else

            '            _strErrorMessage = "Invalid Dataset: Table contains no columns"
            '            bReturnValue = False

            '        End If

            '    Else

            '        _strErrorMessage = "Invalid Dataset: Contains no tables"
            '        bReturnValue = False

            '    End If

            'Else

            '    _strErrorMessage = "Invalid DataSet: Set to Nothing"
            '    bReturnValue = False

            'End If

        Catch ex As Exception

            m_strErrorMessage = ex.Message
            bReturnValue = False

        End Try

        Return bReturnValue

    End Function

End Class
