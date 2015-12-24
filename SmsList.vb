Imports System.IO
Imports System.Data

' This class is designed to manage scan lists which will be created when capturing weight & barcodes data.
Public Class SmsList

    ' Members necessary for stream operation
    Private _bOpenForRead As Boolean = True 'True = Open for Read, False = Open for Write.  In write mode it will either create a new file or append to an existing one.
    Private _strFileName As String
    Private _strErrorMessage As String
    Private _oReader As StreamReader = Nothing
    Private _oWriter As StreamWriter = Nothing

    ' Members necessary for record parsing.  These variables will hold the current record being written or read.
    Private _CurrentRecord As SmsRecord
    Private _Records As New SmsRecords
    Private _iRecordCount As Integer

    ' Collection of SmsRecord Records

    Public ReadOnly Property Records() As SmsRecords
        Get
            Return _Records
        End Get
    End Property

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return _strErrorMessage
        End Get
    End Property

    Public ReadOnly Property RecordCount() As Integer
        Get
            Return _iRecordCount
        End Get
    End Property

    Public ReadOnly Property OpenForRead() As Boolean
        Get
            Return _bOpenForRead
        End Get
    End Property

    Public ReadOnly Property FileName() As String
        Get
            Return _strFileName
        End Get
    End Property

    Public Property Delimiter() As String
        Get
            Return _CurrentRecord.Delimiter
        End Get
        Set(ByVal Value As String)
            _CurrentRecord.Delimiter = Value
        End Set
    End Property

    Public Property Barcode() As String
        Get
            Return _CurrentRecord.Barcode
        End Get
        Set(ByVal Value As String)
            _CurrentRecord.Barcode = Value
        End Set
    End Property

    Public ReadOnly Property BarcodeName() As String
        Get
            Return _CurrentRecord.BarcodeName
        End Get
    End Property

    Public Property Weight() As Decimal
        Get
            Return _CurrentRecord.Weight
        End Get
        Set(ByVal Value As Decimal)
            _CurrentRecord.Weight = Value
        End Set
    End Property

    Public Property OperatorId() As String
        Get
            Return _CurrentRecord.OperatorId
        End Get
        Set(ByVal Value As String)
            _CurrentRecord.OperatorId = Value
        End Set
    End Property

    Public Property PointId() As String
        Get
            Return _CurrentRecord.PointId
        End Get
        Set(ByVal Value As String)
            _CurrentRecord.PointId = Value
        End Set
    End Property

    Public Property ProcessDate() As Date
        Get
            Return _CurrentRecord.ProcessDate
        End Get
        Set(ByVal Value As Date)
            _CurrentRecord.ProcessDate = Value
        End Set
    End Property

    Public Property BarcodeInstance() As Integer
        Get
            Return _CurrentRecord.X
        End Get
        Set(ByVal Value As Integer)
            _CurrentRecord.X = Value
        End Set
    End Property

    Public Property ScanError() As Integer
        Get
            Return _CurrentRecord.ScanError
        End Get
        Set(ByVal Value As Integer)
            _CurrentRecord.ScanError = Value
        End Set
    End Property

    Public Property BatchID() As Integer
        Get
            Return _CurrentRecord.BatchID
        End Get
        Set(ByVal Value As Integer)
            _CurrentRecord.BatchID = Value
        End Set
    End Property

    Public Property Processed() As Integer
        Get
            Return _CurrentRecord.Processed
        End Get
        Set(ByVal Value As Integer)
            _CurrentRecord.Processed = Value
        End Set
    End Property

    ' This constructor will populate it _Records collection with Scan List Records delivered via a DataSet
    ' The DataSet is assumed to contain 1 table at index 0 with 1 column named 'RecordString'.  The RecordString row
    ' has the exact same format and values as a single line of a text-based ScanList file would have
    Sub New(ByRef p_dsDataSet As DataSet)

        _strErrorMessage = String.Empty

        Try

            _CurrentRecord = New SmsRecord 'Make sure there is a valid record for immediate use

            If PopulateFromDataSet(p_dsDataSet) = False Then

                _Records = Nothing

            End If

        Catch ex As Exception

            _strErrorMessage = ex.Message
            _Records = Nothing

        End Try


    End Sub

    ' This constructor will attempt to open the file in Read mode
    Sub New(ByVal p_strFileName As String)

        _strFileName = p_strFileName

        _CurrentRecord = New SmsRecord 'Make sure there is a valid record for immediate use

        OpenFileForReading() 'Will open file and read into _Records
    End Sub

    ' This constructor will allow the caller to specify whether it should open for Read (True) or Write (False)
    Sub New(ByVal p_strFileName As String, ByVal p_bOpenForRead As Boolean)

        _strFileName = p_strFileName

        _CurrentRecord = New SmsRecord 'Make sure there is a valid record for immediate use

        If p_bOpenForRead = True Then
            OpenFileForReading() 'Will open file and read into _Records
        Else
            OpenFileForWriting()
        End If

    End Sub

    Private Sub OpenFileForReading()

        Try

            ' Open the file for reading
            _oReader = New StreamReader(_strFileName)
            _bOpenForRead = True
            _Records = New SmsRecords
            _iRecordCount = 0

            ' Read the contents of the file into the Records collection
            Dim strCurrentLine As String
            Do
                strCurrentLine = _oReader.ReadLine()
                If strCurrentLine Is Nothing Then
                    _oReader.Close()
                    Exit Do
                Else
                    Dim objRecord As New SmsRecord
                    If objRecord.AssignRecord(strCurrentLine) Then
                        _Records.Add(objRecord)
                        _iRecordCount += 1
                    Else
                        _strErrorMessage = "Error in Record at Line " & _iRecordCount + 1 & ":  " & objRecord.ErrorMessage
                        _Records = Nothing
                        _strFileName = ""
                        _oReader.Close()
                        Exit Do
                    End If
                End If
            Loop

        Catch ex As Exception
            _strErrorMessage = ex.Message
            _strFileName = ""
            If Not IsNothing(_oReader) Then _oReader.Close()
        End Try

    End Sub

    Private Sub OpenFileForWriting()

        Try

            'TO DO: If file exists, read current contents into _Records.  
            '       _Records should be kept in-sync with File Contents to allow for faster searching and decision-making.

            ' Open file for writing
            _oWriter = New StreamWriter(_strFileName, True)
            _bOpenForRead = False

        Catch ex As Exception

            _strErrorMessage = ex.Message
            _strFileName = ""
            _oWriter.Close()

        End Try

    End Sub

    Private Function ComposeRecord() As String

        _CurrentRecord.Stamp()
        Return _CurrentRecord.RecordString()

    End Function

    Public Function WriteRecord() As Boolean

        Try

            If Not (_oWriter Is Nothing) Then

                _oWriter.WriteLine(ComposeRecord())
                _oWriter.Flush()
                _Records.Add(_CurrentRecord)
                _CurrentRecord = New SmsRecord
                Return True

            Else

                _strErrorMessage = "Unable to Write Record to File : [WriteRecord()]"
                _strFileName = ""
                _oWriter.Close()
                Return False

            End If

        Catch ex As Exception

            _strErrorMessage = ex.Message & " : [WriteRecord()]"
            _strFileName = ""
            _oWriter.Close()

        End Try

    End Function

    Public Sub CloseFile()

        ' Execute code that will close the open files here
        If Not (_oReader Is Nothing) Then
            _oReader.Close()
            _oReader = Nothing
        End If

        If Not (_oWriter Is Nothing) Then
            _oWriter.Close()
            _oWriter = Nothing
        End If

        ' Clear the _Records collection
        _Records = Nothing

    End Sub

    Private Function PopulateFromDataSet(ByRef p_dsDataSet As DataSet) As Boolean

        Dim bReturnValue As Boolean = True

        Try

            ' Test for Valid DataSet
            If Not p_dsDataSet Is Nothing Then

                ' Test for Valid Table
                If p_dsDataSet.Tables.Count > 0 Then

                    'Test for expected number of columns
                    If p_dsDataSet.Tables(0).Columns.Count > 0 Then

                        ' Prepare Required Objects
                        _Records = New SmsRecords
                        _iRecordCount = 0

                        ' Read the contents of the file into the Records collection
                        Dim strCurrentLine As String
                        For Each dr As DataRow In p_dsDataSet.Tables(0).Rows

                            strCurrentLine = CStr(dr.Item(1))

                            If strCurrentLine Is Nothing Then

                                Exit For

                            Else

                                Dim objRecord As New SmsRecord

                                If objRecord.AssignRecord(strCurrentLine) Then

                                    _Records.Add(objRecord)
                                    _iRecordCount += 1

                                Else

                                    _strErrorMessage = "Error in Record at RowId =  " & dr.Item("RowId") & ":  " & objRecord.ErrorMessage
                                    _Records = Nothing
                                    bReturnValue = False

                                    Exit For

                                End If

                            End If

                        Next

                    Else

                        _strErrorMessage = "Invalid Dataset: Table contains no columns"
                        bReturnValue = False

                    End If

                Else

                    _strErrorMessage = "Invalid Dataset: Contains no tables"
                    bReturnValue = False

                End If

            Else

                _strErrorMessage = "Invalid DataSet: Set to Nothing"
                bReturnValue = False

            End If

        Catch ex As Exception

            _strErrorMessage = ex.Message
            bReturnValue = False

        End Try

        Return bReturnValue

    End Function

End Class

