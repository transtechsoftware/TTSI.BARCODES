Imports System.Data
Imports System.IO

Public Class PickupManifestMapper
    Inherits PickupManifest

    Sub New(ByVal p_strFileName As String, ByVal p_sFileVersion As String)

        m_strFileName = p_strFileName
        m_sFileVersion = p_sFileVersion

        m_CurrentRecord = New PickupManifestRecord(m_strFileName, m_sFileVersion) 'Make sure there is a valid record for immediate use

        OpenFileForReading() 'Will open file and read into _Records

        'Change version string parent type if applicable
        If m_sFileVersion.CompareTo("V2s") = 0 Then m_sFileVersion = "V2"

    End Sub

    Protected Overrides Sub OpenFileForReading()

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
                    Dim objRecord As New PickupManifestRecord(m_strFileName, m_sFileVersion)
                    objRecord.LineNumber = m_iRecordCount
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


End Class
