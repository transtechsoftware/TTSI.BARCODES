Imports System
Imports System.Collections

Public Class PickupManifestRecords

    Implements IEnumerable

    'Private m_elements As PickupManifestRecord()
    Private m_elements As New ArrayList

    Public Sub New()

    End Sub

    Public Sub Add(ByVal p_objRecord As PickupManifestRecord)
        m_elements.Add(p_objRecord)
    End Sub

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return New PickupManifestRecordEnumerator(Me)
    End Function

    Private Class PickupManifestRecordEnumerator
        Implements IEnumerator

        Private m_iPos As Integer = -1
        Private m_objPickupManifestRecords As PickupManifestRecords

        Public Sub New(ByVal p_objPickupManifestRecords As PickupManifestRecords)
            Me.m_objPickupManifestRecords = p_objPickupManifestRecords
        End Sub

        Public ReadOnly Property Current() As Object Implements System.Collections.IEnumerator.Current
            Get
                'Return m_objPickupManifestRecords.m_elements(m_iPos)
                Return m_objPickupManifestRecords.m_elements.Item(m_iPos)
            End Get
        End Property

        Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
            'If (m_iPos < m_objPickupManifestRecords.m_elements.Length - 1) Then
            If (m_iPos < m_objPickupManifestRecords.m_elements.Count - 1) Then
                m_iPos += 1
                Return True
            Else
                Return False
            End If
        End Function

        Public Sub Reset() Implements System.Collections.IEnumerator.Reset
            m_iPos = -1
        End Sub
    End Class

End Class
