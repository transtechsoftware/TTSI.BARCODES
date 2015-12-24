Imports System
Imports System.Collections

Public Class Cartons
    Implements IEnumerable

    'Private m_elements As Cartons()
    Private m_elements As New ArrayList
    Private m_bAllowUnknown As Boolean

    Public Sub New(Optional ByVal p_bAllowUnknown As Boolean = False)
        m_bAllowUnknown = p_bAllowUnknown
    End Sub

    Public Sub New(ByVal p_sImplementation As String, Optional ByVal p_bAllowUnknown As Boolean = False)

        Me.New(p_bAllowUnknown)

        If String.Compare(p_sImplementation, "IngramMicro") = 0 Then
            'Initialize Valid Carton Codes
            Me.Add(New Carton(12, 9, 4, "01"))
            Me.Add(New Carton(12, 10, 6, "02"))
            Me.Add(New Carton(16, 14, 5, "03"))
            Me.Add(New Carton(24, 12, 6, "04"))
            Me.Add(New Carton(16, 14, 10, "05"))
            Me.Add(New Carton(22, 14, 12, "06"))
            Me.Add(New Carton(24, 17, 11, "07"))
            Me.Add(New Carton(17, 12, 6, "08"))
            Me.Add(New Carton(28, 22, 16, "11"))
            Me.Add(New Carton(24, 17, 6, "75"))
            Me.Add(New Carton(10, 13, 0.75, "AJ"))
            Me.Add(New Carton(9.8125, 6.375, 3, "F5"))
            Me.Add(New Carton(10.75, 9.125, 3.5, "F6"))
            Me.Add(New Carton(0, 0, 0, "CB"))
            Me.Add(New Carton(0, 0, 0, "CS"))
            Me.Add(New Carton(0, 0, 0, "CU"))
            Me.Add(New Carton(0, 0, 0, "E2"))
            Me.Add(New Carton(0, 0, 0, "M"))
            Me.Add(New Carton(0, 0, 0, "E0"))
        End If

    End Sub

    Public Sub Add(ByVal p_objRecord As Carton)
        m_elements.Add(p_objRecord)
    End Sub

    Private Sub AddX(ByVal p_sCode As String)
        Me.Add(New Carton(0, 0, 0, p_sCode))
    End Sub

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return New CartonEnumerator(Me)
    End Function

    Public Function ValidCode(ByVal p_sCode As String) As Boolean

        Dim bReturn As Boolean = False
        Dim c As Carton

        For Each c In Me
            If String.Compare(c.Code, p_sCode) = 0 Then
                bReturn = True
                Exit For
            End If
        Next

        If (bReturn = False And m_bAllowUnknown = True) Then
            Me.AddX(p_sCode)
            bReturn = True
        End If

        Return bReturn

    End Function

    Public Function GetCartonByCode(ByVal p_sCode As String) As Carton

        Dim c As Carton

        For Each c In Me
            If String.Compare(c.Code, p_sCode) = 0 Then Return c
        Next

        If m_bAllowUnknown = True Then
            Me.AddX(p_sCode)
            Return Me.GetCartonByCode(p_sCode)
        End If

        Return Nothing

    End Function

    Private Class CartonEnumerator
        Implements IEnumerator

        Private m_iPos As Integer = -1
        Private m_objCartons As Cartons

        Public Sub New(ByVal p_objCartons As Cartons)
            Me.m_objCartons = p_objCartons
        End Sub

        Public ReadOnly Property Current() As Object Implements System.Collections.IEnumerator.Current
            Get
                'Return m_objScanRecords.m_elements(m_iPos)
                Return m_objCartons.m_elements.Item(m_iPos)
            End Get
        End Property

        Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
            'If (m_iPos < m_objScanRecords.m_elements.Length - 1) Then
            If (m_iPos < m_objCartons.m_elements.Count - 1) Then
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
