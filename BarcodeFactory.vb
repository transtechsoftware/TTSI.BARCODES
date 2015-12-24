Public Module BarcodeFactory

    Public Enum BarcodeFormat
        Unknown = 0
        Unity_HOSS = 1
        TPC_Tracking = 2
        TPC_Operator = 3
        TPC_Point = 4
        Unreadable = 5
        IngramMicro = 6
    End Enum

    ' This function will determine which type of barcode is being processed and instantiate an instance
    ' of that type of barcode object
    Public Function NewBarcodeObject(ByVal p_strBarcode As String) As Barcode

        Dim oBarcode As Barcode

        Dim iLength As Integer = p_strBarcode.Length

        'Check for Ingram Company Code
        If iLength > 2 Then
            If String.Compare(p_strBarcode.Substring(1, 2), "IM") = 0 Then
                oBarcode = New IngramMicroBarcode(p_strBarcode)
                Return oBarcode
            End If
        End If

        'If iLength = 16 Then

        '    oBarcode = New HossBarcode(p_strBarcode)
        '    Return oBarcode

        'End If

        If iLength = 15 Then

            'Check Company Code
            Dim sCompanyCode As String = p_strBarcode.Substring(0, 3)
            sCompanyCode = sCompanyCode.ToUpper()

            If sCompanyCode.Equals("TPC") Then
                oBarcode = New TPCBarcode(p_strBarcode)
                Return oBarcode
            End If

        End If

        If iLength = 8 Then

            Dim strFormatCode As String = p_strBarcode.Substring(0, 1)
            strFormatCode = strFormatCode.ToUpper()

            If strFormatCode.Equals("P") Then

                oBarcode = New TPCPointBC(p_strBarcode)
                Return oBarcode

            End If

            If strFormatCode.Equals("E") Then

                oBarcode = New TPCOperatorBC(p_strBarcode)
                Return oBarcode

            End If

        End If

        ' Return a Generic Barcode with the Unknown Condition Set
        oBarcode = New Barcode(p_strBarcode)

        Return oBarcode

    End Function

End Module
