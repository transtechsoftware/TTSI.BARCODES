Imports System.Text
Imports System.Globalization

' This class will instantiate a single record of type "TTSI PU MANIFEST IM r4"
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

Public Class PickupManifestRecord

    ' Used by Ingram Micro
    Public Enum FieldsV0
        TrackingNumber = 0
        '[EVENT].ThirdPartyBarcode  25
        '[MANIFEST].TrackingNum     25
        '[MANIFESTINVOICE].Trackingnum  25
        OrderId = 1
        '[EVENT].RefNum             40
        '[MANIFEST].RefNum          40
        '[MANIFESTINVOICE].Ref1     40
        FromCustId = 2
        '[EVENT].FromCustId         10
        '[MANIFEST].FromCustId      10
        '[MANIFESTINVOICE].FromCustId   10
        FromCustName = 3
        '[EVENT].FromCustName       70
        '[MANIFEST].FromCustName    70
        FromAddRowId = 4
        '[EVENT].FromAddId          int
        '[MANIFEST].FromAddId       int
        '[MANIFESTINVOICE].FromAddId    int
        FromLocId = 5
        '[EVENT].FromLocId          10
        '[MANIFEST].FromLocId       10
        '[MANIFESTINVOICE].FromLocId    10
        FromLocName = 6
        '[EVENT].FromLocName        70
        '[MANIFEST].FromLocName     70
        FromLocStreet = 7
        '[MANIFEST].FromAdd1        40
        FromLocCity = 8
        '[MANIFEST].FromCity        32
        FromLocAddress2 = 9
        '[MANIFEST].FromAdd2        30
        FromLocState = 10
        '[MANIFEST].FromState       2
        FromLocZip = 11
        '[MANIFEST].FromZip         10
        '[MANIFESTINVOICE].FromZip  50
        FromLocPhone = 12
        '[MANIFEST].FromPhone       20
        FromLocContact = 13
        '[MAINFEST].FromContact     40
        FromLocEmail = 14
        '[MANIFEST].FromEmail       30
        ToAddRowId = 15
        '[EVENT].ToAddId            int
        '[MANIFEST].ToAddId         int
        '[MANIFESTINVOICE].ToAddId  int
        ToCustId = 16
        '[MANIFEST].ToCustId        10 
        '[MANIFESTINVOICE].ToCustId 10
        ToCustName = 17
        '[MANIFEST].ToCustName      70
        ToLocId = 18
        '[EVENT].ToLocId            10
        '[MANIFEST].ToLocid         10
        '[MANIFESTINVOICE].ToLocId  10
        ToLocName = 19
        '[EVENT].ToLocName          70
        '[MANIFEST].ToLocName       70
        ToLocStreet = 20
        '[MANIFEST].ToAdd1          40
        ToLocAddress2 = 21
        '[MANIFEST].ToAdd2          30
        ToLocCity = 22
        '[EVENT].ToCity             32
        '[MANIFEST].ToCity          32
        ToLocState = 23
        '[MANIFEST].ToState         2
        ToLocZip = 24
        '[MANIFEST].ToZip           10
        '[MANIFESTINVOICE].ToZip    50
        ToLocContact = 25
        '[MANIFEST].ToContact       40
        ToLocPhone = 26
        '[MANIFEST].ToPhone         20
        ToLocEmail = 27
        '[MANIFEST].ToEmail         30
        Weight = 28
        '[EVENT].Weight             dec(6,2)
        '[MANIFEST].Weight          dec(6,2)
        '[MANIFESTINVOICE].Weight   dec(6,2)
        Pieces = 29
        '[EVENT].Pieces             10
        '[MANIFEST].Pieces          10
        '[MANIFESTINVOICE].Pieces   50
        SentByName = 30
        '[MANIFEST].SentBy          30
        CartonCode = 31
        '[EVENT].ParcelType         20
        '[MANIFEST].ParcelType      20
        '[MANIFESTINVOICE].ParcelType   20
        Dimensions = 32
        '[MANIFEST].SpecialHandle       150 (shared)
        '[MANIFESTINVOICE].SpecialHandle       150 (shared)
        ServiceLevel = 33
        '[MANIFEST].ServiceLevel        20
        '[MANIFESTINVOICE].ServiceLevel 20
        BillType = 34
        '[MANIFEST].BillType        20
        '[MANIFESTINVOICE].BillType 20
        BillNum = 35
        '[MANIFEST].BillNum         50
        '[MANIFESTINVOICE].BillNum  50
        TranDate = 36
        '[EVENT].ScanDate           Date
        '[MANIFEST].[DateTime]      Date
        '[MANIFESTINVOICE].[DateTime]   Date
        UniqueRecordId = 37
        '[MANIFEST].RowId           29
        '[MANIFESTINVOICE].RowId    29
        Void = 38
        '[EVENT].Void               1
        '[MANIFEST].Void            1
        '[MANIFESTINVOICE].Void     1
        ReferenceNumber = 39
        '[MANIFESTINVOICE].Ref2         40
        PONumber = 40
        '[MANIFESTINVOICE].Ref3         40 
        ThirdPartyBillNum = 41
        '[MANIFESTINVOICE].Ref4         40 
        Modifiers = 42
        '[MANIFEST].SpecialHandle       150 (shared)
        '[MANIFESTINVOICE].SpecialHandle       150 (shared)
        DeclaredValue = 43
        '[MANIFEST].SpecialHandle       150 (shared)
        '[MANIFESTINVOICE].SpecialHandle    150 (shared)
        EmptyField = 44
        ' Compensate for the fact Ingram-Micro is adding an extra tilde at the end of each record
        ' 0 bytes expected
    End Enum

    ' Used by Partner Books
    Public Enum FieldsV1
        TranDate = 0
        '[EVENT].ScanDate           Date
        '[MANIFEST].[DateTime]      Date
        '[MANIFESTINVOICE].[DateTime]   Date
        FromCustId = 1
        '[EVENT].FromCustId         10
        '[MANIFEST].FromCustId      10
        '[MANIFESTINVOICE].FromCustId   10
        FromCustName = 2
        '[EVENT].FromCustName       70
        '[MANIFEST].FromCustName    70
        FromLocStreet = 3
        '[MANIFEST].FromAdd1        40
        FromLocAddress2 = 4
        '[MANIFEST].FromAdd2        30
        FromLocCity = 5
        '[MANIFEST].FromCity        32
        FromLocState = 6
        '[MANIFEST].FromState       2
        FromLocZip = 7
        '[MANIFEST].FromZip         10
        '[MANIFESTINVOICE].FromZip  50
        ToLocId = 8
        '[EVENT].ToLocId            10
        '[MANIFEST].ToLocid         10
        '[MANIFESTINVOICE].ToLocId  10
        ToLocName = 9
        '[EVENT].ToLocName          70
        '[MANIFEST].ToLocName       70
        ToLocStreet = 10
        '[MANIFEST].ToAdd1          40
        ToLocAddress2 = 11
        '[MANIFEST].ToAdd2          30
        ToLocCity = 12
        '[EVENT].ToCity             32
        '[MANIFEST].ToCity          32
        ToLocState = 13
        '[MANIFEST].ToState         2
        ToLocZip = 14
        '[MANIFEST].ToZip           10
        '[MANIFESTINVOICE].ToZip    50
        TrackingNumber = 15
        '[EVENT].ThirdPartyBarcode  25
        '[MANIFEST].TrackingNum     25
        '[MANIFESTINVOICE].Trackingnum  25
        OrderId = 16
        '[EVENT].RefNum             40
        '[MANIFEST].RefNum          40
        '[MANIFESTINVOICE].Ref1     40
        Weight = 17
        '[EVENT].Weight             dec(6,2)
        '[MANIFEST].Weight          dec(6,2)
        '[MANIFESTINVOICE].Weight   dec(6,2)
        ServiceLevel = 18
        '[MANIFEST].ServiceLevel        20
        '[MANIFESTINVOICE].ServiceLevel 20
        CartonCode = 19
        '[EVENT].ParcelType         20
        '[MANIFEST].ParcelType      20
        '[MANIFESTINVOICE].ParcelType   20
    End Enum

    ' Used by MedExpress
    Public Enum FieldsV2
        TrackingNumber = 5
        '[EVENT].ThirdPartyBarcode  25
        '[MANIFEST].TrackingNum     25
        '[MANIFESTINVOICE].Trackingnum  25
        ToLocName = 0
        '[EVENT].ToLocName          70
        '[MANIFEST].ToLocName       70
        ToLocStreet = 1
        '[MANIFEST].ToAdd1          40
        ToLocCity = 2
        '[EVENT].ToCity             32
        '[MANIFEST].ToCity          32
        ToLocState = 3
        '[MANIFEST].ToState         2
        ToLocZip = 4
        '[MANIFEST].ToZip           10
        '[MANIFESTINVOICE].ToZip    50
        'OrderId = 6
        ''[EVENT].RefNum             40
        ''[MANIFEST].RefNum          40
        ''[MANIFESTINVOICE].Ref1     40
        'ServiceLevel = 7
        ''[MANIFEST].ServiceLevel        20
        ''[MANIFESTINVOICE].ServiceLevel 20


        'ToCustName = 17
        ''[MANIFEST].ToCustName      70
        'FromCustId = 2
        ''[EVENT].FromCustId         10
        ''[MANIFEST].FromCustId      10
        ''[MANIFESTINVOICE].FromCustId   10
        'FromCustName = 3
        ''[EVENT].FromCustName       70
        ''[MANIFEST].FromCustName    70
        'FromAddRowId = 4
        ''[EVENT].FromAddId          int
        ''[MANIFEST].FromAddId       int
        ''[MANIFESTINVOICE].FromAddId    int
        'FromLocId = 5
        ''[EVENT].FromLocId          10
        ''[MANIFEST].FromLocId       10
        ''[MANIFESTINVOICE].FromLocId    10
        'FromLocName = 6
        ''[EVENT].FromLocName        70
        ''[MANIFEST].FromLocName     70
        'FromLocStreet = 7
        ''[MANIFEST].FromAdd1        40
        'FromLocCity = 8
        ''[MANIFEST].FromCity        32
        'FromLocAddress2 = 9
        ''[MANIFEST].FromAdd2        30
        'FromLocState = 10
        ''[MANIFEST].FromState       2
        'FromLocZip = 11
        ''[MANIFEST].FromZip         10
        ''[MANIFESTINVOICE].FromZip  50
        'FromLocPhone = 12
        ''[MANIFEST].FromPhone       20
        'FromLocContact = 13
        ''[MAINFEST].FromContact     40
        'FromLocEmail = 14
        ''[MANIFEST].FromEmail       30
        'ToAddRowId = 15
        ''[EVENT].ToAddId            int
        ''[MANIFEST].ToAddId         int
        ''[MANIFESTINVOICE].ToAddId  int
        'ToCustId = 16
        ''[MANIFEST].ToCustId        10 
        ''[MANIFESTINVOICE].ToCustId 10
        'ToLocId = 18
        ''[EVENT].ToLocId            10
        ''[MANIFEST].ToLocid         10
        ''[MANIFESTINVOICE].ToLocId  10
        'ToLocAddress2 = 21
        ''[MANIFEST].ToAdd2          30
        'ToLocContact = 25
        ''[MANIFEST].ToContact       40
        'ToLocPhone = 26
        ''[MANIFEST].ToPhone         20
        'ToLocEmail = 27
        ''[MANIFEST].ToEmail         30
        'Weight = 28
        ''[EVENT].Weight             dec(6,2)
        ''[MANIFEST].Weight          dec(6,2)
        ''[MANIFESTINVOICE].Weight   dec(6,2)
        'Pieces = 29
        ''[EVENT].Pieces             10
        ''[MANIFEST].Pieces          10
        ''[MANIFESTINVOICE].Pieces   50
        'SentByName = 30
        ''[MANIFEST].SentBy          30
        'CartonCode = 31
        ''[EVENT].ParcelType         20
        ''[MANIFEST].ParcelType      20
        ''[MANIFESTINVOICE].ParcelType   20
        'Dimensions = 32
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle       150 (shared)
        'BillType = 34
        ''[MANIFEST].BillType        20
        ''[MANIFESTINVOICE].BillType 20
        'BillNum = 35
        ''[MANIFEST].BillNum         50
        ''[MANIFESTINVOICE].BillNum  50
        'TranDate = 36
        ''[EVENT].ScanDate           Date
        ''[MANIFEST].[DateTime]      Date
        ''[MANIFESTINVOICE].[DateTime]   Date
        'UniqueRecordId = 37
        ''[MANIFEST].RowId           29
        ''[MANIFESTINVOICE].RowId    29
        'Void = 38
        ''[EVENT].Void               1
        ''[MANIFEST].Void            1
        ''[MANIFESTINVOICE].Void     1
        'ReferenceNumber = 39
        ''[MANIFESTINVOICE].Ref2         40
        'PONumber = 40
        ''[MANIFESTINVOICE].Ref3         40 
        'ThirdPartyBillNum = 41
        ''[MANIFESTINVOICE].Ref4         40 
        'Modifiers = 42
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle       150 (shared)
        'DeclaredValue = 43
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle    150 (shared)
        'EmptyField = 44
        '' Compensate for the fact Ingram-Micro is adding an extra tilde at the end of each record
        '' 0 bytes expected
    End Enum

    ' Used by MedExpress
    Public Enum FieldsV2s
        TrackingNumber = 5
        '[EVENT].ThirdPartyBarcode  25
        '[MANIFEST].TrackingNum     25
        '[MANIFESTINVOICE].Trackingnum  25
        ToLocName = 0
        '[EVENT].ToLocName          70
        '[MANIFEST].ToLocName       70
        ToLocStreet = 1
        '[MANIFEST].ToAdd1          40
        ToLocCity = 2
        '[EVENT].ToCity             32
        '[MANIFEST].ToCity          32
        ToLocState = 3
        '[MANIFEST].ToState         2
        ToLocZip = 4
        '[MANIFEST].ToZip           10
        '[MANIFESTINVOICE].ToZip    50
        'OrderId = 6
        ''[EVENT].RefNum             40
        ''[MANIFEST].RefNum          40
        ''[MANIFESTINVOICE].Ref1     40
        'ServiceLevel = 7
        ''[MANIFEST].ServiceLevel        20
        ''[MANIFESTINVOICE].ServiceLevel 20


        'ToCustName = 17
        ''[MANIFEST].ToCustName      70
        'FromCustId = 2
        ''[EVENT].FromCustId         10
        ''[MANIFEST].FromCustId      10
        ''[MANIFESTINVOICE].FromCustId   10
        'FromCustName = 3
        ''[EVENT].FromCustName       70
        ''[MANIFEST].FromCustName    70
        'FromAddRowId = 4
        ''[EVENT].FromAddId          int
        ''[MANIFEST].FromAddId       int
        ''[MANIFESTINVOICE].FromAddId    int
        'FromLocId = 5
        ''[EVENT].FromLocId          10
        ''[MANIFEST].FromLocId       10
        ''[MANIFESTINVOICE].FromLocId    10
        'FromLocName = 6
        ''[EVENT].FromLocName        70
        ''[MANIFEST].FromLocName     70
        'FromLocStreet = 7
        ''[MANIFEST].FromAdd1        40
        'FromLocCity = 8
        ''[MANIFEST].FromCity        32
        'FromLocAddress2 = 9
        ''[MANIFEST].FromAdd2        30
        'FromLocState = 10
        ''[MANIFEST].FromState       2
        'FromLocZip = 11
        ''[MANIFEST].FromZip         10
        ''[MANIFESTINVOICE].FromZip  50
        'FromLocPhone = 12
        ''[MANIFEST].FromPhone       20
        'FromLocContact = 13
        ''[MAINFEST].FromContact     40
        'FromLocEmail = 14
        ''[MANIFEST].FromEmail       30
        'ToAddRowId = 15
        ''[EVENT].ToAddId            int
        ''[MANIFEST].ToAddId         int
        ''[MANIFESTINVOICE].ToAddId  int
        'ToCustId = 16
        ''[MANIFEST].ToCustId        10 
        ''[MANIFESTINVOICE].ToCustId 10
        'ToLocId = 18
        ''[EVENT].ToLocId            10
        ''[MANIFEST].ToLocid         10
        ''[MANIFESTINVOICE].ToLocId  10
        'ToLocAddress2 = 21
        ''[MANIFEST].ToAdd2          30
        'ToLocContact = 25
        ''[MANIFEST].ToContact       40
        'ToLocPhone = 26
        ''[MANIFEST].ToPhone         20
        'ToLocEmail = 27
        ''[MANIFEST].ToEmail         30
        'Weight = 28
        ''[EVENT].Weight             dec(6,2)
        ''[MANIFEST].Weight          dec(6,2)
        ''[MANIFESTINVOICE].Weight   dec(6,2)
        'Pieces = 29
        ''[EVENT].Pieces             10
        ''[MANIFEST].Pieces          10
        ''[MANIFESTINVOICE].Pieces   50
        'SentByName = 30
        ''[MANIFEST].SentBy          30
        'CartonCode = 31
        ''[EVENT].ParcelType         20
        ''[MANIFEST].ParcelType      20
        ''[MANIFESTINVOICE].ParcelType   20
        'Dimensions = 32
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle       150 (shared)
        'BillType = 34
        ''[MANIFEST].BillType        20
        ''[MANIFESTINVOICE].BillType 20
        'BillNum = 35
        ''[MANIFEST].BillNum         50
        ''[MANIFESTINVOICE].BillNum  50
        'TranDate = 36
        ''[EVENT].ScanDate           Date
        ''[MANIFEST].[DateTime]      Date
        ''[MANIFESTINVOICE].[DateTime]   Date
        'UniqueRecordId = 37
        ''[MANIFEST].RowId           29
        ''[MANIFESTINVOICE].RowId    29
        'Void = 38
        ''[EVENT].Void               1
        ''[MANIFEST].Void            1
        ''[MANIFESTINVOICE].Void     1
        'ReferenceNumber = 39
        ''[MANIFESTINVOICE].Ref2         40
        'PONumber = 40
        ''[MANIFESTINVOICE].Ref3         40 
        'ThirdPartyBillNum = 41
        ''[MANIFESTINVOICE].Ref4         40 
        'Modifiers = 42
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle       150 (shared)
        'DeclaredValue = 43
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle    150 (shared)
        'EmptyField = 44
        '' Compensate for the fact Ingram-Micro is adding an extra tilde at the end of each record
        '' 0 bytes expected
    End Enum

    ' Used by Blue Package
    Public Enum FieldsV3
        ToLocName = 1
        '[EVENT].ToLocName          70
        '[MANIFEST].ToLocName       70
        ToLocStreet = 2
        '[MANIFEST].ToAdd1          40
        'ToLocAddress2 = 21
        ''[MANIFEST].ToAdd2          30
        ToLocCity = 3
        '[EVENT].ToCity             32
        '[MANIFEST].ToCity          32
        'ToLocState = 23
        ''[MANIFEST].ToState         2
        ToLocZip = 4
        '[MANIFEST].ToZip           10
        '[MANIFESTINVOICE].ToZip    50
        Weight = 6
        '[EVENT].Weight             dec(6,2)
        '[MANIFEST].Weight          dec(6,2)
        '[MANIFESTINVOICE].Weight   dec(6,2)
        Pieces = 5
        '[EVENT].Pieces             10
        '[MANIFEST].Pieces          10
        '[MANIFESTINVOICE].Pieces   50
        ReferenceNumber = 0
        '[MANIFESTINVOICE].Ref2         40
    End Enum

    ' Used by AutoZone
    Public Enum FieldsV4
        'Uploaded Fields
        FromLocId = 0
        '[EVENT].FromLocId          10
        '[MANIFEST].FromLocId       10
        '[MANIFESTINVOICE].FromLocId    10
        ToLocId = 1
        '[EVENT].ToLocId            10
        '[MANIFEST].ToLocid         10
        '[MANIFESTINVOICE].ToLocId  10
        TranDate = 2
        '[EVENT].ScanDate           Date
        '[MANIFEST].[DateTime]      Date
        '[MANIFESTINVOICE].[DateTime]   Date
        Pieces = 3
        '[EVENT].Pieces             10
        '[MANIFEST].Pieces          10
        '[MANIFESTINVOICE].Pieces   50
        CartonCode = 4
        '[EVENT].ParcelType         20
        '[MANIFEST].ParcelType      20
        '[MANIFESTINVOICE].ParcelType   20
        ServiceLevel = 5
        '[MANIFEST].ServiceLevel        20
        '[MANIFESTINVOICE].ServiceLevel 20
    End Enum

    ' Used by DSC Delivery
    Public Enum FieldsV5
        TrackingNumber = 0
        '[EVENT].ThirdPartyBarcode  25
        '[MANIFEST].TrackingNum     25
        '[MANIFESTINVOICE].Trackingnum  25
        ThirdPartyBillNum = 1
        '[MANIFESTINVOICE].Ref4         40 
        ServiceLevel = 2
        '[MANIFEST].ServiceLevel        20
        '[MANIFESTINVOICE].ServiceLevel 20
        FromLocName = 3
        '[EVENT].FromLocName        70
        '[MANIFEST].FromLocName     70
        FromLocCity = 4
        '[MANIFEST].FromCity        32
        ToLocName = 5
        '[EVENT].ToLocName          70
        '[MANIFEST].ToLocName       70
        ToLocCity = 6
        '[EVENT].ToCity             32
        '[MANIFEST].ToCity          32
        ToLocStreet = 7
        '[MANIFEST].ToAdd1          40
        ToLocAddress2 = 8
        '[MANIFEST].ToAdd2          30
        ToLocState = 9
        '[MANIFEST].ToState         2
        ToLocZip = 10
        '[MANIFEST].ToZip           10
        '[MANIFESTINVOICE].ToZip    50
        Pieces = 11
        '[EVENT].Pieces             10
        '[MANIFEST].Pieces          10
        '[MANIFESTINVOICE].Pieces   50
        Weight = 12
        '[EVENT].Weight             dec(6,2)
        '[MANIFEST].Weight          dec(6,2)
        '[MANIFESTINVOICE].Weight   dec(6,2)
        Modifiers = 13
        '[MANIFEST].SpecialHandle       150 (shared)
        '[MANIFESTINVOICE].SpecialHandle       150 (shared)
    End Enum

    ' Used by CALPIA
    Public Enum FieldsV6
        'In order fields appear in file
        FromCustName = 0
        '[EVENT].FromCustName       70
        '[MANIFEST].FromCustName    70
        FromCustMailingAddress = 1
        'Not Used
        FromLocCity = 2
        '[MANIFEST].FromCity        32
        FromLocState = 3
        '[MANIFEST].FromState       2
        FromLocPhone = 4
        '[MANIFEST].FromPhone       20
        ToLocName = 5
        '[EVENT].ToLocName          70
        '[MANIFEST].ToLocName       70
        ToLocStreet = 6
        '[MANIFEST].ToAdd1          40
        ToLocAddress2 = 7
        '[MANIFEST].ToAdd2          30
        ToLocCity = 8
        '[EVENT].ToCity             32
        '[MANIFEST].ToCity          32
        ToLocState = 9
        '[MANIFEST].ToState         2
        ToLocZip = 10
        '[MANIFEST].ToZip           10
        '[MANIFESTINVOICE].ToZip    50
        ToDepot = 11
        'Not Used
        TrackingNumber = 12
        '[EVENT].ThirdPartyBarcode  25
        '[MANIFEST].TrackingNum     25
        '[MANIFESTINVOICE].Trackingnum  25
        ToLocId = 13
        '[EVENT].ToLocId            10
        '[MANIFEST].ToLocid         10
        '[MANIFESTINVOICE].ToLocId  10
        TranDate = 14
        '[EVENT].ScanDate           Date
        '[MANIFEST].[DateTime]      Date
        '[MANIFESTINVOICE].[DateTime]   Date
        EmptyField = 15
        ' Compensate for the fact CalPia is adding an extra tilde at the end of each record. 0 bytes expected
    End Enum

    ' Used by ProCourier
    Public Enum FieldsV7
        TrackingNumber = 0
        '[EVENT].ThirdPartyBarcode  25
        '[MANIFEST].TrackingNum     25
        '[MANIFESTINVOICE].Trackingnum  25
        FromLocName = 1
        '[EVENT].FromLocName        70
        '[MANIFEST].FromLocName     70
        FromLocStreet = 2
        '[MANIFEST].FromAdd1        40
        FromLocCity = 3
        '[MANIFEST].FromCity        32
        FromLocState = 4
        '[MANIFEST].FromState       2
        FromLocZip = 5
        '[MANIFEST].FromZip         10
        '[MANIFESTINVOICE].FromZip  50
        ToLocName = 6
        '[EVENT].ToLocName          70
        '[MANIFEST].ToLocName       70
        ToLocStreet = 7
        '[MANIFEST].ToAdd1          40
        ToLocCity = 8
        '[EVENT].ToCity             32
        '[MANIFEST].ToCity          32
        ToLocState = 9
        '[MANIFEST].ToState         2
        ToLocZip = 10
        '[MANIFEST].ToZip           10
        '[MANIFESTINVOICE].ToZip    50
        ServiceLevel = 11
        '[MANIFEST].ServiceLevel        20
        '[MANIFESTINVOICE].ServiceLevel 20
        Pieces = 12
        '[EVENT].Pieces             10
        '[MANIFEST].Pieces          10
        '[MANIFESTINVOICE].Pieces   50

        ''Fields Above Populated From File
        'OrderId = 1
        ''[EVENT].RefNum             40
        ''[MANIFEST].RefNum          40
        ''[MANIFESTINVOICE].Ref1     40
        'FromCustId = 2
        ''[EVENT].FromCustId         10
        ''[MANIFEST].FromCustId      10
        ''[MANIFESTINVOICE].FromCustId   10
        'FromCustName = 3
        ''[EVENT].FromCustName       70
        ''[MANIFEST].FromCustName    70
        'FromAddRowId = 4
        ''[EVENT].FromAddId          int
        ''[MANIFEST].FromAddId       int
        ''[MANIFESTINVOICE].FromAddId    int
        'FromLocId = 5
        ''[EVENT].FromLocId          10
        ''[MANIFEST].FromLocId       10
        ''[MANIFESTINVOICE].FromLocId    10
        'FromLocAddress2 = 9
        ''[MANIFEST].FromAdd2        30
        'FromLocPhone = 12
        ''[MANIFEST].FromPhone       20
        'FromLocContact = 13
        ''[MAINFEST].FromContact     40
        'FromLocEmail = 14
        ''[MANIFEST].FromEmail       30
        'ToAddRowId = 15
        ''[EVENT].ToAddId            int
        ''[MANIFEST].ToAddId         int
        ''[MANIFESTINVOICE].ToAddId  int
        'ToCustId = 16
        ''[MANIFEST].ToCustId        10 
        ''[MANIFESTINVOICE].ToCustId 10
        'ToCustName = 17
        ''[MANIFEST].ToCustName      70
        'ToLocId = 18
        ''[EVENT].ToLocId            10
        ''[MANIFEST].ToLocid         10
        ''[MANIFESTINVOICE].ToLocId  10
        'ToLocAddress2 = 21
        ''[MANIFEST].ToAdd2          30
        'ToLocContact = 25
        ''[MANIFEST].ToContact       40
        'ToLocPhone = 26
        ''[MANIFEST].ToPhone         20
        'ToLocEmail = 27
        ''[MANIFEST].ToEmail         30
        'Weight = 28
        ''[EVENT].Weight             dec(6,2)
        ''[MANIFEST].Weight          dec(6,2)
        ''[MANIFESTINVOICE].Weight   dec(6,2)
        'SentByName = 30
        ''[MANIFEST].SentBy          30
        'CartonCode = 31
        ''[EVENT].ParcelType         20
        ''[MANIFEST].ParcelType      20
        ''[MANIFESTINVOICE].ParcelType   20
        'Dimensions = 32
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle       150 (shared)
        'BillType = 34
        ''[MANIFEST].BillType        20
        ''[MANIFESTINVOICE].BillType 20
        'BillNum = 35
        ''[MANIFEST].BillNum         50
        ''[MANIFESTINVOICE].BillNum  50
        'TranDate = 36
        ''[EVENT].ScanDate           Date
        ''[MANIFEST].[DateTime]      Date
        ''[MANIFESTINVOICE].[DateTime]   Date
        'UniqueRecordId = 37
        ''[MANIFEST].RowId           29
        ''[MANIFESTINVOICE].RowId    29
        'Void = 38
        ''[EVENT].Void               1
        ''[MANIFEST].Void            1
        ''[MANIFESTINVOICE].Void     1
        'ReferenceNumber = 39
        ''[MANIFESTINVOICE].Ref2         40
        'PONumber = 40
        ''[MANIFESTINVOICE].Ref3         40 
        'ThirdPartyBillNum = 41
        ''[MANIFESTINVOICE].Ref4         40 
        'Modifiers = 42
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle       150 (shared)
        'DeclaredValue = 43
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle    150 (shared)
        'EmptyField = 44
        '' Compensate for the fact Ingram-Micro is adding an extra tilde at the end of each record
        '' 0 bytes expected
    End Enum

    ' Used to reference fields in the import file DIO2CS
    Public Enum FieldsV8

        TrackingNumber = 0 'aka OrderNumber
        '[EVENT].ThirdPartyBarcode  25
        '[MANIFEST].TrackingNum     25
        '[MANIFESTINVOICE].Trackingnum  25
        ChargedBySub = 1 ' aka AmountCharged
        '[MANIFESTINVOICE].Ref5 40
        Weight = 2
        '[EVENT].Weight             dec(6,2)
        '[MANIFEST].Weight          dec(6,2)
        '[MANIFESTINVOICE].Weight   dec(6,2)
        FromLocId = 3 'CustomerCode
        '[EVENT].FromLocId          10
        '[MANIFEST].FromLocId       10
        '[MANIFESTINVOICE].FromLocId    10
        FromLocName = 4 'aka OriginName
        '[EVENT].FromLocName        70
        '[MANIFEST].FromLocName     70
        FromLocStreet = 5 'aka OriginAddress
        '[MANIFEST].FromAdd1        40
        FromLocCity = 6 'aka OriginCity
        '[MANIFEST].FromCity        32
        FromLocZip = 7 'aka OriginZip
        '[MANIFEST].FromZip         10
        '[MANIFESTINVOICE].FromZip  50
        ToLocName = 8 'aka DestName
        '[EVENT].ToLocName          70
        '[MANIFEST].ToLocName       70
        ToLocStreet = 9 'aka DestAddress
        '[MANIFEST].ToAdd1          40
        ToLocCity = 10 ' aka DestCity
        '[EVENT].ToCity             32
        '[MANIFEST].ToCity          32
        ToLocState = 11 'aka DestState
        '[MANIFEST].ToState         2
        ToLocZip = 12 'aka DestZip
        '[MANIFEST].ToZip           10
        '[MANIFESTINVOICE].ToZip    50
        Pieces = 13
        '[EVENT].Pieces             10
        '[MANIFEST].Pieces          10
        '[MANIFESTINVOICE].Pieces   50
        OrderId = 14 'aka OrderAlias
        '[EVENT].RefNum             40
        '[MANIFEST].RefNum          40
        '[MANIFESTINVOICE].Ref1     40
        TranDate = 15 'aka OrderDate
        '[EVENT].ScanDate           Date
        '[MANIFEST].[DateTime]      Date
        '[MANIFESTINVOICE].[DateTime]   Date
        CartonCode = 16
        '[EVENT].ParcelType         20
        '[MANIFEST].ParcelType      20
        '[MANIFESTINVOICE].ParcelType   20


#Region "Not Used"
        'FromCustId = 2
        ''[EVENT].FromCustId         10
        ''[MANIFEST].FromCustId      10
        ''[MANIFESTINVOICE].FromCustId   10
        'FromCustName = 3
        ''[EVENT].FromCustName       70
        ''[MANIFEST].FromCustName    70
        'FromAddRowId = 4
        ''[EVENT].FromAddId          int
        ''[MANIFEST].FromAddId       int
        ''[MANIFESTINVOICE].FromAddId    int
        'FromLocAddress2 = 9
        ''[MANIFEST].FromAdd2        30
        'FromLocState = 10
        ''[MANIFEST].FromState       2
        'FromLocPhone = 12
        ''[MANIFEST].FromPhone       20
        'FromLocContact = 13
        ''[MAINFEST].FromContact     40
        'FromLocEmail = 14
        ''[MANIFEST].FromEmail       30
        'ToAddRowId = 15
        ''[EVENT].ToAddId            int
        ''[MANIFEST].ToAddId         int
        ''[MANIFESTINVOICE].ToAddId  int
        'ToCustId = 16
        ''[MANIFEST].ToCustId        10 
        ''[MANIFESTINVOICE].ToCustId 10
        'ToCustName = 17
        ''[MANIFEST].ToCustName      70
        'ToLocId = 18
        ''[EVENT].ToLocId            10
        ''[MANIFEST].ToLocid         10
        ''[MANIFESTINVOICE].ToLocId  10
        'ToLocAddress2 = 21
        ''[MANIFEST].ToAdd2          30
        'ToLocContact = 25
        ''[MANIFEST].ToContact       40
        'ToLocPhone = 26
        ''[MANIFEST].ToPhone         20
        'ToLocEmail = 27
        ''[MANIFEST].ToEmail         30
        'SentByName = 30
        ''[MANIFEST].SentBy          30
        'Dimensions = 32
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle       150 (shared)
        'ServiceLevel = 33
        ''[MANIFEST].ServiceLevel        20
        ''[MANIFESTINVOICE].ServiceLevel 20
        'BillType = 34
        ''[MANIFEST].BillType        20
        ''[MANIFESTINVOICE].BillType 20
        'BillNum = 35
        ''[MANIFEST].BillNum         50
        ''[MANIFESTINVOICE].BillNum  50
        'UniqueRecordId = 37
        ''[MANIFEST].RowId           29
        ''[MANIFESTINVOICE].RowId    29
        'Void = 38
        ''[EVENT].Void               1
        ''[MANIFEST].Void            1
        ''[MANIFESTINVOICE].Void     1
        'ReferenceNumber = 39
        ''[MANIFESTINVOICE].Ref2         40
        'PONumber = 40
        ''[MANIFESTINVOICE].Ref3         40 
        'ThirdPartyBillNum = 41
        ''[MANIFESTINVOICE].Ref4         40 
        'Modifiers = 42
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle       150 (shared)
        'DeclaredValue = 43
        ''[MANIFEST].SpecialHandle       150 (shared)
        ''[MANIFESTINVOICE].SpecialHandle    150 (shared)
        'EmptyField = 44
        '' Compensate for the fact Ingram-Micro is adding an extra tilde at the end of each record
        '' 0 bytes expected
#End Region

    End Enum


    ' Members necessary for record parsing.  These variables will hold the current record being written or read.
    Private _oTrackingNumber As Barcode
    Private _sOrderId As String
    Private _sFromCustId As String '10
    Private _sFromCustName As String '70
    Private _iFromAddRowId As Integer
    Private _sFromLocId As String '10
    Private _sFromLocName As String '70
    Private _sFromLocStreet As String '40
    Private _sFromLocAddress2 As String '30
    Private _sFromLocCity As String '32
    Private _sFromLocState As String '2
    Private _sFromLocZip As String '10
    Private _sFromLocPhone As String '20
    Private _sFromLocContact As String '40
    Private _sFromLocEmail As String '30
    Private _iToAddRowId As Integer
    Private _sToCustId As String '10
    Private _sToCustName As String '70
    Private _sToLocId As String '10
    Private _sToLocName As String '70
    Private _sToLocStreet As String '40
    Private _sToLocAddress2 As String '30
    Private _sToLocCity As String '32
    Private _sToLocState As String '2
    Private _sToLocZip As String '10
    Private _sToLocContact As String '40
    Private _sToLocPhone As String '20
    Private _sToLocEmail As String '30
    Private _fWeight As Decimal 'dec(6,2)
    Private _sPieces As String '10
    Private _sSentByName As String '30
    Private _sCartonCode As String '20
    Private _sServiceLevel As String '20
    Private _sBillType As String '20
    Private _sBillNum As String '50
    Private _dtTranDate As Date
    Private _sUniqueRecordId As String '
    Private _sVoid As String '1 (Y|N)
    Private _sReferenceNumber As String '40
    Private _sPONumber As String '40
    Private _sThirdPartyBillNum As String '40

    Private _oSpecialHandle As SpecialHandleColumn
    'Private _sModifiers As String '150 (shared)
    'Private _fDeclaredValue As Decimal '150 (shared)
    'Private _sDimensions As String '150 (shared)
    Private _sEmptyField As String '0
    Private _fChargedBySub As Decimal

    ' Members that hold statistical data about the currently loaded record
    Private _chrDelimiter As Char
    Private _strRecordString As String
    Private _sFileName As String
    Private _iExpectedFieldCount As Integer
    Private _strErrorMessage As String
    Private _bHasError As Boolean
    Private _sFileFormat As String
    Private _iLineNum As Integer

    Private Sub PrepareForIngramMicroRecords(ByVal p_sFileName As String) ' V0 - Version used by Ingram-Micro

        Try

            'Defalut Values
            _chrDelimiter = "~"
            _fWeight = 0D
            _iExpectedFieldCount = 45
            _strErrorMessage = ""
            _bHasError = False
            _oSpecialHandle = New SpecialHandleColumn

            'Data Derived from File Name
            Dim sb As New StringBuilder

            '' Extract the TranDate from the File's Name and Validate it
            Dim sTranDate As String
            Dim iFileNameLength As Integer = p_sFileName.Length


            sTranDate = p_sFileName.Substring(iFileNameLength - 18, 14)

            'sb.Append(p_sFileName.Substring(9, 2)) ' MM
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(11, 2)) ' DD
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(5, 4)) ' YYYY
            'sb.Append(" ")
            'sb.Append(p_sFileName.Substring(13, 2)) ' hh
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(15, 2)) ' mm
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(17, 2)) ' ss

            sb.Append(sTranDate.Substring(4, 2)) ' MM
            sb.Append("/")
            sb.Append(sTranDate.Substring(6, 2)) ' DD
            sb.Append("/")
            sb.Append(sTranDate.Substring(0, 4)) ' YYYY
            sb.Append(" ")
            sb.Append(sTranDate.Substring(8, 2)) ' hh
            sb.Append(":")
            sb.Append(sTranDate.Substring(10, 2)) ' mm
            sb.Append(":")
            sb.Append(sTranDate.Substring(12, 2)) ' ss

            TranDate = CDate(sb.ToString())


        Catch ex As Exception

            Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

        End Try

    End Sub

    Private Sub PrepareForPartnerBooksRecords(ByVal p_sFileName As String) ' V1 - Version used by PartnerBooks

        Try

            'Defalut Values
            _chrDelimiter = Chr(9)  'Tab
            _fWeight = 0D
            _iExpectedFieldCount = 20
            _strErrorMessage = ""
            _bHasError = False
            _oSpecialHandle = New SpecialHandleColumn

            'Data Derived from File Name
            Dim sb As New StringBuilder

            '' Extract the TranDate from the File's Name and Validate it
            Dim sTranDate As String
            Dim iFileNameLength As Integer = p_sFileName.Length


            sTranDate = p_sFileName.Substring(iFileNameLength - 10, 6)

            'sb.Append(p_sFileName.Substring(9, 2)) ' MM
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(11, 2)) ' DD
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(5, 4)) ' YYYY
            'sb.Append(" ")
            'sb.Append(p_sFileName.Substring(13, 2)) ' hh
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(15, 2)) ' mm
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(17, 2)) ' ss

            sb.Append(sTranDate.Substring(2, 2)) ' MM
            sb.Append("/")
            sb.Append(sTranDate.Substring(4, 2)) ' DD
            sb.Append("/20")
            sb.Append(sTranDate.Substring(0, 2)) ' YY
            sb.Append(" 00:00:00")

            TranDate = CDate(sb.ToString())

        Catch ex As Exception

            Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

        End Try

    End Sub

    Private Sub PrepareForMedExRecords(ByVal p_sFileName As String) ' V2 - Version used by Med Express

        Try

            _iLineNum = 0

            'Defalut Values
            _chrDelimiter = Chr(126)  'Tilde
            _fWeight = 0D
            _iExpectedFieldCount = 6
            _strErrorMessage = ""
            _bHasError = False
            _oSpecialHandle = New SpecialHandleColumn

            _sFromCustId = "11117"
            _sFromCustName = "MED EXPRESS DRUG SYSTEMS"
            _sToCustId = _sFromCustId
            _sToCustName = _sFromCustName
            _iFromAddRowId = 38960
            _sFromLocId = "001"
            _sFromLocName = "MED EXPRESS DRUG SYSTEMS"
            _sFromLocStreet = "425 W RIDER ST"
            _sFromLocAddress2 = "STE B-2"
            _sFromLocCity = "PERRIS"
            _sFromLocState = "CA"
            _sFromLocZip = "92571"
            _sToCustId = "11117"
            _sPieces = "1"
            _sCartonCode = "BAG"
            _sBillType = "ACCOUNT"
            _sBillNum = "11117"

            _oTrackingNumber = New Barcode
            '_sServiceLevel = "OVERNIGHT"
            _sServiceLevel = String.Empty
            _sOrderId = String.Empty

            _sFromLocPhone = String.Empty
            _sFromLocContact = String.Empty
            _sFromLocEmail = String.Empty
            _sToLocId = String.Empty
            _sToLocAddress2 = String.Empty
            _sToLocContact = String.Empty
            _sToLocPhone = String.Empty
            _sToLocEmail = String.Empty
            _sSentByName = String.Empty
            _sVoid = "F"
            _sReferenceNumber = String.Empty
            _sPONumber = String.Empty
            _sThirdPartyBillNum = String.Empty


            'Data Derived from File Name
            Dim sb As New StringBuilder

            '' Extract the TranDate from the File's Name and Validate it
            Dim sTranDate As String
            Dim iFileNameLength As Integer = p_sFileName.Length


            sTranDate = p_sFileName.Substring(iFileNameLength - 15, 8)

            'sb.Append(p_sFileName.Substring(9, 2)) ' MM
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(11, 2)) ' DD
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(5, 4)) ' YYYY
            'sb.Append(" ")
            'sb.Append(p_sFileName.Substring(13, 2)) ' hh
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(15, 2)) ' mm
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(17, 2)) ' ss

            sb.Append(sTranDate.Substring(4, 2)) ' MM
            sb.Append("/")
            sb.Append(sTranDate.Substring(6, 2)) ' DD
            sb.Append("/")
            sb.Append(sTranDate.Substring(0, 4)) ' YYYY
            sb.Append(" 00:00:00")

            TranDate = CDate(sb.ToString())

        Catch ex As Exception

            Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

        End Try

    End Sub

    Private Sub PrepareForMedExSdRecords(ByVal p_sFileName As String) ' V2s - Version used by Med Express

        Try

            _iLineNum = 0

            'Defalut Values
            _chrDelimiter = Chr(126)  'Tilde
            _fWeight = 0D
            _iExpectedFieldCount = 6
            _strErrorMessage = ""
            _bHasError = False
            _oSpecialHandle = New SpecialHandleColumn

            _sFromCustId = "12005"
            _sFromCustName = "MEDEX SAME DAY"
            _sToCustId = _sFromCustId
            _sToCustName = _sFromCustName
            _iFromAddRowId = 50135
            _sFromLocId = "001"
            _sFromLocName = "MEDEX SAME DAY"
            _sFromLocStreet = "425 W RIDER ST"
            _sFromLocAddress2 = "STE B-2"
            _sFromLocCity = "PERRIS"
            _sFromLocState = "CA"
            _sFromLocZip = "92571"
            _sToCustId = "12005"
            _sPieces = "1"
            _sCartonCode = "BAG"
            _sBillType = "ACCOUNT"
            _sBillNum = "12005"

            _oTrackingNumber = New Barcode
            '_sServiceLevel = "OVERNIGHT"
            _sServiceLevel = String.Empty
            _sOrderId = String.Empty

            _sFromLocPhone = String.Empty
            _sFromLocContact = String.Empty
            _sFromLocEmail = String.Empty
            _sToLocId = String.Empty
            _sToLocAddress2 = String.Empty
            _sToLocContact = String.Empty
            _sToLocPhone = String.Empty
            _sToLocEmail = String.Empty
            _sSentByName = String.Empty
            _sVoid = "F"
            _sReferenceNumber = String.Empty
            _sPONumber = String.Empty
            _sThirdPartyBillNum = String.Empty


            'Data Derived from File Name
            Dim sb As New StringBuilder

            '' Extract the TranDate from the File's Name and Validate it
            Dim sTranDate As String
            Dim iFileNameLength As Integer = p_sFileName.Length


            sTranDate = p_sFileName.Substring(iFileNameLength - 15, 8)

            'sb.Append(p_sFileName.Substring(9, 2)) ' MM
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(11, 2)) ' DD
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(5, 4)) ' YYYY
            'sb.Append(" ")
            'sb.Append(p_sFileName.Substring(13, 2)) ' hh
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(15, 2)) ' mm
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(17, 2)) ' ss

            sb.Append(sTranDate.Substring(4, 2)) ' MM
            sb.Append("/")
            sb.Append(sTranDate.Substring(6, 2)) ' DD
            sb.Append("/")
            sb.Append(sTranDate.Substring(0, 4)) ' YYYY
            sb.Append(" 00:00:00")

            TranDate = CDate(sb.ToString())

        Catch ex As Exception

            Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

        End Try

    End Sub

    Private Sub PrepareForBluePackageRecords(ByVal p_sFileName As String) ' V3 - Version used by Blue Package

        Try

            _iLineNum = 0

            'Defalut Values
            _chrDelimiter = Chr(126)  'Tilde
            _iExpectedFieldCount = 7
            _strErrorMessage = ""
            _bHasError = False
            _oSpecialHandle = New SpecialHandleColumn

            _sFromCustId = "13015"
            _sFromCustName = "BLUE PACKAGE DELIVERY"
            _sToCustId = _sFromCustId
            _sToCustName = _sFromCustName
            _iFromAddRowId = 49909
            _sFromLocId = "001"
            _sFromLocName = "BLUE PACKAGE DELIVERY"
            _sFromLocStreet = "1636 GERVAIS AVE"
            _sFromLocAddress2 = "STE 3"
            _sFromLocCity = "SAINT PAUL"
            _sFromLocState = "MN"
            _sFromLocZip = "55109"
            _sToCustId = "13015"
            _sToLocState = "CA"
            _sCartonCode = "BOX"
            _sBillType = "ACCOUNT"
            _sBillNum = "13015"
            _sServiceLevel = "SAMEDAY"
            _sVoid = "F"

            'Fields that will be imported from the File
            _sReferenceNumber = String.Empty
            _sToLocName = String.Empty
            _sToLocStreet = String.Empty
            _sToLocCity = String.Empty
            _sToLocZip = String.Empty
            _sPieces = String.Empty
            _fWeight = 0D

            'Fields that will be dynamically generated
            _oTrackingNumber = New Barcode

            'Fields that will remain empty
            _sOrderId = String.Empty
            _sFromLocPhone = String.Empty
            _sFromLocContact = String.Empty
            _sFromLocEmail = String.Empty
            _sToLocId = String.Empty
            _sToLocAddress2 = String.Empty
            _sToLocContact = String.Empty
            _sToLocPhone = String.Empty
            _sToLocEmail = String.Empty
            _sSentByName = String.Empty
            _sPONumber = String.Empty
            _sThirdPartyBillNum = String.Empty

            'Data Derived from File Name
            Dim sb As New StringBuilder
            '' Extract the TranDate from the File's Name and Validate it
            Dim sTranDate As String
            Dim iFileNameLength As Integer = p_sFileName.Length

            sTranDate = p_sFileName.Substring(iFileNameLength - 15, 8)

            sb.Append(sTranDate.Substring(4, 2)) ' MM
            sb.Append("/")
            sb.Append(sTranDate.Substring(6, 2)) ' DD
            sb.Append("/")
            sb.Append(sTranDate.Substring(0, 4)) ' YYYY
            sb.Append(" 00:00:00")

            TranDate = CDate(sb.ToString())

        Catch ex As Exception

            Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

        End Try

    End Sub

    Private Sub PrepareForAutoZoneRecords(ByVal p_sFileName As String) ' V4 - Version used by AutoZone

        Try

            _iLineNum = 0

            'Defalut Values
            _chrDelimiter = Chr(126)  'Tilde
            _iExpectedFieldCount = 6
            _strErrorMessage = ""
            _bHasError = False
            _oSpecialHandle = New SpecialHandleColumn

            'Fields that are constant
            _sFromCustId = "13017"
            _sFromCustName = "EXPAK LOGISTICS"
            _sToCustId = _sFromCustId
            _sToCustName = _sFromCustName
            _sBillType = "ACCOUNT"
            _sBillNum = _sFromCustId
            _sVoid = "F"

            'Fields that will be imported from the File
            _sFromLocId = String.Empty
            _sToLocId = String.Empty
            _dtTranDate = Date.Now 'TranDate = CDate(sb.ToString())
            _sPieces = String.Empty
            _sCartonCode = String.Empty
            _sServiceLevel = String.Empty

            'Fields that will be dynamically generated
            _oTrackingNumber = New Barcode
            _iFromAddRowId = 0
            _sFromLocName = String.Empty '
            _sFromLocStreet = String.Empty
            _sFromLocAddress2 = String.Empty
            _sFromLocCity = String.Empty
            _sFromLocState = String.Empty
            _sFromLocZip = String.Empty
            _sToLocName = String.Empty
            _sToLocStreet = String.Empty
            _sToLocCity = String.Empty
            _sToLocZip = String.Empty
            _sToLocState = String.Empty

            'Fields that will remain empty
            _sOrderId = String.Empty
            _sFromLocPhone = String.Empty
            _sFromLocContact = String.Empty
            _sFromLocEmail = String.Empty
            _sToLocAddress2 = String.Empty
            _sToLocContact = String.Empty
            _sToLocPhone = String.Empty
            _sToLocEmail = String.Empty
            _sSentByName = String.Empty
            _sPONumber = String.Empty
            _sThirdPartyBillNum = String.Empty
            _sReferenceNumber = String.Empty
            _fWeight = 0D

        Catch ex As Exception

            Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

        End Try

    End Sub

    'Private Sub PrepareForDSCRecords(ByVal p_sFileName) ' V5 - Version used by DSC Delivery (Rivers)

    '    Try

    '        _iLineNum = 0

    '        'Defalut Values
    '        _chrDelimiter = Chr(126)  'Tilde
    '        _fWeight = 0D
    '        _iExpectedFieldCount = 12
    '        _strErrorMessage = ""
    '        _bHasError = False
    '        _oSpecialHandle = New SpecialHandleColumn

    '        _sFromCustId = "13025"
    '        _sFromCustName = "DSC DELIVERY"
    '        _sToCustId = _sFromCustId
    '        _sToCustName = _sFromCustName
    '        _iFromAddRowId = 52503
    '        _sFromLocId = "001"
    '        _sFromLocName = "RIVER`S EDGE PHARMACY"
    '        _sFromLocStreet = "36919 COOK ST"
    '        _sFromLocAddress2 = "STE 102"
    '        _sFromLocCity = "PALM DESERT"
    '        _sFromLocState = "CA"
    '        _sFromLocZip = "92211"
    '        _sToCustId = "13025"
    '        '_sPieces = "1"
    '        _sCartonCode = "BOX"
    '        _sBillType = "ACCOUNT"
    '        _sBillNum = "13025"

    '        _oTrackingNumber = New Barcode
    '        '_sServiceLevel = "OVERNIGHT"
    '        _sServiceLevel = String.Empty
    '        _sOrderId = String.Empty

    '        _sFromLocPhone = String.Empty
    '        _sFromLocContact = String.Empty
    '        _sFromLocEmail = String.Empty
    '        _sToLocId = String.Empty
    '        _sToLocAddress2 = String.Empty
    '        _sToLocContact = String.Empty
    '        _sToLocPhone = String.Empty
    '        _sToLocEmail = String.Empty
    '        _sSentByName = String.Empty
    '        _sVoid = "F"
    '        _sReferenceNumber = String.Empty
    '        _sPONumber = String.Empty
    '        _sThirdPartyBillNum = String.Empty


    '        'Data Derived from File Name
    '        Dim sb As New StringBuilder

    '        '' Extract the TranDate from the File's Name and Validate it
    '        Dim sTranDate As String
    '        Dim iFileNameLength As Integer = p_sFileName.Length


    '        sTranDate = p_sFileName.Substring(iFileNameLength - 15, 8)

    '        'sb.Append(p_sFileName.Substring(9, 2)) ' MM
    '        'sb.Append("/")
    '        'sb.Append(p_sFileName.Substring(11, 2)) ' DD
    '        'sb.Append("/")
    '        'sb.Append(p_sFileName.Substring(5, 4)) ' YYYY
    '        'sb.Append(" ")
    '        'sb.Append(p_sFileName.Substring(13, 2)) ' hh
    '        'sb.Append(":")
    '        'sb.Append(p_sFileName.Substring(15, 2)) ' mm
    '        'sb.Append(":")
    '        'sb.Append(p_sFileName.Substring(17, 2)) ' ss

    '        sb.Append(sTranDate.Substring(4, 2)) ' MM
    '        sb.Append("/")
    '        sb.Append(sTranDate.Substring(6, 2)) ' DD
    '        sb.Append("/")
    '        sb.Append(sTranDate.Substring(0, 4)) ' YYYY
    '        sb.Append(" 00:00:00")

    '        TranDate = CDate(sb.ToString())

    '    Catch ex As Exception

    '        Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

    '    End Try

    'End Sub

    'Private Sub PrepareForDSCRecords(ByVal p_sFileName) ' V5 - Version used by DSC Delivery (Oso Irvine)

    '    Try

    '        _iLineNum = 0

    '        'Defalut Values
    '        _chrDelimiter = Chr(126)  'Tilde
    '        _fWeight = 0D
    '        _iExpectedFieldCount = 14
    '        _strErrorMessage = ""
    '        _bHasError = False
    '        _oSpecialHandle = New SpecialHandleColumn

    '        _sFromCustId = "13025A"
    '        '_sFromCustId = "13025B"
    '        _sFromCustName = "DSC DELIVERY"
    '        _sToCustId = _sFromCustId
    '        _sToCustName = _sFromCustName
    '        _iFromAddRowId = 53162
    '        _sFromLocId = "001"
    '        _sFromLocName = "OSO HOMECARE"
    '        _sFromLocStreet = "17175 GILLETTE AVE"
    '        _sFromLocAddress2 = String.Empty
    '        _sFromLocCity = "IRVINE"
    '        _sFromLocState = "CA"
    '        _sFromLocZip = "92614"
    '        _sToCustId = "13025A"
    '        '_sFromLocStreet = "1214 BURBANK BLVD"
    '        '_sFromLocAddress2 = String.Empty
    '        '_sFromLocCity = "BURBANK"
    '        '_sFromLocState = "CA"
    '        '_sFromLocZip = "91506"
    '        '_sToCustId = "13025B"
    '        ''_sPieces = "1"
    '        _sCartonCode = "BOX"
    '        _sBillType = "ACCOUNT"
    '        _sBillNum = "13025A"
    '        '_sBillNum = "13025B"

    '        _oTrackingNumber = New Barcode
    '        '_sServiceLevel = "OVERNIGHT"
    '        _sServiceLevel = String.Empty
    '        _sOrderId = String.Empty

    '        _sFromLocPhone = String.Empty
    '        _sFromLocContact = String.Empty
    '        _sFromLocEmail = String.Empty
    '        _sToLocId = String.Empty
    '        _sToLocAddress2 = String.Empty
    '        _sToLocContact = String.Empty
    '        _sToLocPhone = String.Empty
    '        _sToLocEmail = String.Empty
    '        _sSentByName = String.Empty
    '        _sVoid = "F"
    '        _sReferenceNumber = String.Empty
    '        _sPONumber = String.Empty
    '        _sThirdPartyBillNum = String.Empty


    '        'Data Derived from File Name
    '        Dim sb As New StringBuilder

    '        '' Extract the TranDate from the File's Name and Validate it
    '        Dim sTranDate As String
    '        Dim iFileNameLength As Integer = p_sFileName.Length


    '        sTranDate = p_sFileName.Substring(iFileNameLength - 15, 8)

    '        'sb.Append(p_sFileName.Substring(9, 2)) ' MM
    '        'sb.Append("/")
    '        'sb.Append(p_sFileName.Substring(11, 2)) ' DD
    '        'sb.Append("/")
    '        'sb.Append(p_sFileName.Substring(5, 4)) ' YYYY
    '        'sb.Append(" ")
    '        'sb.Append(p_sFileName.Substring(13, 2)) ' hh
    '        'sb.Append(":")
    '        'sb.Append(p_sFileName.Substring(15, 2)) ' mm
    '        'sb.Append(":")
    '        'sb.Append(p_sFileName.Substring(17, 2)) ' ss

    '        sb.Append(sTranDate.Substring(4, 2)) ' MM
    '        sb.Append("/")
    '        sb.Append(sTranDate.Substring(6, 2)) ' DD
    '        sb.Append("/")
    '        sb.Append(sTranDate.Substring(0, 4)) ' YYYY
    '        sb.Append(" 00:00:00")

    '        TranDate = CDate(sb.ToString())

    '    Catch ex As Exception

    '        Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

    '    End Try

    'End Sub

    Private Sub PrepareForDSCRecords(ByVal p_sFileName As String) ' V5 - Version used by DSC Delivery (Oso Burbank & Irvine)

        Try

            _iLineNum = 0

            'Defalut Values
            _chrDelimiter = Chr(126)  'Tilde
            _fWeight = 0D
            _iExpectedFieldCount = 14
            _strErrorMessage = ""
            _bHasError = False
            _oSpecialHandle = New SpecialHandleColumn

            '_sFromCustId = "13025A"
            '_sFromCustId = "13025B"
            _sFromCustId = String.Empty

            '_sFromCustName = "DSC DELIVERY"
            _sFromCustName = String.Empty

            '_sToCustId = _sFromCustId
            _sToCustId = String.Empty

            _sToCustName = _sFromCustName

            '_iFromAddRowId = 53162
            _iFromAddRowId = 0

            '_sFromLocId = "001"
            _sFromLocId = String.Empty

            '_sFromLocName = "OSO HOMECARE"
            _sFromLocName = String.Empty

            '_sFromLocStreet = "17175 GILLETTE AVE"
            _sFromLocStreet = String.Empty

            _sFromLocAddress2 = String.Empty
            '_sFromLocCity = "IRVINE"
            _sFromLocCity = String.Empty

            '_sFromLocState = "CA"
            _sFromLocState = String.Empty

            '_sFromLocZip = "92614"
            _sFromLocZip = String.Empty

            '_sToCustId = "13025A"
            _sToCustId = String.Empty


            '_sFromLocStreet = "1214 BURBANK BLVD"
            _sFromLocStreet = String.Empty

            _sFromLocAddress2 = String.Empty

            '_sFromLocCity = "BURBANK"
            _sFromLocCity = String.Empty

            '_sFromLocState = "CA"
            '_sFromLocZip = "91506"
            '_sToCustId = "13025B"
            ''_sPieces = "1"
            _sPieces = String.Empty

            _sCartonCode = "BOX"
            _sBillType = "ACCOUNT"

            '_sBillNum = "13025A"
            '_sBillNum = "13025B"
            _sBillNum = String.Empty


            _oTrackingNumber = New Barcode
            '_sServiceLevel = "OVERNIGHT"
            _sServiceLevel = String.Empty
            _sOrderId = String.Empty

            _sFromLocPhone = String.Empty
            _sFromLocContact = String.Empty
            _sFromLocEmail = String.Empty
            _sToLocId = String.Empty
            _sToLocAddress2 = String.Empty
            _sToLocContact = String.Empty
            _sToLocPhone = String.Empty
            _sToLocEmail = String.Empty
            _sSentByName = String.Empty
            _sVoid = "F"
            _sReferenceNumber = String.Empty
            _sPONumber = String.Empty
            _sThirdPartyBillNum = String.Empty


            'Data Derived from File Name
            Dim sb As New StringBuilder

            '' Extract the TranDate from the File's Name and Validate it
            Dim sTranDate As String
            Dim iFileNameLength As Integer = p_sFileName.Length


            sTranDate = p_sFileName.Substring(iFileNameLength - 15, 8)

            'sb.Append(p_sFileName.Substring(9, 2)) ' MM
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(11, 2)) ' DD
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(5, 4)) ' YYYY
            'sb.Append(" ")
            'sb.Append(p_sFileName.Substring(13, 2)) ' hh
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(15, 2)) ' mm
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(17, 2)) ' ss

            sb.Append(sTranDate.Substring(4, 2)) ' MM
            sb.Append("/")
            sb.Append(sTranDate.Substring(6, 2)) ' DD
            sb.Append("/")
            sb.Append(sTranDate.Substring(0, 4)) ' YYYY
            sb.Append(" 00:00:00")

            TranDate = CDate(sb.ToString())

        Catch ex As Exception

            Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

        End Try

    End Sub

    Private Sub PrepareForPIARecords(ByVal p_sFileName As String) ' V6 - Version used by CALPIA

        Try

            _iLineNum = 0

            'Defalut Values
            _chrDelimiter = Chr(126)  'Tilde
            _fWeight = 0D
            _iExpectedFieldCount = 16
            _strErrorMessage = ""
            _bHasError = False
            _oSpecialHandle = New SpecialHandleColumn

            'Hard-coded Values
            _sFromLocId = "001"
            _sFromLocAddress2 = String.Empty
            _sPieces = "1"
            _sCartonCode = "BOX"
            _sBillType = "ACCOUNT"
            _sBillNum = String.Empty
            _sServiceLevel = "NEXT DAY"
            _sOrderId = String.Empty
            _sFromLocContact = String.Empty
            _sFromLocEmail = String.Empty
            _sToLocContact = String.Empty
            _sToLocPhone = String.Empty
            _sToLocEmail = String.Empty
            _sSentByName = String.Empty
            _sVoid = "F"
            _sReferenceNumber = String.Empty
            _sPONumber = String.Empty
            _sThirdPartyBillNum = String.Empty

            'Values Dependent on Values of Other Fields
            _sFromCustId = String.Empty 'Will either be 14007 (Chowchilla) or 14008 (Vacaville)
            _sFromCustName = String.Empty 'Will either be "PIA CHOWCHILLA" or "PIA VACAVILLE"
            _sToCustId = String.Empty 'Will either be 14007 (Chowchilla) or 14008 (Vacaville)
            _sToCustName = String.Empty 'Will either be "PIA CHOWCHILLA" or "PIA VACAVILLE"
            _iFromAddRowId = 0 'Wlll either be 168796 (Chowchilla) or 168797 (Vacaville)
            _sFromLocName = String.Empty 'Will either be "PIA CHOWCHILLA" or "CSP SOLANO"
            _sFromLocStreet = String.Empty 'Will either be "23370 ROAD 22" (Chowchilla) or "2100 PEABODY ROAD" (CSP Solano)
            _sFromLocZip = String.Empty 'Will either be "93610" (Chowchilla) or "95687"


            'Values Read Directly From File
            _sFromLocCity = String.Empty
            _sFromLocState = String.Empty
            _sFromLocPhone = String.Empty
            _sToLocName = String.Empty
            _sToLocStreet = String.Empty
            _sToLocAddress2 = String.Empty
            _sToLocCity = String.Empty
            _sToLocState = String.Empty
            _sToLocZip = String.Empty
            _oTrackingNumber = New Barcode
            _sToLocId = String.Empty
            _dtTranDate = Date.Now

            'Data Derived from File Name
            Dim sb As New StringBuilder

            '' Extract the TranDate from the File's Name and Validate it
            'Dim sTranDate As String
            'Dim iFileNameLength As Integer = p_sFileName.Length


            'sTranDate = p_sFileName.Substring(iFileNameLength - 15, 8)

            ''sb.Append(p_sFileName.Substring(9, 2)) ' MM
            ''sb.Append("/")
            ''sb.Append(p_sFileName.Substring(11, 2)) ' DD
            ''sb.Append("/")
            ''sb.Append(p_sFileName.Substring(5, 4)) ' YYYY
            ''sb.Append(" ")
            ''sb.Append(p_sFileName.Substring(13, 2)) ' hh
            ''sb.Append(":")
            ''sb.Append(p_sFileName.Substring(15, 2)) ' mm
            ''sb.Append(":")
            ''sb.Append(p_sFileName.Substring(17, 2)) ' ss

            'sb.Append(sTranDate.Substring(4, 2)) ' MM
            'sb.Append("/")
            'sb.Append(sTranDate.Substring(6, 2)) ' DD
            'sb.Append("/")
            'sb.Append(sTranDate.Substring(0, 4)) ' YYYY
            'sb.Append(" 00:00:00")

            'TranDate = CDate(sb.ToString())

        Catch ex As Exception

            Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

        End Try

    End Sub

    Private Sub PrepareForProCourierRecords(ByVal p_sFileName As String) ' V7 - Version used by ProCourier

        Try

            _iLineNum = 0

            'Defalut Values
            _chrDelimiter = Chr(126)  'Tilde
            _fWeight = 0D
            _iExpectedFieldCount = 13
            _strErrorMessage = ""
            _bHasError = False
            _oSpecialHandle = New SpecialHandleColumn
            _sFromCustId = "14016"
            _sFromCustName = "PRO COURIER"
            _sToCustId = _sFromCustId
            _sToCustName = _sFromCustName
            _sCartonCode = "PIECE"
            _sBillType = "ACCOUNT"
            _sBillNum = _sFromCustId
            _sVoid = "F"

            'Initialize Non-default Values
            _oTrackingNumber = New Barcode
            _sServiceLevel = String.Empty
            _sOrderId = String.Empty
            _iFromAddRowId = 0
            _sFromLocId = String.Empty
            _sFromLocName = String.Empty
            _sFromLocStreet = String.Empty
            _sFromLocAddress2 = String.Empty
            _sFromLocCity = String.Empty
            _sFromLocState = String.Empty
            _sFromLocZip = String.Empty
            _sFromLocPhone = String.Empty
            _sFromLocContact = String.Empty
            _sFromLocEmail = String.Empty
            _sToLocId = String.Empty
            _sToLocAddress2 = String.Empty
            _sToLocContact = String.Empty
            _sToLocPhone = String.Empty
            _sToLocEmail = String.Empty
            _sSentByName = String.Empty
            _sReferenceNumber = String.Empty
            _sPONumber = String.Empty
            _sThirdPartyBillNum = String.Empty


            'Data Derived from File Name
            Dim sb As New StringBuilder

            '' Extract the TranDate from the File's Name and Validate it
            Dim sTranDate As String
            Dim iFileNameLength As Integer = p_sFileName.Length


            sTranDate = p_sFileName.Substring(iFileNameLength - 15, 8)

            'sb.Append(p_sFileName.Substring(9, 2)) ' MM
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(11, 2)) ' DD
            'sb.Append("/")
            'sb.Append(p_sFileName.Substring(5, 4)) ' YYYY
            'sb.Append(" ")
            'sb.Append(p_sFileName.Substring(13, 2)) ' hh
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(15, 2)) ' mm
            'sb.Append(":")
            'sb.Append(p_sFileName.Substring(17, 2)) ' ss

            sb.Append(sTranDate.Substring(4, 2)) ' MM
            sb.Append("/")
            sb.Append(sTranDate.Substring(6, 2)) ' DD
            sb.Append("/")
            sb.Append(sTranDate.Substring(0, 4)) ' YYYY
            sb.Append(" 00:00:00")

            TranDate = CDate(sb.ToString())

        Catch ex As Exception

            Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

        End Try

    End Sub

    Private Sub PrepareForDIO2CSRecords(ByVal p_sFileName As String) ' V8 - Version used by Deliver-it to Courier Systems

        Try

            _iLineNum = 0

            'Defalut Values
            _chrDelimiter = Chr(126)  'Tilde
            _iExpectedFieldCount = 17
            _strErrorMessage = ""
            _bHasError = False
            _oSpecialHandle = New SpecialHandleColumn

            _sFromCustId = "13024A"
            _sFromCustName = "DELIVER-IT OVERNIGHT LLC"
            _sToCustId = _sFromCustId
            _sToCustName = _sFromCustName
            _sFromLocState = "CA"
            _sBillType = "ACCOUNT"
            _sBillNum = _sFromCustId
            _sServiceLevel = "OVERNIGHT"
            _sVoid = "F"

            _iFromAddRowId = 0
            _sFromLocId = String.Empty
            _sFromLocName = String.Empty
            _sFromLocStreet = String.Empty
            _sFromLocAddress2 = String.Empty
            _sFromLocCity = String.Empty
            _sFromLocZip = String.Empty
            _sPieces = String.Empty
            _sCartonCode = String.Empty
            _oTrackingNumber = New Barcode
            _sOrderId = String.Empty
            _sFromLocPhone = String.Empty
            _sFromLocContact = String.Empty
            _sFromLocEmail = String.Empty
            _sToLocId = String.Empty
            _sToLocAddress2 = String.Empty
            _sToLocContact = String.Empty
            _sToLocPhone = String.Empty
            _sToLocEmail = String.Empty
            _sSentByName = String.Empty
            _sReferenceNumber = String.Empty
            _sPONumber = String.Empty
            _sThirdPartyBillNum = String.Empty


            'Data Derived from File Name
            Dim sb As New StringBuilder

            '' Extract the TranDate from the File's Name and Validate it
            Dim sTranDate As String
            Dim iFileNameLength As Integer = p_sFileName.Length


            sTranDate = p_sFileName.Substring(iFileNameLength - 15, 8)

            sb.Append(sTranDate.Substring(4, 2)) ' MM
            sb.Append("/")
            sb.Append(sTranDate.Substring(6, 2)) ' DD
            sb.Append("/")
            sb.Append(sTranDate.Substring(0, 4)) ' YYYY
            sb.Append(" 00:00:00")

            TranDate = CDate(sb.ToString())

        Catch ex As Exception

            Throw New ArgumentException("Invalid FileName Specified. " & p_sFileName)

        End Try

    End Sub


    'Default Constructor
    Public Sub New(ByVal p_sFileName As String, Optional ByVal p_sFileVersion As String = "V0")

        _sFileFormat = p_sFileVersion

        If String.Compare(_sFileFormat, "V0") = 0 Then
            PrepareForIngramMicroRecords(p_sFileName)
        End If

        If String.Compare(_sFileFormat, "V1") = 0 Then
            PrepareForPartnerBooksRecords(p_sFileName)
        End If

        If String.Compare(_sFileFormat, "V2") = 0 Then
            PrepareForMedExRecords(p_sFileName)
        End If

        If String.Compare(_sFileFormat, "V2s") = 0 Then
            PrepareForMedExSdRecords(p_sFileName)
        End If

        If String.Compare(_sFileFormat, "V3") = 0 Then
            PrepareForBluePackageRecords(p_sFileName)
        End If

        If String.Compare(_sFileFormat, "V4") = 0 Then
            PrepareForAutoZoneRecords(p_sFileName)
        End If

        If String.Compare(_sFileFormat, "V5") = 0 Then
            PrepareForDSCRecords(p_sFileName)
        End If

        If String.Compare(_sFileFormat, "V6") = 0 Then
            PrepareForPIARecords(p_sFileName)
        End If

        If String.Compare(_sFileFormat, "V7") = 0 Then
            PrepareForProCourierRecords(p_sFileName)
        End If

        If String.Compare(_sFileFormat, "V8") = 0 Then
            PrepareForDIO2CSRecords(p_sFileName)
        End If

    End Sub

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

    Private Sub SetError(ByVal p_errMsg As String)

        _bHasError = True
        _strErrorMessage = p_errMsg

    End Sub

    'Properties to expose and format records fields
    Public Property Delimiter() As String
        Get
            Return _chrDelimiter
        End Get
        Set(ByVal Value As String)
            _chrDelimiter = Value
        End Set
    End Property

    Public ReadOnly Property BarcodeName() As String
        Get
            If Not IsNothing(_oTrackingNumber) Then
                Return _oTrackingNumber.BarcodeName
            Else
                Return ""
            End If
        End Get
    End Property

    Public Property TrackingNumber() As String
        Get
            Return _oTrackingNumber.Barcode
        End Get
        Set(ByVal value As String)
            _oTrackingNumber = Nothing
            _oTrackingNumber = NewBarcodeObject(value)
        End Set
    End Property

    Public Property OrderId() As String
        Get
            Return _sOrderId
        End Get
        Set(ByVal value As String)
            If value.Length > 40 Then Throw New ArgumentOutOfRangeException("OrderId", "Order Id Cannot Exceed 40 Characters [" & value & "]") Else _sOrderId = value
        End Set
    End Property

    Public Property FromCustId() As String
        Get
            Return _sFromCustId
        End Get
        Set(ByVal value As String)
            If value.Length > 10 Then Throw New ArgumentOutOfRangeException("FromCustId", "Customer Id Cannot Exceed 10 Characters [" & value & "]") Else _sFromCustId = value
        End Set
    End Property

    Public Property FromCustName() As String
        Get
            Return _sFromCustName
        End Get
        Set(ByVal value As String)
            If value.Length > 70 Then Throw New ArgumentOutOfRangeException("FromCustName", "Customer Name Cannot Exceed 70 Characters [" & value & "]") Else _sFromCustName = value
        End Set
    End Property

    Public Property FromAddRowId() As Integer
        Get
            Return _iFromAddRowId
        End Get
        Set(ByVal value As Integer)
            If (value < 0 Or value > 2147483647) Then Throw New ArgumentOutOfRangeException("FromAddRowId", "Address Id must be between 0 and 2,147,483,647") Else _iFromAddRowId = value
        End Set
    End Property

    Public Property FromLocId() As String   '10
        Get
            Return _sFromLocId
        End Get
        Set(ByVal value As String)
            If value.Length > 10 Then Throw New ArgumentOutOfRangeException("FromLocId", "Customer Name Cannot Exceed 10 Characters [" & value & "]") Else _sFromLocId = value
        End Set
    End Property

    Public Property FromLocName() As String '70
        Get
            Return _sFromLocName
        End Get
        Set(ByVal value As String)
            If value.Length > 70 Then Throw New ArgumentOutOfRangeException("FromLocName", "Cannot Exceed 70 Characters [" & value & "]") Else _sFromLocName = value
        End Set
    End Property

    Public Property FromLocStreet() As String   '40
        Get
            Return _sFromLocStreet
        End Get
        Set(ByVal value As String)
            If value.Length > 40 Then Throw New ArgumentOutOfRangeException("FromLocStreet", "Cannot Exceed 40 Characters [" & value & "]") Else _sFromLocStreet = value
        End Set
    End Property

    Public Property FromLocAddress2() As String '30
        Get
            Return _sFromLocAddress2
        End Get
        Set(ByVal value As String)
            If value.Length > 30 Then Throw New ArgumentOutOfRangeException("FromLocAddress2", "Cannot Exceed 30 Characters [" & value & "]") Else _sFromLocAddress2 = value
        End Set
    End Property

    Public Property FromLocCity() As String '32
        Get
            Return _sFromLocCity
        End Get
        Set(ByVal value As String)
            If value.Length > 32 Then Throw New ArgumentOutOfRangeException("FromLocCity", "Cannot Exceed 32 Characters [" & value & "]") Else _sFromLocCity = value
        End Set
    End Property

    Public Property FromLocState() As String    '2
        Get
            Return _sFromLocState
        End Get
        Set(ByVal value As String)
            If value.Length > 2 Then Throw New ArgumentOutOfRangeException("FromLocState", "Cannot Exceed 2 Characters [" & value & "]") Else _sFromLocState = value
        End Set
    End Property

    Public Property FromLocZip() As String '10
        Get
            Return _sFromLocZip
        End Get
        Set(ByVal value As String)
            If value.Length > 10 Then Throw New ArgumentOutOfRangeException("FromLocZip", "Cannot Exceed 10 Characters [" & value & "]") Else _sFromLocZip = value
        End Set
    End Property

    Public Property FromLocPhone() As String '20
        Get
            Return _sFromLocPhone
        End Get
        Set(ByVal value As String)
            If value.Length > 20 Then Throw New ArgumentOutOfRangeException("FromLocPhone", "Cannot Exceed 20 Characters [" & value & "]") Else _sFromLocPhone = value
        End Set
    End Property

    Public Property FromLocContact() As String '40
        Get
            Return _sFromLocContact
        End Get
        Set(ByVal value As String)
            If value.Length > 40 Then Throw New ArgumentOutOfRangeException("FromLocContact", "Cannot Exceed 40 Characters [" & value & "]") Else _sFromLocContact = value
        End Set
    End Property

    Public Property FromLocEmail() As String '30
        Get
            Return _sFromLocEmail
        End Get
        Set(ByVal value As String)
            If value.Length > 30 Then Throw New ArgumentOutOfRangeException("FromLocEmail", "Cannot Exceed 30 Characters [" & value & "]") Else _sFromLocEmail = value
        End Set
    End Property

    Public Property ToAddRowId() As Integer
        Get
            Return _iToAddRowId
        End Get
        Set(ByVal value As Integer)
            If (value < 0 Or value > 2147483647) Then Throw New ArgumentOutOfRangeException("ToAddRowId", "Address Id must be between 0 and 2,147,483,647") Else _iToAddRowId = value
        End Set
    End Property

    Public Property ToCustId() As String '10
        Get
            Return _sToCustId
        End Get
        Set(ByVal value As String)
            If value.Length > 10 Then Throw New ArgumentOutOfRangeException("ToCustId", "Cannot Exceed 10 Characters [" & value & "]") Else _sToCustId = value
        End Set
    End Property

    Public Property ToCustName() As String '70
        Get
            Return _sToCustName
        End Get
        Set(ByVal value As String)
            If value.Length > 70 Then Throw New ArgumentOutOfRangeException("ToCustName", "Cannot Exceed 70 Characters [" & value & "]") Else _sToCustName = value
        End Set
    End Property

    Public Property ToLocId() As String '10
        Get
            Return _sToLocId
        End Get
        Set(ByVal value As String)
            If value.Length > 10 Then Throw New ArgumentOutOfRangeException("ToLocId", "Cannot Exceed 10 Characters [" & value & "]") Else _sToLocId = value
        End Set
    End Property

    Public Property ToLocName() As String '70
        Get
            Return _sToLocName
        End Get
        Set(ByVal value As String)
            If value.Length > 70 Then Throw New ArgumentOutOfRangeException("ToLocName", "Cannot Exceed 70 Characters [" & value & "]") Else _sToLocName = value
        End Set
    End Property

    Public Property ToLocStreet() As String '40
        Get
            Return _sToLocStreet
        End Get
        Set(ByVal value As String)
            If value.Length > 40 Then Throw New ArgumentOutOfRangeException("ToLocStreet", "Cannot Exceed 40 Characters [" & value & "]") Else _sToLocStreet = value
        End Set
    End Property

    Public Property ToLocAddress2() As String '30
        Get
            Return _sToLocAddress2
        End Get
        Set(ByVal value As String)
            If value.Length > 30 Then Throw New ArgumentOutOfRangeException("ToLocAddress2", "Cannot Exceed 30 Characters [" & value & "]") Else _sToLocAddress2 = value
        End Set
    End Property

    Public Property ToLocCity() As String '32
        Get
            Return _sToLocCity
        End Get
        Set(ByVal value As String)
            If value.Length > 32 Then Throw New ArgumentOutOfRangeException("ToLocCity", "Cannot Exceed 32 Characters [" & value & "]") Else _sToLocCity = value
        End Set
    End Property

    Public Property ToLocState() As String '2
        Get
            Return _sToLocState
        End Get
        Set(ByVal value As String)
            If value.Length > 2 Then Throw New ArgumentOutOfRangeException("ToLocState", "Cannot Exceed 2 Characters [" & value & "]") Else _sToLocState = value
        End Set
    End Property

    Public Property ToLocZip() As String '10
        Get
            Return _sToLocZip
        End Get
        Set(ByVal value As String)
            If value.Length > 10 Then Throw New ArgumentOutOfRangeException("ToLocZip", "Cannot Exceed 10 Characters [" & value & "]") Else _sToLocZip = value
        End Set
    End Property

    Public Property ToLocContact() As String '40
        Get
            Return _sToLocContact
        End Get
        Set(ByVal value As String)
            If value.Length > 40 Then Throw New ArgumentOutOfRangeException("ToLocContact", "Cannot Exceed 40 Characters [" & value & "]") Else _sToLocContact = value
        End Set
    End Property

    Public Property ToLocPhone() As String '20
        Get
            Return _sToLocPhone
        End Get
        Set(ByVal value As String)
            If value.Length > 20 Then Throw New ArgumentOutOfRangeException("ToLocPhone", "Cannot Exceed 20 Characters [" & value & "]") Else _sToLocPhone = value
        End Set
    End Property

    Public Property ToLocEmail() As String '30
        Get
            Return _sToLocEmail
        End Get
        Set(ByVal value As String)
            If value.Length > 30 Then Throw New ArgumentOutOfRangeException("ToLocEmail", "Cannot Exceed 30 Characters [" & value & "]") Else _sToLocEmail = value
        End Set
    End Property

    Public Property Weight() As Decimal 'dec(6,2)
        Get
            Return _fWeight
        End Get
        Set(ByVal value As Decimal)
            If value < 0 Then Throw New ArgumentException("Weight", "Weight cannot be negative") Else _fWeight = value
        End Set
    End Property

    Public Property Pieces() As String '10
        Get
            Return _sPieces
        End Get
        Set(ByVal value As String)
            If value.Length > 10 Then Throw New ArgumentOutOfRangeException("Pieces", "Cannot Exceed 30 Characters [" & value & "]") Else _sPieces = value
        End Set
    End Property

    Public Property SentByName() As String '30
        Get
            Return _sSentByName
        End Get
        Set(ByVal value As String)
            If value.Length > 30 Then Throw New ArgumentOutOfRangeException("SentByName", "Cannot Exceed 30 Characters [" & value & "]") Else _sSentByName = value
        End Set
    End Property

    Public Property CartonCode() As String '20
        Get
            Return _sCartonCode
        End Get
        Set(ByVal value As String)
            If IngramMicroCartonCodes.ValidCode(value) Then _sCartonCode = value Else Throw New ArgumentException("Invalid Carton Code Specified")
        End Set
    End Property

    Public Property Dimensions() As String '26 (shared)
        Get
            Return _oSpecialHandle.Dimensions
        End Get
        Set(ByVal value As String)
            If value.Length > 26 Then
                Throw New ArgumentOutOfRangeException("Dimensions", "Cannot Exceed 26 Characters [" & value & "]")
            Else
                _oSpecialHandle.Dimensions = value
            End If
        End Set
    End Property

    Public Property ServiceLevel() As String '20
        Get
            Return _sServiceLevel
        End Get
        Set(ByVal value As String)
            If value.Length > 20 Then Throw New ArgumentOutOfRangeException("ServiceLevel", "Cannot Exceed 20 Characters [" & value & "]") Else _sServiceLevel = value
        End Set
    End Property

    Public Property BillType() As String '20
        Get
            Return _sBillType
        End Get
        Set(ByVal value As String)
            Dim v As String = Nothing
            If value.Equals(" ") Or value = String.Empty Then v = "ACCOUNT"
            If v.Length > 20 Then Throw New ArgumentOutOfRangeException("BillType", "Cannot Exceed 20 Characters [" & value & "]") Else _sBillType = v
        End Set
    End Property

    Public Property BillNum() As String '50
        Get
            Return _sBillNum
        End Get
        Set(ByVal value As String)
            Dim v As String
            If value.Equals(" ") Or value = String.Empty Then v = CStr(FromCustId) Else v = value
            If v.Length > 50 Then Throw New ArgumentOutOfRangeException("BillNum", "Cannot Exceed 50 Characters [" & value & "]") Else _sBillNum = v
        End Set
    End Property

    Public Property TranDate() As Date
        Get
            Return _dtTranDate
        End Get
        Set(ByVal value As Date)
            _dtTranDate = value
        End Set
    End Property

    Public ReadOnly Property UniqueRecordId() As String '29
        Get
            Return _sUniqueRecordId
        End Get
    End Property

    Public Property Void() As String '1 (T|F)
        Get
            Return _sVoid
        End Get
        Set(ByVal value As String)
            If (String.Compare(value.ToUpper(), "T") = 0) Or (String.Compare(value.ToUpper(), "F") = 0) Then
                _sVoid = value.ToUpper()
            Else
                _sVoid = "F"
                'Throw New ArgumentException("Invalid Value Specified.  Must be either 'Y' or 'N'. [" & value & "]")
            End If
        End Set
    End Property

    Public Property ReferenceNumber() As String '40
        Get
            Return _sReferenceNumber
        End Get
        Set(ByVal value As String)
            If value.Length > 40 Then Throw New ArgumentOutOfRangeException("ReferenceNumber", "Cannot Exceed 40 Characters") Else _sReferenceNumber = value
        End Set
    End Property

    Public Property PONumber() As String '40
        Get
            Return _sPONumber
        End Get
        Set(ByVal value As String)
            If value.Length > 40 Then Throw New ArgumentOutOfRangeException("PONumber", "Cannot Exceed 40 Characters") Else _sPONumber = value
        End Set
    End Property

    Public Property ThirdPartyBillNum() As String '40
        Get
            Return _sThirdPartyBillNum
        End Get
        Set(ByVal value As String)
            If value.Length > 40 Then Throw New ArgumentOutOfRangeException("ThirdPartyBillNum", "Cannot Exceed 40 Characters") Else _sThirdPartyBillNum = value
        End Set
    End Property

    Public Property Modifiers() As String '113 (shared)
        Get
            Return _oSpecialHandle.Modifiers
        End Get
        Set(ByVal value As String)
            If value.Length > 113 Then
                Throw New ArgumentOutOfRangeException("Modifiers", "Cannot Exceed 113 Characters")
            Else
                _oSpecialHandle.Modifiers = value
            End If
        End Set
    End Property

    Public Property DeclaredValue() As Decimal '150 (shared)
        Get
            Return _oSpecialHandle.DeclaredValue
        End Get
        Set(ByVal value As Decimal)
            Dim v As Decimal = Decimal.Round(value, 2)
            If v.ToString("f2").Length > 9 Then
                Throw New ArgumentOutOfRangeException("Declared Value", "Cannot Exceed 999999.99 Characters")
            Else
                _oSpecialHandle.DeclaredValue = v
            End If
        End Set
    End Property

    Public Property EmptyField() As String '0
        Get
            Return Nothing
        End Get
        Set(ByVal Value As String)
            _sEmptyField = Nothing
        End Set
    End Property

    Public Property LineNumber() As Integer
        Get
            Return _iLineNum
        End Get
        Set(ByVal Value As Integer)
            _iLineNum = Value
        End Set
    End Property

    Public Property ChargedBySub() As Decimal
        'Get
        '    Return _fChargedBySub
        'End Get
        'Set(value As Decimal)
        '    Dim v As Decimal = Decimal.Round(value, 2)
        '    If v.ToString("f2").Length > 9 Then
        '        Throw New ArgumentOutOfRangeException("Charged By Sub", "Cannot Exceed 999999.99")
        '    Else
        '        _fChargedBySub = v
        '    End If
        'End Set
        Get
            Return DeclaredValue
        End Get
        Set(value As Decimal)
            DeclaredValue = value
        End Set
    End Property

    Public Function AssignRecord(ByVal p_strLine As String) As Boolean

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

    Private Function DecomposeIngramMicroRecordString(ByVal p_strSubstrings() As String) As Boolean

        Dim sFromAddRowId As String
        Dim sToAddRowId As String
        Dim sWeight As String
        Dim sDeclaredValue As String

        Try
            TrackingNumber = p_strSubstrings(FieldsV0.TrackingNumber)
            OrderId = p_strSubstrings(FieldsV0.OrderId)
            FromCustId = p_strSubstrings(FieldsV0.FromCustId)
            FromCustName = p_strSubstrings(FieldsV0.FromCustName)
            sFromAddRowId = p_strSubstrings(FieldsV0.FromAddRowId)
            If IsNumeric(sFromAddRowId) Then FromAddRowId = CInt(sFromAddRowId) Else FromAddRowId = 0
            FromLocId = p_strSubstrings(FieldsV0.FromLocId)
            FromLocName = p_strSubstrings(FieldsV0.FromLocName)
            FromLocStreet = p_strSubstrings(FieldsV0.FromLocStreet)
            FromLocAddress2 = p_strSubstrings(FieldsV0.FromLocAddress2)
            FromLocCity = p_strSubstrings(FieldsV0.FromLocCity)
            FromLocState = p_strSubstrings(FieldsV0.FromLocState)
            FromLocZip = p_strSubstrings(FieldsV0.FromLocZip)
            FromLocPhone = p_strSubstrings(FieldsV0.FromLocPhone)
            FromLocContact = p_strSubstrings(FieldsV0.FromLocContact)
            FromLocEmail = p_strSubstrings(FieldsV0.FromLocEmail)
            sToAddRowId = p_strSubstrings(FieldsV0.ToAddRowId)
            If IsNumeric(sToAddRowId) Then ToAddRowId = CInt(sToAddRowId) Else ToAddRowId = 0
            ToCustId = p_strSubstrings(FieldsV0.ToCustId)
            ToCustName = p_strSubstrings(FieldsV0.ToCustName)
            ToLocId = p_strSubstrings(FieldsV0.ToLocId)
            ToLocName = p_strSubstrings(FieldsV0.ToLocName)
            ToLocStreet = p_strSubstrings(FieldsV0.ToLocStreet)
            ToLocAddress2 = p_strSubstrings(FieldsV0.ToLocAddress2)
            ToLocCity = p_strSubstrings(FieldsV0.ToLocCity)
            ToLocState = p_strSubstrings(FieldsV0.ToLocState)
            ToLocZip = p_strSubstrings(FieldsV0.ToLocZip)
            ToLocContact = p_strSubstrings(FieldsV0.ToLocContact)
            ToLocPhone = p_strSubstrings(FieldsV0.ToLocPhone)
            ToLocEmail = p_strSubstrings(FieldsV0.ToLocEmail)
            sWeight = p_strSubstrings(FieldsV0.Weight)
            If IsNumeric(sWeight) Then Weight = CDec(sWeight) Else Weight = 0D
            Pieces = p_strSubstrings(FieldsV0.Pieces)
            SentByName = p_strSubstrings(FieldsV0.SentByName)
            CartonCode = p_strSubstrings(FieldsV0.CartonCode)
            If p_strSubstrings(FieldsV0.Dimensions) = String.Empty Or String.Compare(p_strSubstrings(FieldsV0.Dimensions), " ") = 0 Then
                Dim oCarton As Carton = IngramMicroCartonCodes.GetCartonByCode(CartonCode)
                Dimensions = oCarton.DimensionString
            Else
                Dimensions = p_strSubstrings(FieldsV0.Dimensions)
            End If
            ServiceLevel = p_strSubstrings(FieldsV0.ServiceLevel)
            BillType = p_strSubstrings(FieldsV0.BillType)
            BillNum = p_strSubstrings(FieldsV0.BillNum)

            'The TranDate is provided via the File's Name. Value should be set explicitly outside of this method.

            '' UniqueRecordId - This value is derived from TrackingNum and TranDate
            Dim sb As New StringBuilder
            sb.Append(TranDate.ToString("MMddyyyyHHmmss", CultureInfo.InvariantCulture))
            sb.Append(TrackingNumber)
            _sUniqueRecordId = sb.ToString()

            Void = p_strSubstrings(FieldsV0.Void)
            ReferenceNumber = p_strSubstrings(FieldsV0.ReferenceNumber)
            PONumber = p_strSubstrings(FieldsV0.PONumber)
            ThirdPartyBillNum = p_strSubstrings(FieldsV0.ThirdPartyBillNum)
            Modifiers = p_strSubstrings(FieldsV0.Modifiers)
            sDeclaredValue = p_strSubstrings(FieldsV0.DeclaredValue)
            If IsNumeric(sDeclaredValue) Then DeclaredValue = CDec(sDeclaredValue) Else DeclaredValue = 0D

            EmptyField = p_strSubstrings(FieldsV0.EmptyField)


        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges

            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function

    Private Function DecomposePartnerBooksRecordString(ByVal p_strSubstrings() As String) As Boolean

        Dim sFromAddRowId As String
        'Dim sToAddRowId As String
        Dim sWeight As String
        Dim sDeclaredValue As String

        Try

            'Toss Header
            If String.Compare(p_strSubstrings(FieldsV1.FromCustId), "ShipperID") = 0 Then
                _bHasError = False
                Return False
            End If

            'TranDate is currently set by the filename.  TO DO:  Change to set from field 0

            FromCustId = p_strSubstrings(FieldsV1.FromCustId)
            FromCustName = p_strSubstrings(FieldsV1.FromCustName)

            sFromAddRowId = "43501" 'CAUTION: Tied to unique RowId of Unison Address table.  TO DO:  Make this dynamic at some point.
            FromLocId = "001"
            FromLocName = FromCustName
            FromLocStreet = p_strSubstrings(FieldsV1.FromLocStreet)
            FromLocAddress2 = p_strSubstrings(FieldsV1.FromLocAddress2)
            FromLocCity = p_strSubstrings(FieldsV1.FromLocCity)
            FromLocState = p_strSubstrings(FieldsV1.FromLocState)
            FromLocZip = p_strSubstrings(FieldsV1.FromLocZip)
            FromLocPhone = String.Empty
            FromLocContact = String.Empty
            FromLocEmail = String.Empty
            SentByName = String.Empty

            ToCustId = FromCustId
            ToCustName = FromCustName
            ToAddRowId = 0 'Not Provided
            ToLocId = p_strSubstrings(FieldsV1.ToLocId)
            ToLocName = p_strSubstrings(FieldsV1.ToLocName)
            ToLocStreet = p_strSubstrings(FieldsV1.ToLocStreet)
            ToLocAddress2 = p_strSubstrings(FieldsV1.ToLocAddress2)
            ToLocCity = p_strSubstrings(FieldsV1.ToLocCity)
            ToLocState = p_strSubstrings(FieldsV1.ToLocState)
            ToLocZip = p_strSubstrings(FieldsV1.ToLocZip)
            ToLocContact = String.Empty
            ToLocPhone = String.Empty
            ToLocEmail = String.Empty

            TrackingNumber = p_strSubstrings(FieldsV1.TrackingNumber)
            OrderId = p_strSubstrings(FieldsV1.OrderId)
            sWeight = p_strSubstrings(FieldsV1.Weight)
            If IsNumeric(sWeight) Then Weight = CDec(sWeight) Else Weight = 0D
            ServiceLevel = p_strSubstrings(FieldsV1.ServiceLevel)
            CartonCode = p_strSubstrings(FieldsV1.CartonCode)
            Pieces = "1"
            Dimensions = String.Empty

            BillType = String.Empty
            BillNum = FromCustId


            '' UniqueRecordId - This value is derived from TrackingNum and TranDate
            Dim sb As New StringBuilder
            sb.Append(TranDate.ToString("MMddyyyyHHmmss", CultureInfo.InvariantCulture))
            sb.Append(TrackingNumber)
            _sUniqueRecordId = sb.ToString()

            Void = "F"
            ReferenceNumber = String.Empty
            PONumber = String.Empty
            ThirdPartyBillNum = String.Empty
            Modifiers = String.Empty
            sDeclaredValue = 0D

            EmptyField = String.Empty

        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges

            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function

    Private Function DecomposeMedExRecordString(ByVal p_strSubstrings() As String) As Boolean

        'Dim sFromAddRowId As String
        'Dim sToAddRowId As String
        'Dim sWeight As String
        'Dim sDeclaredValue As String

        Try
            'Toss Header
            If String.Compare(p_strSubstrings(FieldsV2.ToLocName), "PATIENT") = 0 Then
                _bHasError = False
                Return False
            End If

            ToLocName = p_strSubstrings(FieldsV2.ToLocName)
            Dim sWords As String() = p_strSubstrings(FieldsV2.ToLocStreet).Split(New Char() {">"c})
            ToLocStreet = sWords(0)
            ToLocCity = p_strSubstrings(FieldsV2.ToLocCity)
            ToLocState = p_strSubstrings(FieldsV2.ToLocState)
            ToLocZip = p_strSubstrings(FieldsV2.ToLocZip)
            If IsNumeric(sWords(1)) Then ServiceLevel = "OVRNGT-RES" Else ServiceLevel = "OVRNGT-BUS"

            'OrderId = p_strSubstrings(FieldsV2.OrderId)
            'ServiceLevel = p_strSubstrings(FieldsV2.ServiceLevel)
            'TrackingNumber = "MDX" & _iLineNum.ToString() & ToLocName.Substring(0, 2) & ToLocStreet.Substring(0, 2) & ToLocCity.Substring(0, 2) & ToLocZip.Substring(3, 2) & TranDate.Day.ToString()

            ' The Formatted File will append -DDLN (dash + day of month + line number) to the tracking number, but that will be stripped away
            ' when imported into the "ThirdPartyBarcode" field.  It will be included when construction the UniqueRowId.
            'TrackingNumber = p_strSubstrings(FieldsV2.TrackingNumber)
            Dim sTrNumParts As String() = p_strSubstrings(FieldsV2.TrackingNumber).Split(New Char() {"-"c})
            TrackingNumber = sTrNumParts(0)

            'The TranDate is provided via the File's Name. Value should be set explicitly outside of this method.

            '' UniqueRecordId - This value is derived from TrackingNum and TranDate
            Dim sb As New StringBuilder
            sb.Append(TranDate.ToString("MMddyyyyHHmmss", CultureInfo.InvariantCulture))
            'sb.Append(TrackingNumber)
            sb.Append(p_strSubstrings(FieldsV2.TrackingNumber))
            _sUniqueRecordId = sb.ToString()


        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges

            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function

    Private Function DecomposeMedExSdRecordString(ByVal p_strSubstrings() As String) As Boolean

        'Dim sFromAddRowId As String
        'Dim sToAddRowId As String
        'Dim sWeight As String
        'Dim sDeclaredValue As String

        Try
            'Toss Header
            If String.Compare(p_strSubstrings(FieldsV2.ToLocName), "PATIENT") = 0 Then
                _bHasError = False
                Return False
            End If

            ToLocName = p_strSubstrings(FieldsV2.ToLocName)
            Dim sWords As String() = p_strSubstrings(FieldsV2.ToLocStreet).Split(New Char() {">"c})
            ToLocStreet = sWords(0)
            ToLocCity = p_strSubstrings(FieldsV2.ToLocCity)
            ToLocState = p_strSubstrings(FieldsV2.ToLocState)
            ToLocZip = p_strSubstrings(FieldsV2.ToLocZip)
            If IsNumeric(sWords(1)) Then ServiceLevel = "SAMEDAY-RES" Else ServiceLevel = "SAMEDAY-BUS"

            'OrderId = p_strSubstrings(FieldsV2.OrderId)
            'ServiceLevel = p_strSubstrings(FieldsV2.ServiceLevel)
            'TrackingNumber = "MDX" & _iLineNum.ToString() & ToLocName.Substring(0, 2) & ToLocStreet.Substring(0, 2) & ToLocCity.Substring(0, 2) & ToLocZip.Substring(3, 2) & TranDate.Day.ToString()

            ' The Formatted File will append -DDLN (dash + day of month + line number) to the tracking number, but that will be stripped away
            ' when imported into the "ThirdPartyBarcode" field.  It will be included when construction the UniqueRowId.
            'TrackingNumber = p_strSubstrings(FieldsV2.TrackingNumber)
            Dim sTrNumParts As String() = p_strSubstrings(FieldsV2.TrackingNumber).Split(New Char() {"-"c})
            TrackingNumber = sTrNumParts(0)

            'The TranDate is provided via the File's Name. Value should be set explicitly outside of this method.

            '' UniqueRecordId - This value is derived from TrackingNum and TranDate
            Dim sb As New StringBuilder
            sb.Append(TranDate.ToString("MMddyyyyHHmmss", CultureInfo.InvariantCulture))
            'sb.Append(TrackingNumber)
            sb.Append(p_strSubstrings(FieldsV2.TrackingNumber))
            _sUniqueRecordId = sb.ToString()


        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges

            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function

    Private Function DecomposeBluePackageRecordString(ByVal p_strSubstrings() As String) As Boolean

        'Dim sFromAddRowId As String
        'Dim sToAddRowId As String
        Dim sWeight As String
        'Dim sDeclaredValue As String

        Try

            'Toss Header
            If String.Compare(p_strSubstrings(FieldsV3.ReferenceNumber), "Sortcode") = 0 Then
                _bHasError = False
                Return False
            End If

            'TranDate is set by the filename

            ReferenceNumber = p_strSubstrings(FieldsV3.ReferenceNumber)
            ToLocId = p_strSubstrings(FieldsV3.ToLocZip)
            ToLocName = p_strSubstrings(FieldsV3.ToLocName)
            ToLocStreet = p_strSubstrings(FieldsV3.ToLocStreet)
            ToLocCity = p_strSubstrings(FieldsV3.ToLocCity)
            ToLocZip = ToLocId
            sWeight = p_strSubstrings(FieldsV3.Weight)
            Pieces = p_strSubstrings(FieldsV3.Pieces)

            TrackingNumber = p_strSubstrings(FieldsV3.ToLocZip) + TranDate.Day.ToString + LineNumber.ToString()

            '' UniqueRecordId - This value is derived from TrackingNum and TranDate
            Dim sb As New StringBuilder
            sb.Append(TranDate.ToString("MMddyyyyHHmmss", CultureInfo.InvariantCulture))
            sb.Append(TrackingNumber)
            _sUniqueRecordId = sb.ToString()

            If IsNumeric(sWeight) Then Weight = CDec(sWeight) Else Weight = 0D

        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges

            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function

    Private Function DecomposeAutoZoneRecordString(ByVal p_strSubstrings() As String) As Boolean

        'Dim sFromAddRowId As String
        'Dim sToAddRowId As String
        'Dim sWeight As String
        'Dim sDeclaredValue As String

        Try


            FromLocId = p_strSubstrings(FieldsV4.FromLocId)
            ToLocId = p_strSubstrings(FieldsV4.ToLocId)
            TranDate = CDate(p_strSubstrings(FieldsV4.TranDate))
            Pieces = p_strSubstrings(FieldsV4.Pieces)
            CartonCode = p_strSubstrings(FieldsV4.CartonCode)
            ServiceLevel = p_strSubstrings(FieldsV4.ServiceLevel)

            TrackingNumber = ToLocId + TranDate.Ticks.ToString
            If TrackingNumber.Length > 15 Then TrackingNumber = TrackingNumber.Substring(1, 15)

            '' UniqueRecordId - This value is derived from TrackingNum and TranDate
            Dim sb As New StringBuilder
            sb.Append(TranDate.ToString("MMddyyyyHHmmss", CultureInfo.InvariantCulture))
            sb.Append(TrackingNumber)
            _sUniqueRecordId = sb.ToString()


        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges

            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function

    Private Function DecomposeDSCRecordString(ByVal p_strSubstrings() As String) As Boolean

        'Dim sFromAddRowId As String
        'Dim sToAddRowId As String
        'Dim sWeight As String
        'Dim sDeclaredValue As String

        Try
            'Toss Header
            If String.Compare(p_strSubstrings(FieldsV5.TrackingNumber), "Tracking Number") = 0 Then
                _bHasError = False
                Return False
            End If

            'Differentiate between Irvine (56) and Burbank (228) and San Gabriel (347) and Parkview (660) Locations
            If p_strSubstrings(FieldsV5.ThirdPartyBillNum).CompareTo("56") = 0 Then
                FromCustName = "DSC DELIVERY"
                FromCustId = "13025A"
                ToCustName = FromCustName
                ToCustId = FromCustId
                FromAddRowId = 53162
                FromLocId = "001"
                FromLocName = "OSO HOMECARE"
                FromLocStreet = "17175 GILLETTE AVE"
                FromLocCity = "IRVINE"
                FromLocState = "CA"
                FromLocZip = "92614"
                BillNum = FromCustId
                Modifiers = p_strSubstrings(FieldsV5.Modifiers)
            ElseIf p_strSubstrings(FieldsV5.ThirdPartyBillNum).CompareTo("228") = 0 Then
                FromCustName = "DSC DELIVERY"
                FromCustId = "13025B"
                ToCustName = FromCustName
                ToCustId = FromCustId
                FromAddRowId = 56674
                FromLocId = "001"
                FromLocName = "OSO HOMECARE"
                FromLocStreet = "1214 BURBANK BLVD"
                FromLocCity = "BURBANK"
                FromLocState = "CA"
                FromLocZip = "91506"
                BillNum = FromCustId
                Modifiers = p_strSubstrings(FieldsV5.Modifiers)
            ElseIf p_strSubstrings(FieldsV5.ThirdPartyBillNum).CompareTo("347") = 0 Then
                FromCustId = "13025C"
                ToCustId = FromCustId
                FromAddRowId = 179816
                FromLocStreet = "546 W LAS TUNAS DR"
                FromLocAddress2 = "STE 102"
                FromLocCity = "SAN GABRIEL"
                FromLocState = "CA"
                FromLocZip = "91776"
                BillNum = FromCustId
                Modifiers = p_strSubstrings(FieldsV5.Modifiers)
            ElseIf p_strSubstrings(FieldsV5.ThirdPartyBillNum).CompareTo("660") = 0 Then
                FromCustName = "DCS DELIVERY"
                FromCustId = "14017"
                ToCustName = FromCustName
                ToCustId = FromCustId
                FromAddRowId = 201205
                FromLocId = "002"
                FromLocName = "PARK VIEW MEDICAL CTR PHARMACY"
                FromLocStreet = "3975 JACKSON ST"
                FromLocAddress2 = "STE 109"
                FromLocCity = "RIVERSIDE"
                FromLocState = "CA"
                FromLocZip = "92503"
                BillNum = FromCustId
                Modifiers = p_strSubstrings(FieldsV5.Modifiers)
            ElseIf p_strSubstrings(FieldsV5.ThirdPartyBillNum).CompareTo("662") = 0 Then
                FromCustName = "DCS DELIVERY"
                FromCustId = "14017"
                ToCustName = FromCustName
                ToCustId = FromCustId
                FromAddRowId = 207600
                FromLocId = "003"
                FromLocName = "PARKVIEW MEDICAL PLAZA PHARMACY"
                FromLocStreet = "3975 JACKSON ST"
                FromLocAddress2 = "STE 111"
                FromLocCity = "RIVERSIDE"
                FromLocState = "CA"
                FromLocZip = "92503"
                BillNum = FromCustId
                Modifiers = p_strSubstrings(FieldsV5.Modifiers)
            ElseIf p_strSubstrings(FieldsV5.ThirdPartyBillNum).CompareTo("661") = 0 Then
                FromCustName = "DCS DELIVERY"
                FromCustId = "14018"
                ToCustName = FromCustName
                ToCustId = FromCustId
                FromAddRowId = 201204
                FromLocId = "002"
                FromLocName = "PARK VISTA PHARMACY"
                FromLocStreet = "3838 SHERMAN DR"
                FromLocAddress2 = "STE 1"
                FromLocCity = "RIVERSIDE"
                FromLocState = "CA"
                FromLocZip = "92503"
                BillNum = FromCustId
                Modifiers = p_strSubstrings(FieldsV5.Modifiers)
            ElseIf p_strSubstrings(FieldsV5.ThirdPartyBillNum).CompareTo("902") = 0 Then
                FromCustName = "DCS DELIVERY"
                FromCustId = "14020"
                ToCustName = FromCustName
                ToCustId = FromCustId
                FromAddRowId = 207363
                FromLocId = "002"
                FromLocName = "BIO LOGIC INFUSION PHARMACY"
                FromLocStreet = "8851 WATSON ST"
                FromLocCity = "CYPRESS"
                FromLocState = "CA"
                FromLocZip = "90630"
                BillNum = FromCustId
                Modifiers = String.Empty
                CartonCode = RTrim(p_strSubstrings(FieldsV5.Modifiers))
            Else
                Return False 'Unknown Customer
            End If

            'Remove the " (Oso)" from the end of the string
            Dim sLastSix As String = p_strSubstrings(FieldsV5.ServiceLevel).Substring(Len(p_strSubstrings(FieldsV5.ServiceLevel)) - 6, 6)
            If sLastSix.CompareTo(" (Oso)") = 0 Then
                ServiceLevel = p_strSubstrings(FieldsV5.ServiceLevel).Substring(0, Len(p_strSubstrings(FieldsV5.ServiceLevel)) - 6)
            Else
                ServiceLevel = p_strSubstrings(FieldsV5.ServiceLevel)
            End If
            'ServiceLevel = p_strSubstrings(FieldsV5.ServiceLevel)


            TrackingNumber = p_strSubstrings(FieldsV5.TrackingNumber)
            ThirdPartyBillNum = p_strSubstrings(FieldsV5.ThirdPartyBillNum)
            ToLocName = p_strSubstrings(FieldsV5.ToLocName)
            ToLocCity = p_strSubstrings(FieldsV5.ToLocCity)
            ToLocStreet = p_strSubstrings(FieldsV5.ToLocStreet)
            ToLocAddress2 = p_strSubstrings(FieldsV5.ToLocAddress2)
            ToLocState = p_strSubstrings(FieldsV5.ToLocState)
            ToLocZip = p_strSubstrings(FieldsV5.ToLocZip)
            Pieces = p_strSubstrings(FieldsV5.Pieces)
            If IsNumeric(p_strSubstrings(FieldsV5.Weight)) Then
                Weight = CDec(p_strSubstrings(FieldsV5.Weight))
            Else
                Weight = 0D
            End If

            'The TranDate is provided via the File's Name. Value should be set explicitly outside of this method.

            '' UniqueRecordId - This value is derived from TrackingNum and TranDate
            Dim sb As New StringBuilder
            sb.Append(TranDate.ToString("MMddyyyyHHmmss", CultureInfo.InvariantCulture))
            'sb.Append(TrackingNumber)
            sb.Append(p_strSubstrings(FieldsV5.TrackingNumber))
            _sUniqueRecordId = sb.ToString()


        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges

            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function

    Private Function DecomposePiaRecordString(ByVal p_strSubstrings() As String) As Boolean

        'Dim sFromAddRowId As String
        'Dim sToAddRowId As String
        'Dim sWeight As String
        'Dim sDeclaredValue As String

        Dim bContinue As Boolean = False

        Try

            'Validate Contents Beyond Number of Fields (which have already been matched)
            If (String.Compare(p_strSubstrings(FieldsV6.FromCustName), "VSPW") = 0) Or (String.Compare(p_strSubstrings(FieldsV6.FromCustName), "CSP SOLANO") = 0) Then
                bContinue = True
            Else
                bContinue = False
                _bHasError = False
                Return False
            End If

            'Differentiate between Chowchilla & Vacaville Locations
            If p_strSubstrings(FieldsV6.FromLocCity).CompareTo("CHOWCHILLA") = 0 Then
                FromCustId = "14007"
                BillNum = FromCustId
                FromCustName = "VSPW"
                ToCustId = FromCustId
                ToCustName = FromCustName
                FromAddRowId = 168796
                FromLocName = "VSPW"
                FromLocStreet = "23370 ROAD 22"
                FromLocZip = "93610"
            ElseIf p_strSubstrings(FieldsV6.FromLocCity).CompareTo("VACAVILLE") = 0 Then
                FromCustId = "14009"
                BillNum = FromCustId
                FromCustName = "CSP SOLANO"
                ToCustId = FromCustId
                ToCustName = FromCustName
                FromAddRowId = 168797
                FromLocName = "CSP SOLANO"
                FromLocStreet = "2100 PEABODY ROAD"
                FromLocZip = "95687"
            Else
                _bHasError = False
                Return False
            End If


            FromLocCity = p_strSubstrings(FieldsV6.FromLocCity)
            FromLocState = p_strSubstrings(FieldsV6.FromLocState)
            FromLocPhone = p_strSubstrings(FieldsV6.FromLocPhone)
            ToLocName = p_strSubstrings(FieldsV6.ToLocName)
            ToLocStreet = p_strSubstrings(FieldsV6.ToLocStreet)
            ToLocAddress2 = p_strSubstrings(FieldsV6.ToLocAddress2)
            ToLocCity = p_strSubstrings(FieldsV6.ToLocCity)
            ToLocState = p_strSubstrings(FieldsV6.ToLocState)
            ToLocZip = p_strSubstrings(FieldsV6.ToLocZip)
            TrackingNumber = p_strSubstrings(FieldsV6.TrackingNumber)
            ToLocId = p_strSubstrings(FieldsV6.ToLocId)
            If IsNumeric(ToLocId) Then
                ToLocId = CInt(ToLocId).ToString
            End If
            TranDate = CDate(p_strSubstrings(FieldsV6.TranDate))

            '' UniqueRecordId - This value is derived from TrackingNum and TranDate
            Dim sb As New StringBuilder
            sb.Append(TranDate.ToString("MMddyyyyHHmmss", CultureInfo.InvariantCulture))
            sb.Append(p_strSubstrings(FieldsV6.TrackingNumber))
            _sUniqueRecordId = sb.ToString()

        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges

            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function

    Private Function DecomposeProCourierRecordString(ByVal p_strSubstrings() As String) As Boolean

        'Dim sFromAddRowId As String
        'Dim sToAddRowId As String
        'Dim sWeight As String
        'Dim sDeclaredValue As String

        Try

            ''Toss Header - This section is not necessary because first two rows will not be included in delimited export as per SOP
            'If String.Compare(p_strSubstrings(FieldsV2.ToLocName), "PATIENT") = 0 Then
            '    _bHasError = False
            '    Return False
            'End If

            TrackingNumber = p_strSubstrings(FieldsV7.TrackingNumber)

            FromLocName = p_strSubstrings(FieldsV7.FromLocName)
            FromLocStreet = p_strSubstrings(FieldsV7.FromLocStreet)
            FromLocCity = p_strSubstrings(FieldsV7.FromLocCity)
            FromLocState = p_strSubstrings(FieldsV7.FromLocState)
            FromLocZip = p_strSubstrings(FieldsV7.FromLocZip)

            ToLocName = p_strSubstrings(FieldsV7.ToLocName)
            ToLocStreet = p_strSubstrings(FieldsV7.ToLocStreet)
            ToLocCity = p_strSubstrings(FieldsV7.ToLocCity)
            ToLocState = p_strSubstrings(FieldsV7.ToLocState)
            ToLocZip = p_strSubstrings(FieldsV7.ToLocZip)

            ServiceLevel = p_strSubstrings(FieldsV7.ServiceLevel)
            Pieces = p_strSubstrings(FieldsV7.Pieces)

            'The TranDate is provided via the File's Name. Value should be set explicitly outside of this method.

            '' UniqueRecordId - This value is derived from TrackingNum and TranDate
            Dim sb As New StringBuilder
            sb.Append(TranDate.ToString("MMddyyyyHHmmss", CultureInfo.InvariantCulture))
            sb.Append(TrackingNumber)
            _sUniqueRecordId = sb.ToString()

        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges
            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function

    Private Function DecomposeDIO2CSRecordString(ByVal p_strSubstrings() As String) As Boolean

        'Dim sFromAddRowId As String
        'Dim sToAddRowId As String
        'Dim sWeight As String
        'Dim sDeclaredValue As String

        Try

            'The Header section is not necessary because first two rows will not be included in delimited export as per SOP

            'These values come directly from the import file
            TrackingNumber = p_strSubstrings(FieldsV8.TrackingNumber) 'aka OrderNumber
            ChargedBySub = p_strSubstrings(FieldsV8.ChargedBySub) ' aka AmountCharged
            Weight = p_strSubstrings(FieldsV8.Weight)
            FromLocId = p_strSubstrings(FieldsV8.FromLocId) 'CustomerCode
            FromLocName = p_strSubstrings(FieldsV8.FromLocName) 'aka OriginName
            FromLocStreet = p_strSubstrings(FieldsV8.FromLocStreet) 'aka OriginAddress
            FromLocCity = p_strSubstrings(FieldsV8.FromLocCity) 'aka OriginCity
            FromLocZip = p_strSubstrings(FieldsV8.FromLocZip) 'aka OriginZip
            ToLocName = p_strSubstrings(FieldsV8.ToLocName) 'aka DestName
            ToLocStreet = p_strSubstrings(FieldsV8.ToLocStreet) 'aka DestAddress
            ToLocCity = p_strSubstrings(FieldsV8.ToLocCity) ' aka DestCity
            ToLocState = p_strSubstrings(FieldsV8.ToLocState) 'aka DestState
            ToLocZip = p_strSubstrings(FieldsV8.ToLocZip) 'aka DestZip
            Pieces = p_strSubstrings(FieldsV8.Pieces)
            OrderId = p_strSubstrings(FieldsV8.OrderId) 'aka OrderAlias
            TranDate = p_strSubstrings(FieldsV8.TranDate) 'aka OrderDate
            CartonCode = p_strSubstrings(FieldsV8.CartonCode) 'aka ParcelType

            'The TranDate is provided via the File's Name. Value should be set explicitly outside of this method.

            '' UniqueRecordId - This value is derived from TrackingNum and TranDate
            Dim sb As New StringBuilder
            sb.Append(TranDate.ToString("MMddyyyyHHmmss", CultureInfo.InvariantCulture))
            sb.Append(TrackingNumber)
            _sUniqueRecordId = sb.ToString()

        Catch ex As Exception
            ' This Catch Block Should Catch any Type Mismatches whereas the Assign functions validate value ranges
            SetError(ex.Message & " [ScanRecord::DecomposeRecordString]")
            Reset()
            Return False

        End Try

        Return True

    End Function


    Private Function DecomposeRecordString(ByVal p_strSubStrings() As String) As Boolean

        If String.Compare(_sFileFormat, "V0") = 0 Then Return DecomposeIngramMicroRecordString(p_strSubStrings)
        If String.Compare(_sFileFormat, "V1") = 0 Then Return DecomposePartnerBooksRecordString(p_strSubStrings)
        If String.Compare(_sFileFormat, "V2") = 0 Then Return DecomposeMedExRecordString(p_strSubStrings)
        If String.Compare(_sFileFormat, "V2s") = 0 Then Return DecomposeMedExSdRecordString(p_strSubStrings)
        If String.Compare(_sFileFormat, "V3") = 0 Then Return DecomposeBluePackageRecordString(p_strSubStrings)
        If String.Compare(_sFileFormat, "V4") = 0 Then Return DecomposeAutoZoneRecordString(p_strSubStrings)
        If String.Compare(_sFileFormat, "V5") = 0 Then Return DecomposeDSCRecordString(p_strSubStrings)
        If String.Compare(_sFileFormat, "V6") = 0 Then Return DecomposePiaRecordString(p_strSubStrings)
        If String.Compare(_sFileFormat, "V7") = 0 Then Return DecomposeProCourierRecordString(p_strSubStrings)
        If String.Compare(_sFileFormat, "V8") = 0 Then Return DecomposeDIO2CSRecordString(p_strSubStrings)

    End Function


    Private Class SpecialHandleColumn

        Dim _fDeclaredValue As Decimal
        Dim _oCarton As Carton
        Dim _sModifiers As String

        Public Property DeclaredValue() As Decimal
            Get
                Return _fDeclaredValue
            End Get
            Set(ByVal Value As Decimal)
                _fDeclaredValue = Value
            End Set
        End Property

        Public Property Dimensions() As String
            Get
                If Not _oCarton Is Nothing Then Return _oCarton.DimensionString Else Return "0.00*0.00*0.00"
            End Get
            Set(ByVal Value As String)
                _oCarton = Nothing
                _oCarton = New Carton(Value)
            End Set
        End Property

        Public Property Modifiers() As String
            Get
                Return _sModifiers
            End Get
            Set(ByVal Value As String)
                _sModifiers = Value
            End Set
        End Property

        Public ReadOnly Property SpecialHandleString() As String
            Get

                Dim s As StringBuilder = Nothing

                If Not _oCarton Is Nothing Then s.Append("")
                s.Append("&")
                If _sModifiers <> String.Empty Then s.Append(_sModifiers)
                s.Append("&")
                s.Append(_fDeclaredValue.ToString("f2"))

                Return s.ToString()

            End Get
        End Property

        Sub New()
            _fDeclaredValue = 0.0
            _oCarton = Nothing
            _sModifiers = String.Empty
        End Sub

    End Class

End Class
