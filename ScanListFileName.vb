Imports System.Text

'<ClassHeader scope="Public" name="ScanFileName">
'   <description>
'       This class will help manage the actual file name of a ScanList file.  It will know how to interpret each part of the file name and 
'       it will know how to modify the name based on certain milestones in a typical files life cycle.
'       
'       There are two ways to use this class.  You can either create a blank object {by calling New with no paramters}, initialize its properties 
'       and then retrieve the properly formatted file name (by reading the FileName property).  Or you can initialize a new object with an exising
'       filename (by calling New with the filename parameter) and then inspect the various parts of the file.  If the filename passed to the
'       object is not property formatted and exception will be raised.
'
'       Regardless of which method you used to create an instance of this class, you can always change the "filename" it is dealing with by
'       calling the SetFileName method.
'   </description>
'</ClassHeader>
Public Class ScanListFileName

    '<variable scope="Private" name="m_bUploaded" type="boolean" default="false">
    '   <description>
    '       Flag to indicate whether or not the file has been uploaded to the server.
    '   </description>
    '</variable>
    Private m_bUploaded As Boolean = False

    '<property scope="Public" name="Uploaded" type="Boolean" readonly="false">
    '   <description>
    '       Allows consumers to read or set the m_bUploaded variable
    '   </description>
    '</property>
    Public Property Uploaded() As Boolean
        Get
            Return m_bUploaded
        End Get
        Set(ByVal Value As Boolean)
            m_bUploaded = True
        End Set
    End Property

    '<variable scope="Private" name="m_bImported" type="boolean" default="false">
    '   <description>
    '       Flag to indicate whether or not the file has been successfully imported
    '   </description>
    '</variable>
    Private m_bImported As Boolean = False

    '<property scope="Public" name="Imported" type="Boolean" readonly="false">
    '   <description>
    '       Allows consumers to get or set the value of m_bImported
    '   </description>
    '</property>
    Public Property Imported() As Boolean
        Get
            Return m_bImported
        End Get
        Set(ByVal Value As Boolean)
            m_bImported = Value
        End Set
    End Property

    '<variable scope="Private" name="Problems" type="Boolean" default="false">
    '   <description>
    '       Flag to indicate that some of the records in the file could not be imported
    '   </description>
    '</variable>
    Private m_bProblems As Boolean = False

    '<property scope="Public" name="Problems" type="" readonly="false">
    '   <description>
    '       Allows consmers to get or set the m_bProblems variable.
    '   </description>
    '</property>
    Public Property Problems() As Boolean
        Get
            Return m_bProblems
        End Get
        Set(ByVal Value As Boolean)
            m_bProblems = Value
        End Set
    End Property

    '<variable scope="private" name="BranchId" type="String" default="00">
    '   <description>
    '       Represents the part of the file name that identifies the branch that created it
    '   </description>
    '</variable>
    Private m_sBranchId As String = "00"

    '<property scope="public" name="BranchId" type="Integer" readonly="false">
    '   <description>
    '       Allows consmers to get or set the m_sBranchId variable.  It will convert from integer to string and vice-versa, making sure
    '       there is always a leading "0" in the string version.
    '   </description>
    '</property>
    Public Property BranchId() As Integer

        Get
            Return CInt(m_sBranchId)
        End Get

        Set(ByVal Value As Integer)

            If Value < 0 Or Value > 99 Then
                m_sBranchId = "00"
                Throw New Exception("Invalid BranchId Specified. [ScanListFileName::BranchId]")
            End If

            m_sBranchId = String.Empty

            If Value < 10 Then m_sBranchId = "0"

            m_sBranchId &= CStr(Value)

        End Set

    End Property

    '<variable scope="private" name="m_sWksId" type="string" default="W0">
    '   <description>
    '       Represents the part of the file name that identifies the scanning workstation that created it.
    '   </description>
    '</variable>
    Private m_sWksId As String = "W0"

    '<property scope="Public" name="StationId" type="string" readonly="false">
    '   <description>
    '       Allows consmers to get or set the m_sWksId variable.  Only allows two character descriptions.
    '   </description>
    '</property>
    Public Property StationId() As String

        Get
            Return m_sWksId
        End Get

        Set(ByVal Value As String)

            If Value.Length <> 2 Then
                m_sWksId = "W0"
                Throw New Exception("Invalid Workstation Id Specified. [ScanListFileName::StationId]")
            End If

            m_sWksId = Value

        End Set

    End Property

    '<variable scope="Private" name="m_sMonth" type="string" default="01">
    '   <description>
    '       Represents the part of the file name that identifies the Month it was created.
    '   </description>
    '</variable>
    Private m_sMonth As String = "01"

    '<property scope="Public" name="Month" type="integer" readonly="false">
    '   <description>
    '       Allows consumers to get or set the m_sMonth variable.  This property will properly convert the month into properly formatted string.
    '   </description>
    '</property>
    Public Property BillingMonth() As Integer

        Get
            Return CInt(m_sMonth)
        End Get

        Set(ByVal Value As Integer)

            m_sMonth = String.Empty

            If Value < 1 Or Value > 12 Then
                m_sMonth = "01"
                Throw New Exception(Value & " is not a valid month. [ScanListFileName::Month]")
            End If

            If Value < 10 Then m_sMonth = "0"
            m_sMonth &= CStr(Value)

        End Set

    End Property

    '<variable scope="private" name="m_sDay" type="string" default="01">
    '   <description>
    '       Represents the part of the file name that identifies the day it was created.
    '   </description>
    '</variable>
    Private m_sDay As String = "01"

    '<property scope="Public" name="Day" type="integer" readonly="false">
    '   <description>
    '       Allows consumer to get and/or set the value of m_sDay.  Performs basic validation.
    '   </description>
    '</property>
    Public Property BillingDay() As Integer

        Get
            Return CInt(m_sDay)
        End Get

        Set(ByVal Value As Integer)

            m_sDay = String.Empty

            If Value < 1 Or Value > 31 Then
                m_sDay = "01"
                Throw New Exception(Value & " is not a valid day. [ScanListFileName::Day]")
            End If

            If Value < 10 Then m_sDay = "0"
            m_sDay &= CStr(Value)

        End Set

    End Property

    '<variable scope="private" name="m_sYear" type="string" default="2008">
    '   <description>
    '       Represents the part of the file name that identifies the year it was created
    '   </description>
    '</variable>
    Private m_sYear As String = "2008"

    '<property scope="public" name="Year" type="string" readonly="false">
    '   <description>
    '       Allows consumer to get and/or set the value of m_sYear.  Performs basic validation and formatting.
    '   </description>
    '</property>
    Public Property BillingYear() As Integer

        Get
            Return CInt(m_sYear)
        End Get

        Set(ByVal Value As Integer)

            m_sYear = String.Empty

            ' Bogus Year
            If Value < 1 Or Value > 9999 Then
                m_sYear = "2008"
                Throw New Exception(Value & " is not a valid date. [ScanListFileName::Year]")
            End If

            If Value < 100 Then ' year was specified as a two digit year
                m_sYear = "20"
                If Value < 10 Then m_sYear &= "0"
                m_sYear &= CStr(Value)
            Else
                If Value < 1000 Then m_sYear = "0"
                m_sYear &= CStr(Value) ' Assumes year was specifed with 4 digits
            End If

        End Set
    End Property

    '<property scope="public" name="DateString" type="String" readonly="true">
    '   <description>
    '       Returns the date in a properly formatted string.
    '   </description>
    '</property>
    Public ReadOnly Property BillingDateString() As String
        Get
            Return m_sMonth & m_sDay & m_sYear
        End Get
    End Property

    '<property scope="public" name="BillingDate" type="date" readonly="false" writeonly="true">
    '   <description>
    '       Sets the date variables to match the date Value specified.
    '   </description>
    '</property>
    Public WriteOnly Property BillingDate() As Date

        Set(ByVal Value As Date)

            BillingMonth = Value.Month
            BillingDay = Value.Day
            BillingYear = Value.Year

        End Set

    End Property

    '<variable scope="private" name="m_sHour" type="string" default="00">
    '   <description>
    '       Represents the part of the file that identifies the hour of the day (using 24 hour clock) it was created.
    '   </description>
    '</variable>
    Private m_sHour As String = "00"

    '<property scope="public" name="BillingHour" type="integer" readonly="false">
    '   <description>
    '       Allows consumer to get and/or set the hour of the day in which the billing day starts.  Performs basic validation.
    '   </description>
    '</property>
    Public Property BillingHour() As Integer

        Get
            Return CInt(m_sHour)
        End Get

        Set(ByVal Value As Integer)

            m_sHour = String.Empty

            If Value < 0 Or Value > 23 Then
                m_sHour = "00"
                Throw New Exception(Value & " is not a valid hour of the day. [ScanListFileName::BillingHour]")
            End If

            If Value < 10 Then m_sHour = "0"
            m_sHour &= CStr(Value)

        End Set

    End Property

    '<variable scope="private" name="m_sMinute" type="string" default="00">
    '   <description>
    '       Represents the part of the file name that identifies the minutes of the hour it was created
    '   </description>
    '</variable>
    Private m_sMinute As String = "00"

    '<property scope="public" name="BillingMinute" type="integer" readonly="false">
    '   <description>
    '       Allows consumer to get and/or set the minute of the hour in which the Billing day begins.  Performs basic validation.
    '   </description>
    '</property>
    Public Property BillingMinute() As Integer

        Get
            Return CInt(m_sMinute)
        End Get

        Set(ByVal Value As Integer)

            m_sMinute = String.Empty

            If Value < 0 Or Value > 59 Then
                m_sMinute = "00"
                Throw New Exception(Value & " is not a valid value for minutes. [ScanListFileName::BillingMinute]")
            End If

            If Value < 10 Then m_sMinute = "0"
            m_sMinute &= CStr(Value)

        End Set

    End Property

    '<property scope="public" name="BillingTimeString" type="string" readonly="true">
    '   <description>
    '       Allows the consumer to set get the properly formatted time string.
    '   </description>
    '</property>
    Public ReadOnly Property BillingTimeString() As String
        Get
            Return m_sHour & m_sMinute
        End Get
    End Property

    '<property scope="public" name="BillingTime" type="date" writeonly="true">
    '   <description>
    '       Allows consumer to set the Start of Billing Day hour and minute by specifying a datetime value.
    '   </description>
    '</property>
    Public WriteOnly Property BillingTime() As Date

        Set(ByVal Value As Date)

            BillingHour = Value.Hour
            BillingMinute = Value.Minute

        End Set

    End Property

    '<property scope="public" name="FileName" type="string" readonly="true">
    '   <description>
    '       Returns the properly formatted Scan List File name, based on the value of its member variables.
    '   </description>
    '</property>
    Public ReadOnly Property FileName() As String

        Get

            Dim sb As StringBuilder = Nothing

            sb.Append(m_sBranchId)
            sb.Append(m_sWksId)
            sb.Append(m_sMonth)
            sb.Append(m_sDay)
            sb.Append(m_sYear)
            sb.Append(m_sHour)
            sb.Append(m_sMinute)

            If m_bUploaded = False And m_bImported = False Then
                sb.Append(".SCAN")
            End If

            If m_bUploaded = True And m_bImported = False Then
                sb.Append(".UPLD")
            End If

            If m_bImported = True Then
                If m_bProblems = True Then
                    sb.Append(".PERR")
                Else
                    sb.Append(".PROC")
                End If
            End If

            Return sb.ToString()

        End Get

    End Property



End Class
'<ClassFooter />
