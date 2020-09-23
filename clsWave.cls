VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Originally written September 06, 2006 by Nick Williams (stolenflipflops@yahoo.com)
'
'You can distribute this freely, I don't expect anything aside from perhaps
'saving some people time and effort dealing with similar projects that I am now.
'
'Change Log:
'
'Edited September 07, 2006 by Nick Williams (stolenflipflops@yahoo.com)
'Today I altered this object's properties to directly manipulate the now module-scoped
'WaveHeader udt - this eliminates redundant module-scoped holders for the property data.
'The WriteWave routine now accepts (optionally) all pertinent parameters needed to
'generate the rest of the header and write the passed byte array.  I did this because
'it occured to me that some developers may want to use this class to only write to a
'wave, or use an already instantiated version of this object to write a wave file
'completely independent of the file that has been opened by it (if any).  If the
'optional parameters are not given, the WriteWave routine will default to using the
'WaveHeader udt's data.
'The OpenWave routine no longer populates the byte array immediately when opened.  The
'byte array is only populated the *first* time GetData() is executed, subsequent execut
'ions return the byte array directly rather than converting the data string into a byte
'array each time.  This was done to extend the class's use to that of simply retrieving
'WAVE file info without unnecessarily processing the data any further than reading it.
'
Option Explicit
'Wave Header
Private Header As WaveHeader
'Wave Data
Private m_strData As String 'String version, retrieved immediately when file is opened
Private m_bytData() As Byte 'Only populated from the first call of GetData()
'Memory manipulation
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal olestr As Long, ByVal BLen As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal lngBytes As Long)
'Wave properties
Private m_lngWaveLength As Integer
Private m_lngWaveSamples As Long
Private m_lngWaveDataLength As Long

'Used for error reporting
Private Const WM_APP As Long = &H8000&
'Numerical values of text in the wave file.  FYI - numerica comparisons are ~3 times
'faster than textual comparisons.
Private Const RIFF As Long = 1179011410
Private Const WAVE As Long = 1163280727
Private Const FMT As Long = 544501094
Private Const DATA As Long = 1635017060
'The wave file's header structure
Private Type WaveHeader
    '44 bytes total
    RiffChunkID As Long
    RiffChunkSize As Long
    RiffChunkFormat As Long
    RiffSubChunkID As Long
    RiffSubChunkSize As Long
    WaveAudioFormat As Integer
    WaveNumChannels As Integer
    WaveSampleRate As Long
    WaveByteRate As Long
    WaveBlockAlign As Integer
    WaveBitsPerSample As Integer
    RiffSubChunk2ID As Long
    WaveDataLength As Long
    '
End Type

'Credits jory@joryanick.com - http://www.xbeat.net/vbspeed/c_AllocString.htm
Private Function AllocString(ByVal lngSize As Long) As String
    '
    RtlMoveMemory ByVal VarPtr(AllocString), SysAllocStringByteLen(0&, lngSize + lngSize), 4&
    '
End Function

'Opens a wave file, subsequently fills this object's properties with information
'and reads wave data into m_bytData byte array.
'Returns True on success, False on failure
Public Function OpenWave(strFileName As String) As Boolean
    '
    'See http://ccrma.stanford.edu/courses/422/projects/WaveFormat/ for more
    'information on the wave file format.
    '
    'On Error GoTo ErrHandler
    '
    Dim intFreeFile As Integer
    '
    intFreeFile = FreeFile
    '
    Open strFileName For Binary Access Read As intFreeFile
        'Read the first 44 bytes into the header structure
        Get intFreeFile, , Header
        'Prevent unnecessary accessing of UDT (not sure if this makes things any faster)
        m_lngWaveDataLength = Header.WaveDataLength
        'Allocate a string
        m_strData = AllocString(m_lngWaveDataLength)
        'Read the wave data into the string - We should be at byte 45
        Get intFreeFile, , m_strData
        'We don't need the file anymore, let's close it.
    Close #1
    'Make sure we actually have a WAVE file that we can read.
    If Not Header.RiffChunkID = RIFF Then
        '
        Err.Raise WM_APP + 1984, "clsWave.OpenWave", "The file is not in RIFF format."
        GoTo ErrHandler
        '
    ElseIf Not Header.RiffChunkFormat = WAVE Then
        '
        Err.Raise WM_APP + 1985, "clsWave.OpenWave", "The file is in RIFF format but is not a WAVE file."
        GoTo ErrHandler
        '
    ElseIf Not Header.RiffSubChunkID = FMT Then
        '
        Err.Raise WM_APP + 1986, "clsWave.OpenWave", "The file is a WAVE file, but the header is corrupt."
        GoTo ErrHandler
        '
    ElseIf Not Header.RiffSubChunk2ID = DATA Then
        '
        Err.Raise WM_APP + 1986, "clsWave.OpenWave", "The file is a WAVE file, but the header is corrupt."
        GoTo ErrHandler
        '
    End If
    'WaveDataLength = NumSamples * NumChannels * BitsPerSample/8
    m_lngWaveSamples = Header.WaveDataLength / Header.WaveNumChannels / (Header.WaveBitsPerSample / 8)
    m_lngWaveLength = (m_lngWaveSamples / Header.WaveSampleRate) * 1000
    'All done
    OpenWave = True
    '
Exit Function
    '
ErrHandler:
    '
    If Err.Number > WM_APP Then
        '
        OpenWave = False
        Close intFreeFile
        '
    Else
        '
        OpenWave = False
        Close intFreeFile
        '
        Err.Raise Err.Number, "clsWave.OpenWave", Err.Description
        '
    End If
    '
End Function



'PCM = 1 (i.e. Linear quantization) Values other than 1 indicate some form of compression.
Public Property Get WaveAudioFormat() As Integer
    '
    WaveAudioFormat = Header.WaveAudioFormat
    '
End Property

'PCM = 1 (i.e. Linear quantization) Values other than 1 indicate some form of compression.
Public Property Let WaveAudioFormat(ByVal intWaveAudioFormat As Integer)
    '
    Header.WaveAudioFormat = intWaveAudioFormat
    '
End Property

Public Property Get WaveChannels() As Integer
    '
    WaveChannels = Header.WaveNumChannels
    '
End Property

Public Property Let WaveChannels(ByVal intWaveChannels As Integer)
    '
    Header.WaveNumChannels = intWaveChannels
    '
End Property


Public Property Get WaveSampleRate() As Long
    '
    WaveSampleRate = Header.WaveSampleRate
    '
End Property

Public Property Let WaveSampleRate(ByVal lngWaveSampleRate As Long)
    '
    Header.WaveSampleRate = lngWaveSampleRate
    '
End Property

Public Property Get WaveByteRate() As Long
    '
    WaveByteRate = Header.WaveByteRate
    '
End Property

Public Property Let WaveByteRate(ByVal lngWaveByteRate As Long)
    '
    Header.WaveByteRate = lngWaveByteRate
    '
End Property

Public Property Get WaveBitsPerSample() As Integer
    '
    WaveBitsPerSample = Header.WaveBitsPerSample
    '
End Property

Public Property Let WaveBitsPerSample(ByVal intWaveBitsPerSample As Integer)
    '
    Header.WaveBitsPerSample = intWaveBitsPerSample
    '
End Property

'Length of the wave file in milliseconds
Public Property Get WaveLength() As Long
    '
    WaveLength = m_lngWaveLength
    '
End Property

Public Property Get WaveSamples() As Long
    '
    WaveSamples = m_lngWaveSamples
    '
End Property

'Copies byte array from object's array to external array
Public Function GetData() As Byte()
    'Only convert the string to a byte array once.
    If LenB(m_strData) > 0 Then
        'Convert the data string into a byte array
        m_bytData = StrConv(m_strData, vbFromUnicode)
        m_strData = vbNullString
        '
    End If
    '
    GetData = m_bytData
    '
End Function

Public Property Get WaveDataLength() As Long
    '
    WaveDataLength = Header.WaveDataLength
    '
End Property

'Writes header using this object's properties if optional parameters are not set, then
'writes data contained in bytArray.  If optional parameters are provided, the routine
'will use those values as the WAVE header.
'Returns True on success, False on failure
Public Function WriteWave(bytArray() As Byte, strFileName As String, Optional WaveNumChannels As Integer, Optional WaveSampleRate As Long, Optional WaveBitsPerSample As Long) As Boolean
    '
    On Error GoTo ErrHandler
    '
    Dim intFreeFile As Integer
    Dim WaveInfo As WaveHeader  'We use our own local header udt as not to disturb
                                'the module-level header udt if optional parameters
                                'are provided
    Dim lngWaveDataLength As Long
    Dim strData As String
    Dim bolOmitted As Boolean
    'Set wave header values in no particular order (well, they actually ARE in order)
    WaveInfo.RiffChunkID = RIFF
    WaveInfo.RiffChunkFormat = WAVE
    WaveInfo.RiffSubChunkID = FMT
    WaveInfo.RiffSubChunk2ID = DATA
    WaveInfo.WaveAudioFormat = 1 '1 for PCM/Uncompressed
    'Quickly determine if any of the optional parameters were omitted.
    bolOmitted = bolOmitted Or (WaveNumChannels = 0)
    bolOmitted = bolOmitted Or (WaveSampleRate = 0)
    bolOmitted = bolOmitted Or (WaveBitsPerSample = 0)
    'If the header parameters were omitted, we retrieve them from the header
    If bolOmitted Then
        '
        WaveNumChannels = Header.WaveNumChannels
        WaveSampleRate = Header.WaveSampleRate
        WaveBitsPerSample = Header.WaveBitsPerSample
        '
    End If
    '
    lngWaveDataLength = UBound(bytArray) + 1
    '
    WaveInfo.WaveNumChannels = WaveNumChannels
    WaveInfo.WaveSampleRate = WaveSampleRate
    WaveInfo.WaveByteRate = WaveSampleRate * WaveNumChannels * WaveBitsPerSample / 8
    WaveInfo.WaveBitsPerSample = WaveBitsPerSample
    WaveInfo.WaveBlockAlign = WaveInfo.WaveNumChannels * WaveInfo.WaveBitsPerSample / 8
    WaveInfo.WaveDataLength = lngWaveDataLength
    'This is the size of the rest of the chunk following this number.
    'Since this value is at byte 8, that means it's the value of the entire
    'file's size in bytes, minus 8.  Or, you could say it is the wave data
    'length plus 36.
    WaveInfo.RiffChunkSize = lngWaveDataLength + 36
    WaveInfo.RiffSubChunkSize = 16  '16 for PCM.
    'Convert our byte array into a string
    strData = StrConv(bytArray, vbUnicode)
    '
    intFreeFile = FreeFile
    'And now, we can simply write the header and data
    Open strFileName For Binary Access Write As intFreeFile
        'Write the header
        Put intFreeFile, , WaveInfo
        Put intFreeFile, , strData
        '
    Close intFreeFile
    'I would imagine VB does this cleaning for us.
    'strData = vbNullString
    '
    WriteWave = True
    '
Exit Function
    '
ErrHandler:
    '
    WriteWave = False
    'Err.Raise Err.Number, "clsWave.WriteWave", Err.Description
    '
End Function
