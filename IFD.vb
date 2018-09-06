Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports FileLibrary.Tiff.Enums

Namespace Tiff


    Public Class IFD
        Implements IEquatable(Of IFD)

#Region " Stateless Implementation "
#Region " Shared Members "
        Public Shared Empty As IFD
#End Region

#Region " Shared Constructors "
        Shared Sub New()
            IFD.Empty = New IFD()
        End Sub
#End Region
#End Region

#Region " Properties "
        Public ReadOnly Property PrevIfd() As Integer
            Get
                Return mParentOffset
            End Get
        End Property

        Public ReadOnly Property NextIfd() As Integer
            Get
                Return mNextIFD
            End Get
        End Property

        Public ReadOnly Property IsEmtpy() As Boolean
            Get
                Return mOffset = 0
            End Get
        End Property

        Public ReadOnly Property IsValid() As Boolean
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property Compression() As CompressionValue
            Get
                If Not mEntries.ContainsKey(IFDTag.Compression) Then
                    Return CompressionValue.Unknown
                End If

                Dim val As CompressionValue = TiffUtilities.ByteToUInteger(mEntries(IFDTag.Compression).Value, mEndian)
                If val.ToString() = CInt(val).ToString() Then
                    Return CompressionValue.Unknown
                End If
                Return val
            End Get
        End Property

        Public ReadOnly Property Entries() As Dictionary(Of IFDTag, IFDEntry)
            Get
                Return mEntries
            End Get
        End Property

        Public ReadOnly Property Filename() As String
            Get
                Return mFilename
            End Get
        End Property

        Public ReadOnly Property ImageSize() As System.Drawing.Size
            Get
                Return New System.Drawing.Size(ImageWidth, ImageHeight)
            End Get
        End Property

        Public ReadOnly Property ImageWidth() As UInteger
            Get
                Dim entry As IFDEntry = Nothing
                If mEntries.ContainsKey(IFDTag.ImageWidth) Then
                    entry = mEntries(IFDTag.ImageWidth)
                    If Not entry.ValueIsOffset Then	'This IFDEntry should only contain a short or an integer
                        Return TiffUtilities.ByteToUInteger(entry.Value, entry.Endian)
                    End If
                End If

                Return 0
            End Get
        End Property

        Public ReadOnly Property ImageHeight() As UInteger
            Get
                Dim entry As IFDEntry = Nothing
                If mEntries.ContainsKey(IFDTag.ImageLength) Then
                    entry = mEntries(IFDTag.ImageLength)
                    If Not entry.ValueIsOffset Then	'This IFDEntry should only contain a short or an integer
                        Return TiffUtilities.ByteToUInteger(entry.Value, entry.Endian)
                    End If
                End If

                Return 0
            End Get
        End Property

        Public ReadOnly Property BitDepth() As UInteger
            Get
                'The BitsPerSample entry is only optional for Bitonal images, so its absence
                ' implies 1bpp.
                If Not mEntries.ContainsKey(IFDTag.BitsPerSample) Then Return 1

                Dim data As Byte() = mEntries(IFDTag.BitsPerSample).Value
                Dim ret As Integer = 0

                'The BitsPerSample Value is Short data type, so increment by 2
                For i As Integer = 0 To data.Length - 1 Step 2
                    'Aggregate the short values.  For bitonal and grayscale(8bpp), a single short
                    ' contains the value (1, 4 or 8), but for component images (RGB, etc) we have 1 short
                    ' per pixel component (ex. (8 for Red, 8 for Green, 8 for Blue) = 24 bits per pixel)
                    ' The addition of each value is the images' total BitDepth.
                    ret += TiffUtilities.ByteToShort(TiffUtilities.GetSubArray(data, i, 2), mEndian)
                Next
                Return ret
            End Get
        End Property

        ''' <summary>
        ''' Returns an array of offsets (in order) for each of the strips that compose the
        ''' page's image.  Retreival of the data must also use the StripByteCounts information.
        ''' </summary>
        Public ReadOnly Property StripOffsets() As UInteger()
            Get
                If Me.mEntries.ContainsKey(IFDTag.StripOffsets) Then
                    Dim entry As IFDEntry = mEntries(IFDTag.StripOffsets)
                    Return TiffUtilities.ByteToUIntArray(entry.Value, mEndian, entry.Type)
                End If

                Return Nothing
            End Get
        End Property

        ''' <summary>
        ''' Returns an array of Strip sizes (in order) for each of the strips that compose the
        ''' page's image.  Retreival of the data must also use the StripOffsets information.
        ''' </summary>
        Public ReadOnly Property StripByteCounts() As UInteger()
            Get
                If Me.mEntries.ContainsKey(IFDTag.StripByteCounts) Then
                    Dim entry As IFDEntry = mEntries(IFDTag.StripByteCounts)
                    Return TiffUtilities.ByteToUIntArray(entry.Value, mEndian, entry.Type)
                End If

                Return Nothing
            End Get
        End Property

        Public ReadOnly Property Endian() As EndianStyle
            Get
                Return mEndian
            End Get
        End Property
#End Region

#Region " Public Methods "
        Public Function GetIfdEntry(ByVal tag As IFDTag) As IFDEntry
            If mEntries.ContainsKey(tag) Then
                Return mEntries(tag)
            End If
            Return Nothing
        End Function

        Public Function GetIFD(ByVal endian As EndianStyle) As Byte()
            Dim tempEntry As Byte() = mEntry.Clone()

            If mEndian <> endian Then
                Array.Reverse(mEntry, 0, 2)

                Dim offset As Integer = 2
                For Each pair As KeyValuePair(Of IFDTag, IFDEntry) In mEntries
                    Array.Copy(pair.Value.GetEntry(endian), 0, mEntry, offset, 12)
                    offset += 12
                Next

                Array.Reverse(mEntry, offset, 4)
            End If

            Return tempEntry
        End Function

        Public Function Contains(ByVal tag As IFDTag) As Boolean
            If mEntries.ContainsKey(tag) Then Return True
            Return False
        End Function
#End Region

#Region " Interface Implementations "
#Region " IEquatable Implementation "
        Public Overloads Function Equals(ByVal other As IFD) As Boolean Implements System.IEquatable(Of IFD).Equals

        End Function
#End Region
#End Region

#Region " Private Methods "
        Private Sub Clear()
            mParentOffset = 0
            mOffset = 0
            mNextIFD = 0
            mEntryCount = 0
            mEntry = New Byte(-1) {}
            mEndian = EndianStyle.LittleEndian
            mEntries = New Dictionary(Of IFDTag, IFDEntry)()
            mFilename = String.Empty
        End Sub

        Private Sub ConvertEntry()
            If mEndian = EndianStyle.BigEndian Then
                ConvertEntry(EndianStyle.LittleEndian)
            ElseIf mEndian = EndianStyle.LittleEndian Then
                ConvertEntry(EndianStyle.BigEndian)
            End If
        End Sub

        Private Sub ConvertEntry(ByVal endian As EndianStyle)
            If mEndian <> endian Then
                'Update all the IFD Entries first
                For Each pair As KeyValuePair(Of IFDTag, IFDEntry) In mEntries
                    pair.Value.Endian = endian
                Next

                'Update the internal Entry
                mEntry = GetIFD(endian)
                'Update the EndianStyle
                mEndian = endian
            End If
        End Sub
#End Region

#Region "Constructors"
        Private Sub New()
            Me.Clear()
        End Sub

        ''' <summary>
        ''' Creates a new IFD from the given file, starting at the given file offset.
        ''' </summary>
        ''' <param name="openFile">The file to extract the IFD from</param>
        ''' <param name="offset">The offset where the IFD begins</param>
        Public Sub New(ByVal openFile As FileStream, ByVal offset As Integer, ByVal parentOffset As Integer, ByVal endian As EndianStyle)
            'We must be able to both seek the stream and read from the stream.
            If openFile.CanRead AndAlso openFile.CanSeek AndAlso (offset > 0) Then
                mOffset = offset
                mParentOffset = parentOffset
                mFilename = openFile.Name

                'Get the IFD header, which determines the number of 12-byte entries that follow
                Dim count As Byte() = New Byte(1) {}
                openFile.Seek(mOffset, SeekOrigin.Begin)
                openFile.Read(count, 0, 2)
                mEntryCount = TiffUtilities.ByteToShort(count, endian)
                'Obtain and store the value
                'Calculate the fulle IFD entry size
                'Header
                'Entries
                'Footer
                Dim entrySize As Integer = 2 + (mEntryCount * 12) + 4

                'Read the entire IFD entry from disk at once (including the 2-byte header)
                mEntry = TiffUtilities.CustomByteArray(entrySize)
                openFile.Seek(mOffset, SeekOrigin.Begin)
                openFile.Read(mEntry, 0, entrySize)

                'Iterate through all the 12-byte entries
                Dim tempEntry As IFDEntry
                For i As Integer = 2 To (12 * mEntryCount) - 1 Step 12
                    tempEntry = New IFDEntry(TiffUtilities.GetSubArray(mEntry, i, 12), mEndian, openFile)
                    If Not mEntries.ContainsKey(tempEntry.Tag) Then mEntries.Add(tempEntry.Tag, tempEntry)
                Next

                'Grab the last 4 bytes of the entry to locate the offset of the next IFD.
                mNextIFD = TiffUtilities.ByteToInteger(TiffUtilities.GetSubArray(mEntry, (entrySize - 4), 4), mEndian)
            End If
        End Sub
#End Region

#Region "Member Variables"
        ''' <summary>
        ''' The file offset of this IFD's parent IFD (0 if none)
        ''' </summary>
        Private mParentOffset As Integer = 0

        ''' <summary>
        ''' The file offset of the beginning of this IFD (0 if this instance is Empty or Invalid)
        ''' </summary>
        Private mOffset As Integer = 0

        ''' <summary>
        ''' The file offset to this IFD's child IFD (0 if none)
        ''' </summary>
        Private mNextIFD As Integer = 0

        ''' <summary>
        ''' The number of 12-byte IFD Entries present in the IFD
        ''' </summary>
        Private mEntryCount As Short = 0

        ''' <summary>
        ''' The full in-file IFD entry, including the 2-byte header and 4-byte footer
        ''' </summary>
        Private mEntry As Byte() = New Byte(-1) {}

        ''' <summary>
        ''' The EndianStyle used by this IFD
        ''' </summary>
        Private mEndian As EndianStyle = EndianStyle.LittleEndian

        ''' <summary>
        ''' The collection of de-serialized IFD entries, with their Tags used as keys
        ''' </summary>
        Private mEntries As New Dictionary(Of IFDTag, IFDEntry)()

        ''' <summary>
        ''' The name of the Tiff file this IFD belongs to
        ''' </summary>
        Private mFilename As String = String.Empty
#End Region

    End Class
End NameSpace
