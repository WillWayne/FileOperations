Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.InteropServices
Imports FileLibrary.Tiff.Enums

Namespace Tiff

    Public Class TiffPage
        Implements IEquatable(Of TiffPage)

#Region " Stateless Implementation "
#Region " Shared Members "
        Public Shared Empty As TiffPage
#End Region

#Region " Shared Constructors "
        Shared Sub New()
            TiffPage.Empty = New TiffPage()
        End Sub
#End Region
#End Region

#Region " Properties "
        Public ReadOnly Property Size() As System.Drawing.Size
            Get
                Return New System.Drawing.Size(Width, Height)
            End Get
        End Property

        Public ReadOnly Property Width() As UInteger
            Get
                Return mIFD.ImageWidth
            End Get
        End Property

        Public ReadOnly Property Height() As UInteger
            Get
                Return mIFD.ImageHeight
            End Get
        End Property

        Public ReadOnly Property BitDepth() As UInteger
            Get
                Return mIFD.BitDepth
            End Get
        End Property

        Public ReadOnly Property Compression() As CompressionValue
            Get
                Return Me.mIFD.Compression
            End Get
        End Property

        Public ReadOnly Property PhotometricInterpretation() As Integer
            Get
                Return 0
            End Get
        End Property

        Public ReadOnly Property Image() As Bitmap
            Get
                Using ms As MemoryStream = GenerateImageData()
                    If ms Is Nothing Then Return Nothing
                    Return Bitmap.FromStream(ms, True, True)
                End Using
            End Get
        End Property

        Public ReadOnly Property RawData() As MemoryStream
            Get
                Return GenerateImageData()
            End Get
        End Property
#End Region

#Region " Public Methods "
        Public Function SaveToFile(ByVal filename As String, ByVal overwrite As Boolean) As Boolean
            'If the file exists, but we're instructed not to overwrite, return false.
            If File.Exists(filename) AndAlso Not overwrite Then Return False

            Try
                Using fs As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None)
                    GenerateImageData().WriteTo(fs)	'Write the data into the FileStream
                    fs.Close()
                End Using
            Catch ex As Exception
                Return False
            End Try
            Return True
        End Function

        Public Sub Rotate(ByVal rotation As RotateFlipType)
            If rotation = RotateFlipType.Rotate90FlipX Then
                '
            End If
        End Sub
#End Region

#Region " Interface Implementations "
#Region " IEquatable Implementation "
        Public Overloads Function Equals(ByVal other As TiffPage) As Boolean Implements System.IEquatable(Of TiffPage).Equals
            Return mIFD.Equals(other.mIFD)
        End Function
#End Region
#End Region

#Region " Private Methods "
        Private Function GenerateImageData() As MemoryStream
            If mPageData Is Nothing Then mPageData = AssembleImageData(8, 0)
            If mPageData Is Nothing Then Return Nothing
            Dim data As New MemoryStream
            data.SetLength(8) 'Minumum starting length = 8 for the header.

            'Write the header
            Dim header As Byte() = CType(IIf(mIFD.Endian = EndianStyle.BigEndian, TiffUtilities.BigEndianHeader, TiffUtilities.LittleEndianHeader), Byte())
            data.Write(header, 0, header.Length)

            'Write the offset to the first IFD
            data.Seek(4, SeekOrigin.Begin)
            data.Write(TiffUtilities.IntegerToByteArray(mPageData.IfdOffset + 8, mIFD.Endian), 0, 4)

            'Write the actual page data to the temp stream.
            mPageData.RawPageData.WriteTo(data)
            'mImageData = data
            Return data
        End Function

        Private Sub Clear()
            mIFD = IFD.Empty
        End Sub

        Private Sub ProcessIfd(ByVal info As IFD)
            mIFD = info
            'TODO: Do stuff.  Like extract the Strips using the Strip offsets and assemble the page image
            ExtractPageStrips()
        End Sub

        Private Sub ProcessDimensions()

        End Sub

        Private Function TryGetIfdEntry(ByVal tag As IFDTag, ByRef entry As IFDEntry) As Boolean
            If mIFD.Contains(tag) Then
                entry = mIFD.GetIfdEntry(tag)
                Return True
            Else
                entry = Nothing
                Return False
            End If
        End Function

        Private Function ExtractPageStrips() As Boolean
            Dim offsets() As UInteger = mIFD.StripOffsets
            Dim offsetLengths() As UInteger = mIFD.StripByteCounts
            Dim buffer() As Byte

            Try
                Using fs As New FileStream(mIFD.Filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    For i As Integer = 0 To offsets.Length - 1
                        buffer = TiffUtilities.CustomByteArray(offsetLengths(i))
                        fs.Seek(offsets(i), SeekOrigin.Begin)
                        fs.Read(buffer, 0, offsetLengths(i))
                        mStrips.Add(buffer.Clone)
                    Next
                End Using
            Catch ex As Exception
                mStrips.Clear()
                Return False
            End Try

            Return True
        End Function

        Public Function AssembleImageData(ByVal offset As Integer, ByVal nextIfd As Integer) As TiffPageStructure
            'Write the data into here
            Dim ImageData As New MemoryStream
            'Use this as a temporary location for IFD entry values that are bigger than 4 bytes
            Dim Values As New MemoryStream

            'Store the Strip Offsets here
            Dim StripOffsets As New MemoryStream
            'Store the Strip sizes here
            Dim StripByteCounts As New MemoryStream

            If mStrips Is Nothing OrElse mStrips.Count = 0 Then
                Return Nothing
            End If

            'Write the strips to the MemoryStream and add their offsets to the temp list.
            For Each strip As Byte() In mStrips
                StripOffsets.Write(TiffUtilities.IntegerToByteArray(offset + ImageData.Length, mIFD.Endian), 0, 4)
                StripByteCounts.Write(TiffUtilities.IntegerToByteArray(strip.Length, mIFD.Endian), 0, 4)
                ImageData.Write(strip, 0, strip.Length)
            Next

            'Write the IFD
            Dim IfdOffset As Integer = ImageData.Length
            mOffsets.Add(IfdOffset)
            Dim ValuesOffset As Integer = offset + IfdOffset + 2 + (mIFD.Entries.Count * 12) + 4
            ImageData.Write(TiffUtilities.ShortToByteArray(mIFD.Entries.Count, mIFD.Endian), 0, 2)

            For Each entry As KeyValuePair(Of IFDTag, IFDEntry) In mIFD.Entries
                entry.Value.SetEndian(mIFD.Endian) 'Ensure the entry and all of its data are in the proper byte order

                'Very complicated!!  "StripOffsets" is always an offset, even if the count * sizeOfType > 4
                If entry.Value.ValueIsOffset Then
                    'Determine the next valid Offset value, the offset of the Values stream plus its current length.
                    Dim newOffset As Integer = ValuesOffset + Values.Length

                    If entry.Value.Tag = IFDTag.StripOffsets Then
                        If StripOffsets.Length <= 4 Then
                            'If there is only one Strip, the sole offset is written to the entry's value
                            ' by updating the newOffset object to point to the Strip data.
                            newOffset = TiffUtilities.ByteToInteger(StripOffsets.ToArray(), entry.Value.Endian)
                        Else
                            'If there are multiple offsets, write the offsets to the Values object
                            StripOffsets.WriteTo(Values)

                            mOffsets.AddRange(TiffUtilities.ByteArrayToIntegerArray(StripOffsets.ToArray(), IFDType.LONG, mIFD.Endian))
                        End If
                    ElseIf entry.Value.Tag = IFDTag.StripByteCounts Then
                        'If there are multiple offsets, write the offsets to the Values object
                        StripByteCounts.WriteTo(Values)
                    Else
                        Values.Write(entry.Value.Value, 0, entry.Value.Value.Length)
                    End If

                    entry.Value.SetValueSegment(newOffset)
                    mOffsets.Add(newOffset)
                Else
                    If entry.Value.Tag = IFDTag.StripOffsets Then
                        entry.Value.SetValueSegment(TiffUtilities.ByteToInteger(StripOffsets.ToArray(), entry.Value.Endian))
                    ElseIf entry.Value.Tag = IFDTag.StripByteCounts Then
                        entry.Value.SetValueSegment(TiffUtilities.ByteToInteger(StripByteCounts.ToArray(), entry.Value.Endian))
                    End If
                End If

                'Append the entry to the ImageData
                ImageData.Write(entry.Value.GetEntry(mIFD.Endian), 0, 12)
            Next

            ImageData.Write(TiffUtilities.IntegerToByteArray(nextIfd, mIFD.Endian), 0, 4)
            Values.WriteTo(ImageData)

            mStrips.Clear()
            'Set the offset of this IFD to the offset value passed in
            Return New TiffPageStructure(IfdOffset, ImageData)
        End Function

        Private Sub UpdateImageData(ByVal newPageOffset As Integer)
            'Cycle through the mOffsets list and update the offsets to be
            ' (passed PageOffset) + (Stored Value - mPageData.IfdOffset)
            'Update the IfdOffset.
        End Sub
#End Region

#Region " Constructors "
        Private Sub New()
            Me.Clear()
        End Sub

        Public Sub New(ByVal info As IFD)
            ProcessIfd(info)
        End Sub
#End Region

#Region "Member Variables "
        Private mIFD As IFD = IFD.Empty
        Private mIsValid As Boolean = True
        Private mStrips As New List(Of Byte())
        '' <summary>
        '' The sorted Page entry, including the IFD
        '' </summary>
        '' <remarks>This is intended to ease updates to the Page</remarks>
        'Private mImageData As System.IO.MemoryStream = Nothing
        ''' <summary>
        ''' The offset of the entire Page entry
        ''' </summary>
        Private mEntryOffset As Integer
        ''' <summary>
        ''' List of offset locations in the ImageData, to allow easy updates
        ''' </summary>
        ''' <remarks>
        ''' Intended to be used to 
        ''' </remarks>
        Private mOffsets As New List(Of Integer)
        ''' <summary>
        ''' Contains the 'Nuts and Bolts' of the Tiff Page itself - the entry data and the relative IFD offset
        ''' </summary>
        Private mPageData As TiffPageStructure = Nothing
#End Region
    End Class
End NameSpace