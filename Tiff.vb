Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports FileLibrary.Tiff.Enums

Namespace Tiff

    ''' <summary>
    ''' Reads an entire Tiff image from file into memory
    ''' </summary>
    Public Class Tiff
        Implements IEquatable(Of Tiff)

#Region " Stateless Implementation "
#Region " Shared Members "
        Public Shared Empty As Tiff
#End Region

#Region " Shared Constructors "
        Shared Sub New()
            Tiff.Empty = New Tiff()
        End Sub
#End Region
#End Region

#Region " Properties "
        ''' <summary>
        ''' Gets the Dictionary of TiffPage objects in the Tiff
        ''' </summary>
        Public ReadOnly Property Pages() As Dictionary(Of Integer, TiffPage)
            Get
                Return mPages
            End Get
        End Property

        ''' <summary>
        ''' Gets the Dictionary of TiffInfo objects in the Tiff
        ''' </summary>
        Public ReadOnly Property PageInfos() As Dictionary(Of Integer, IFD)
            Get
                Return mTiffInfo.PageInfos
            End Get
        End Property

        ''' <summary>
        ''' Gets the number of pages present in the Tiff image
        ''' </summary>
        Public ReadOnly Property PageCount() As Integer
            Get
                Return mTiffInfo.PageInfos.Count
            End Get
        End Property

        ''' <summary>
        ''' Gets the CompressionValue of the current Tiff
        ''' </summary>
        ''' <remarks>
        ''' Each page in a Tiff must use a single compression, but a Tiff file does not need
        ''' to assert that a single compression be used by all pages.
        ''' 
        ''' Tiff images that use more than one compression will return the custom 
        ''' <see cref="CompressionValue.Multiple"/> value.  For the compression of a specific 
        ''' page in the Tiff, access the Compression property of that individual page.
        ''' </remarks>
        Public ReadOnly Property Compression() As CompressionValue
            Get
                Return mTiffInfo.Compression
            End Get
        End Property


        Public ReadOnly Property IsEmpty() As Boolean
            Get
                Return Me.Equals(Tiff.Empty)
            End Get
        End Property
#End Region

#Region " Public Methods "
        Public Function Page(ByVal pageNumber As Integer) As TiffPage
            If mPages.ContainsKey(pageNumber) Then
                Return mPages(pageNumber)
            End If
            Throw New IndexOutOfRangeException()
        End Function

        Public Function SaveAs(ByVal filename As String, ByVal overwrite As Boolean) As Boolean
            If File.Exists(filename) AndAlso Not overwrite Then Return False
            Dim currentOffset As Integer = 8
            Using fs As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None)
                Dim header As Byte() = IIf(mEndian = EndianStyle.BigEndian, TiffUtilities.BigEndianHeader, TiffUtilities.LittleEndianHeader)
                fs.Write(header, 0, 8)
                Dim firstOffset As Integer = 0
                Dim data As TiffPageStructure
                Dim prevIfdOffset As Integer = 0
                mPages2.Reverse()
                For i As Integer = 0 To mPages2.Count - 1
                    data = mPages2(i).AssembleImageData(fs.Length, prevIfdOffset)
                    If i = 0 Then firstOffset = data.IfdOffset
                    data.RawPageData.WriteTo(fs)
                    prevIfdOffset = data.IfdOffset

                    Dim data23232 As Byte() = New Byte() {}
                    fs.Write(data23232, 0, data23232.length)
                Next
                mPages2.Reverse() 'Unreverse
            End Using
        End Function
#End Region

#Region " Interface Implementations "
#Region " IEquatable Implementation "
        Public Overloads Function Equals(ByVal other As Tiff) As Boolean Implements System.IEquatable(Of Tiff).Equals

        End Function
#End Region
#End Region

#Region " Private Methods "
        Private Sub ProcessFile(ByVal filename As String)
            mTiffInfo = New TiffInfo(filename)
            If Not mTiffInfo.IsEmpty Then
                For Each entry As KeyValuePair(Of Integer, IFD) In mTiffInfo.PageInfos
                    mPages.Add(entry.Key, New TiffPage(entry.Value))
                Next
            End If
        End Sub

        Private Sub Clear()
            mFileName = String.Empty
            mEndian = EndianStyle.LittleEndian
            mTiffInfo = Nothing
            mPages = Nothing
        End Sub
#End Region

#Region " Constructors "
        Private Sub New()
            Me.Clear()
        End Sub

        Public Sub New(ByVal filename As String)
            If File.Exists(filename) Then
                ProcessFile(filename)
            End If
        End Sub
#End Region

#Region "Member Variables"
        Private mFileName As String = String.Empty
        Private mEndian As EndianStyle = EndianStyle.LittleEndian
        Private mTiffInfo As TiffInfo = Nothing
        Private mPages As New Dictionary(Of Integer, TiffPage)
        Private mPages2 As New List(Of TiffPage)
#End Region

    End Class
End NameSpace