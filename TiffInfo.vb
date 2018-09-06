Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports FileLibrary.Tiff.Enums

Namespace Tiff

    Public Class TiffInfo
        Implements IEquatable(Of TiffInfo)

#Region " Stateless Implementation "
#Region " Shared Members "
        Public Shared Empty As TiffInfo
#End Region

#Region " Shared Constructors "
        Shared Sub New()
            TiffInfo.Empty = New TiffInfo()
        End Sub
#End Region
#End Region

#Region " Properties "
        Public ReadOnly Property PageInfos() As Dictionary(Of Integer, IFD)
            Get
                Return mIfds
            End Get
        End Property

        Public ReadOnly Property PageCount() As Integer
            Get
                Return mIfds.Count
            End Get
        End Property

        Public ReadOnly Property Compression() As CompressionValue
            Get
                Dim val1 As CompressionValue = mIfds(1).Compression
                For Each entry As KeyValuePair(Of Integer, IFD) In mIfds
                    If val1 <> entry.Value.Compression Then
                        Return CompressionValue.Multiple
                    End If
                Next
                Return val1
            End Get
        End Property

        ''' <summary>
        ''' Returns the width of the widest page
        ''' </summary>
        Public ReadOnly Property Width() As UInteger
            Get
                Dim ret As Integer = 0
                For Each Item As KeyValuePair(Of Integer, IFD) In mIfds
                    If Item.Value.ImageWidth > ret Then ret = Item.Value.ImageWidth
                Next
                Return ret
            End Get
        End Property

        ''' <summary>
        ''' Returns the height of the tallest page
        ''' </summary>
        Public ReadOnly Property Height() As UInteger
            Get
                Dim ret As Integer = 0
                For Each Item As KeyValuePair(Of Integer, IFD) In mIfds
                    If Item.Value.ImageHeight > ret Then ret = Item.Value.ImageHeight
                Next
                Return ret
            End Get
        End Property

        ''' <summary>
        ''' Returns the highest BitDepth encountered in the TIFF
        ''' </summary>
        ''' <remarks>
        ''' This property returns the highest Bitdepth of each page because
        ''' there are fewer restrictions for images higher than 1bpp.
        ''' </remarks>
        Public ReadOnly Property BitDepth() As Integer
            Get
                Dim ret As Integer = 0
                For Each Item As KeyValuePair(Of Integer, IFD) In mIfds
                    If Item.Value.BitDepth > ret Then ret = Item.Value.BitDepth
                Next
                Return ret
            End Get
        End Property

        Public ReadOnly Property IsEmpty() As Boolean
            Get
                Return Me.Equals(TiffInfo.Empty)
            End Get
        End Property

        Public ReadOnly Property IsValid() As Boolean
            Get
                Return Not IsEmpty AndAlso Me.mIfdOffset < 8
            End Get
        End Property

        Public ReadOnly Property IsReadOnly() As Boolean
            Get
                Return False
            End Get
        End Property

        Public ReadOnly Property Thumbnail() As System.Drawing.Image
            Get
                If mIfds.Count > 0 Then
                    Return ImageConverter.GetScaledThumbnailImage(Me.GetPageImage(1), 120, 120)
                Else
                    Return Nothing
                End If
            End Get
        End Property
#End Region

#Region " Public Methods "
        Public Function GetPageIfd(ByVal pageNumber As Integer) As IFD
            If mIfds.ContainsKey(pageNumber) Then
                Return mIfds(pageNumber)
            End If
            Throw New IndexOutOfRangeException()
        End Function

        Public Function GetPageImage(ByVal pageNumber As Integer) As System.Drawing.Image
            If Not mIfds.ContainsKey(pageNumber) Then Throw New IndexOutOfRangeException()
            Return New TiffPage(mIfds(pageNumber)).Image
        End Function
#End Region

#Region " Interface Implementations "
#Region " IEquatable Implementation "
        ''' <summary>
        ''' Compares the IFD entries of each TiffInfo.
        ''' </summary>
        Public Overloads Function Equals(ByVal other As TiffInfo) As Boolean Implements System.IEquatable(Of TiffInfo).Equals
            If Me.mIfds.Count <> other.mIfds.Count Then Return False

            For i As Integer = 0 To mIfds.Count
                If Not mIfds(i).Equals(other.mIfds(i)) Then Return False
            Next

            Return True
        End Function
#End Region
#End Region

#Region " Private Methods "
        Private Sub ConvertEntry()

        End Sub

        Private Sub ProcessFile(ByVal filename As String)
            'Validate the file, open the filestream and continue 
            Dim mFS As FileStream = Nothing

            Try
                mFS = New FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read)
            Catch ex As Exception
                If mFS IsNot Nothing Then
                    mFS.Close()
                    mFS.Dispose()
                    mFS = Nothing
                End If

                Me.Clear()
                Exit Sub
            End Try

            mFS.Seek(0, SeekOrigin.Begin)
            mFS.Read(mHeader, 0, 8)

            If ProcessHeader() Then
                Dim pageNumber As Integer = 1
                Dim tempOffset As Integer = mIfdOffset
                Dim prevOffset As Integer = 0
                mIfds = New Dictionary(Of Integer, IFD)()

                While tempOffset > 0
                    mFS.Seek(tempOffset, SeekOrigin.Begin)
                    mIfds.Add(pageNumber, New IFD(mFS, tempOffset, prevOffset, mEndian))
                    prevOffset = tempOffset
                    tempOffset = mIfds(pageNumber).NextIfd
                    pageNumber += 1
                End While
            Else
                'There was an error processing the header, so make this instance empty.
                Me.Clear()
            End If

            If mFS IsNot Nothing Then
                mFS.Close()
                mFS.Dispose()
                mFS = Nothing
            End If
        End Sub

        ''' <summary> 
        ''' Reads the mHeader private variable to determine if it is a valid TIFF header 
        ''' </summary> 
        ''' <returns>True if the header is valid, false otherwise.</returns> 
        Private Function ProcessHeader() As Boolean
            'The first two bytes must be identical
            If mHeader(0) <> mHeader(1) Then Return False

            'First, cast the mEndian value to the value of the first byte
            'Then, if the String values of the mEndian object and the Byte are the same,
            '		we know that that value is undefined
            mEndian = mHeader(0)
            If mEndian.ToString() = mHeader(0).ToString() Then Return False

            'Next, ensure the next two bytes are equal to 'The Answer' (42)
            If TiffUtilities.ByteToShort(TiffUtilities.GetSubArray(mHeader, 2, 2), mEndian) <> 42 Then Return False

            'Assign the next 4 bytes as the numerical offset of the first IFD
            mIfdOffset = TiffUtilities.ByteToInteger(TiffUtilities.GetSubArray(mHeader, 4, 4), mEndian)
            Return True
        End Function

        ''' <summary>
        ''' Clears the current instance to equal the Empty instance
        ''' </summary>
        Private Sub Clear()
            mFileName = String.Empty
            mHeader = New Byte(7) {}
            mEndian = EndianStyle.LittleEndian
            mIfdOffset = 0
            mIfds = New dictionary(Of Integer, IFD)
        End Sub
#End Region

#Region " Constructors "
        ''' <summary>
        ''' Creates an Emtpy TiffInfo
        ''' </summary>
        ''' <remarks></remarks>
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
        Private mHeader As Byte() = New Byte(7) {}
        Private mEndian As EndianStyle = EndianStyle.LittleEndian
        Private mIfdOffset As Integer = 0
        Private mIfds As Dictionary(Of Integer, IFD) = Nothing
        Private mIsValid As Boolean = False
#End Region

    End Class
End NameSpace