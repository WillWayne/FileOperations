Imports System
Imports System.Collections.Generic
Imports System.Text
Imports FileLibrary.Tiff.Enums

Namespace Tiff

    Public Class IFDEntry
        Implements IEquatable(Of IFDEntry)

#Region " Stateless Implementation "
#Region " Shared Members "
        Public Shared Empty As IFDEntry
#End Region

#Region " Shared Constructors "
        Shared Sub New()
            IFDEntry.Empty = New IFDEntry()
        End Sub
#End Region
#End Region

#Region " Properties "
        Public ReadOnly Property Tag() As IFDTag
            Get
                Return mTag
            End Get
        End Property

        Public ReadOnly Property Type() As IFDType
            Get
                Return mType
            End Get
        End Property

        Public ReadOnly Property Count() As Integer
            Get
                Return mCount
            End Get
        End Property

        Public ReadOnly Property ValueOffset() As Integer
            Get
                Return mValueOffset
            End Get
        End Property

        Public ReadOnly Property Value() As Byte()
            Get
                Return mValue
            End Get
        End Property

        Public ReadOnly Property ValueIsOffset() As Boolean
            Get
                If Me.Tag = IFDTag.StripOffsets Then Return True 'Always an offset value
                Return mValueIsOffset
            End Get
        End Property

        Public Property Endian() As EndianStyle
            Get
                Return mEndian
            End Get
            Set(ByVal value As EndianStyle)
                'Make sure this represents an actual change, and the new value is valid.
                If mEndian <> value AndAlso value.ToString() <> DirectCast(value, Integer).ToString() Then
                    ConvertEntry()
                End If
            End Set
        End Property
#End Region

#Region " Public Methods "
        Public Function GetEntry(ByVal endian As EndianStyle) As Byte()
            Dim tempEntry As Byte() = DirectCast(mEntry.Clone, Byte())

            If mEndian <> endian Then
                Array.Reverse(tempEntry, 0, 2)
                Array.Reverse(tempEntry, 2, 2)
                Array.Reverse(tempEntry, 4, 4)
                Array.Reverse(tempEntry, 8, 4)
            End If

            Return tempEntry
        End Function

        Public Sub SetEndian(ByVal endian As EndianStyle)
            Me.ConvertEntry(endian)
        End Sub

        Public Sub SetValueSegment(ByVal value As Integer)
            Me.mValueOffset = value
            TiffUtilities.WriteValueToArray(value, IFDType.LONG, 8, Me.mEntry, mEndian)
        End Sub

        Public Sub SetCount(ByVal value As Integer)
            If value * TiffUtilities.GetSizeOfType(Me.Type) > 4 Then
                mValueIsOffset = True
            Else
                mValueIsOffset = False
            End If

            TiffUtilities.WriteValueToArray(value, IFDType.SHORT, 2, Me.mEntry, mEndian)
        End Sub
#End Region

#Region " Interface Implementations "
#Region " IEquatable Implementation "
        Public Overloads Function Equals(ByVal other As IFDEntry) As Boolean Implements System.IEquatable(Of IFDEntry).Equals
            Dim reversed As Boolean = False
            Dim result As Boolean = True

            If Me.mEndian <> other.mEndian Then
                ConvertEntry()
                reversed = True
            End If

            If Not Me.mEntry.Equals(other.mEntry) Then result = False

            If reversed Then ConvertEntry()

            Return result
        End Function
#End Region
#End Region

#Region " Private Methods "
        Private Sub Clear()
            mTag = IFDTag.Unknown
            mType = IFDType.Unknown
            mCount = 0
            mValueOffset = 0
            mValue = New Byte(-1) {}
            mValueIsOffset = False
            mEndian = EndianStyle.LittleEndian
            mEntry = New Byte(11) {}
        End Sub

        Private Sub ConvertEntry(ByVal endian As EndianStyle)
            If mEndian <> endian Then
                mValue = GetEntry(endian)

                Dim valueBlockSize As Integer = TiffUtilities.GetSizeOfType(mType)
                If valueBlockSize > 1 Then
                    For i As Integer = 0 To (mValue.Length / valueBlockSize)
                        Array.Reverse(mValue, (i * valueBlockSize), valueBlockSize)
                    Next
                End If

                mEndian = endian
            End If
        End Sub

        Private Sub ConvertEntry()
            If mEndian = EndianStyle.BigEndian Then
                ConvertEntry(EndianStyle.LittleEndian)
            ElseIf mEndian = EndianStyle.LittleEndian Then
                ConvertEntry(EndianStyle.BigEndian)
            End If
        End Sub
#End Region

#Region " Constructors "
        ''' <summary>
        ''' Initializes an Empty IFDEntry
        ''' </summary>
        Private Sub New()
            Me.Clear()
        End Sub

        ''' <summary> 
        ''' Constructs a new IFDEntry object. 
        ''' </summary> 
        ''' <param name="entry">The 12-byte IFD Entry</param> 
        ''' <param name="endian">The Byte-ordering of the entry</param> 
        ''' <param name="openFile"> 
        ''' A reference to the open FileStream, allowing The IFDEntry to access the referenced 
        ''' data, if the Value portion of the Entry is a file pointer. 
        ''' </param> 
        Public Sub New(ByVal entry As Byte(), ByVal endian As EndianStyle, ByVal openFile As System.IO.FileStream)
            If entry.Length = 12 Then
                mEntry = entry
                mEndian = endian

                mTag = TiffUtilities.GetIfdTag(TiffUtilities.ByteToUShort(TiffUtilities.GetSubArray(entry, 0, 2), mEndian))
                mType = TiffUtilities.GetIfdType(TiffUtilities.ByteToUShort(TiffUtilities.GetSubArray(entry, 2, 2), mEndian))

                If mType <> IFDType.Unknown Then
                    mCount = TiffUtilities.ByteToInteger(TiffUtilities.GetSubArray(entry, 4, 4), mEndian)

                    Dim valueLength As Integer = TiffUtilities.GetSizeOfType(mType) * mCount
                    If valueLength <= 4 Then
                        mValue = TiffUtilities.GetSubArray(entry, 8, 4)
                    Else
                        mValueIsOffset = True
                        mValueOffset = TiffUtilities.ByteToInteger(TiffUtilities.GetSubArray(entry, 8, 4), mEndian)
                        mValue = TiffUtilities.CustomByteArray(valueLength)
                        If openFile.CanRead AndAlso openFile.CanSeek Then
                            openFile.Seek(mValueOffset, System.IO.SeekOrigin.Begin)
                            openFile.Read(mValue, 0, valueLength)
                        End If
                    End If
                End If
            End If
        End Sub
#End Region

#Region " Member Variables "
        Private mTag As IFDTag = IFDTag.Unknown
        Private mType As IFDType = IFDType.Unknown
        Private mCount As Integer = 0
        Private mValueOffset As Integer = 0
        Private mValue As Byte() = New Byte(-1) {}
        Private mValueIsOffset As Boolean = False
        Private mEndian As EndianStyle = EndianStyle.LittleEndian
        Private mEntry As Byte() = New Byte(11) {}
#End Region

    End Class
End NameSpace