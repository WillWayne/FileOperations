Imports FileLibrary.Tiff.Enums

Namespace Tiff
    Public Class TiffUtilities
        Public Shared LittleEndianHeader() As Byte = New Byte(3) {&H49, &H49, 42, 0}
        Public Shared BigEndianHeader() As Byte = New Byte(3) {&H4D, &H4D, 0, 42}

        Public Shared Function ByteToLong(ByVal data As Byte(), ByVal endian As EndianStyle) As Long
            If data.Length > 8 Then
                Throw New ArgumentException("Data must be at most 8 bytes long")
            End If
            Dim result As Long = 0

            If endian = EndianStyle.BigEndian Then
                For Each x As Byte In data
                    result = (result << 8) + CLng(x)
                Next
            ElseIf endian = EndianStyle.LittleEndian Then
                Dim count As Integer = 0
                For Each x As Byte In data
                    result += CLng(x) << (count * 8)
                    count += 1
                Next
            End If

            Return result
        End Function

        Public Shared Function ByteToUInteger(ByVal data As Byte(), ByVal endian As EndianStyle) As UInteger
            If data.Length > 4 Then
                Throw New ArgumentException("Data must be at most 4 bytes long")
            End If
            Return CUInt(ByteToLong(data, endian))

            'Dim result As UInteger = 0

            'If endian = EndianStyle.BigEndian Then
            '	For Each x As Byte In data
            '		result = (result << 8) + CUInt(x)
            '	Next
            'ElseIf endian = EndianStyle.LittleEndian Then
            '	Dim count As Integer = 0
            '	For Each x As Byte In data
            '		result += CUInt(x) << (count * 8)
            '		count += 1
            '	Next
            'End If

            'Return result
        End Function

        Public Shared Function ByteToInteger(ByVal data As Byte(), ByVal endian As EndianStyle) As Integer
            If data.Length > 4 Then
                Throw New ArgumentException("Data must be at most 4 bytes long")
            End If
            Return CInt(ByteToUInteger(data, endian))
        End Function

        Public Shared Function ByteToUShort(ByVal data As Byte(), ByVal endian As EndianStyle) As UShort
            If data.Length > 2 Then
                Throw New ArgumentException("Data must be at most 2 bytes long")
            End If
            Return CUShort(ByteToUInteger(data, endian))
        End Function

        Public Shared Function ByteToShort(ByVal data As Byte(), ByVal endian As EndianStyle) As Short
            If data.Length > 2 Then
                Throw New ArgumentException("Data must be at most 2 bytes long")
            End If
            Return CShort(ByteToInteger(data, endian))
        End Function

        Public Shared Function ByteToString(ByVal data As Byte()) As String
            Dim chars() As Char = New Char(data.Length - 1) {}

            For i As Integer = 0 To data.Length - 1
                chars(i) = CChar(ChrW(data(i)))
            Next

            Return New String(chars)
        End Function

        ''' <summary>
        ''' This generic method will convert a 1-dimensional byte array into an array of unsigned
        ''' integers based on both the endian style and the designated IFDType.
        ''' </summary>
        ''' <param name="data">The data to break into values.</param>
        ''' <param name="endian">The byte-order the individual values are stored in.</param>
        ''' <param name="type">The IFDType of the individual values.</param>
        Public Shared Function ByteToUIntArray(ByVal data As Byte(), ByVal endian As EndianStyle, ByVal type As IFDType) As UInteger()
            Dim increment As Short = TiffUtilities.GetSizeOfType(type)
            If increment > 4 Then Throw New ArgumentException("The specified type is too large to be cast as an Integer.")
            If data.Length Mod increment <> 0 Then
                Throw New IndexOutOfRangeException("The passed Byte array does not contain the correct number of elements")
            End If

            Dim values() As UInteger = New UInteger((data.Length / increment) - 1) {}

            For i As Integer = 0 To data.Length - 1 Step increment
                values(i / increment) = TiffUtilities.ByteToUInteger(TiffUtilities.GetSubArray(data, i, increment), endian)
            Next

            Return (values)

        End Function

        Public Shared Function ByteToIntArray(ByVal data As Byte(), ByVal endian As EndianStyle, ByVal type As IFDType) As Integer()
            Dim temp() As UInteger = ByteToUIntArray(data, endian, type)
            Dim ret() As Integer = New Integer(temp.Length - 1) {}

            For i As Integer = 0 To temp.Length
                ret(i) = CInt(temp(i))
            Next

            Return ret
        End Function

        Public Shared Function ByteToUShortArray(ByVal data As Byte(), ByVal endian As EndianStyle, ByVal type As IFDType) As UShort()
            If TiffUtilities.GetSizeOfType(type) > 2 Then Throw New ArgumentException("The specified type is too large to be cast as a Short.")

            Dim temp() As UInteger = ByteToUIntArray(data, endian, type)
            Dim ret() As UShort = New UShort(temp.Length - 1) {}

            For i As Integer = 0 To temp.Length
                ret(i) = CUShort(temp(i))
            Next

            Return ret
        End Function

        Public Shared Function ByteToShortArray(ByVal data As Byte(), ByVal endian As EndianStyle, ByVal type As IFDType) As Short()
            Dim temp() As UShort = ByteToUShortArray(data, endian, type)
            Dim ret() As Short = New Short(temp.Length - 1) {}

            For i As Integer = 0 To temp.Length
                ret(i) = CShort(temp(i))
            Next

            Return ret
        End Function

        Public Shared Function ByteArrayToSignedByteArray(ByVal data As Byte()) As SByte()
            Dim ret() As SByte = New SByte(data.Length - 1) {}

            Dim dataType As Type = ret(0).GetType()

            Dim index As Integer = 0
            For Each datum As Byte In data
                ret(index) = ByteToSByte(data(index))
                index += 1
            Next

            Return ret
        End Function

        Public Shared Function SByteArrayToUnsignedByteArray(ByVal data As Byte()) As Byte()
            Dim ret() As Byte = New Byte(data.Length - 1) {}

            Dim dataType As Type = ret(0).GetType()

            Dim index As Integer = 0
            For Each datum As Byte In data
                ret(index) = SByteToByte(data(index))
                index += 1
            Next

            Return ret
        End Function

        ''' <summary>
        ''' Converts an Unsigned Byte into the bit-equivalent Signed Byte
        ''' </summary>
        Public Shared Function ByteToSByte(ByVal unit As Byte) As SByte
            Return CSByte((unit And SByte.MaxValue) - (unit And SByte.MinValue))
        End Function

        ''' <summary>
        ''' Converts a Signed Byte into the bit-equivalent Unsigned Byte
        ''' </summary>
        Public Shared Function SByteToByte(ByVal unit As SByte) As Byte
            Return CByte((unit And Byte.MaxValue) + (unit And Byte.MinValue))
        End Function

        Public Shared Function GetSubArray(ByVal data As Byte(), ByVal offset As Integer, ByVal length As Integer) As Byte()
            If offset + length > data.Length Then
                'Throw New IndexOutOfRangeException()
                length = data.Length - offset
            End If

            Dim result As Byte() = New Byte(length - 1) {}
            For i As Integer = 0 To length - 1
                result(i) = data(i + offset)
            Next

            Return result
        End Function

        Public Shared Function GetSizeOfType(ByVal type As IFDType) As Short
            Select Case type
                Case IFDType.[BYTE], IFDType.ASCII, IFDType.[SBYTE], IFDType.UNDEFINED
                    Return 1
                Case IFDType.[SHORT], IFDType.SSHORT
                    Return 2
                Case IFDType.[LONG], IFDType.SLONG, IFDType.FLOAT
                    Return 4
                Case IFDType.RATIONAL, IFDType.SRATIONAL, IFDType.[DOUBLE]
                    Return 8
                Case Else
                    Return 0
            End Select
        End Function

        Public Shared Function TypeIsSigned(ByVal type As IFDType) As Boolean
            Select Case type
                Case IFDType.[SBYTE], IFDType.SSHORT, IFDType.SLONG, IFDType.SRATIONAL
                    Return True
            End Select

            Return False
        End Function

        Public Shared Function GetDotNetType(ByVal type As IFDType) As System.Type
            Select Case type
                Case IFDType.Unknown
                    Return GetType(Object)
                Case IFDType.[BYTE]
                    Return GetType(Byte)
                Case IFDType.ASCII
                    Return GetType(Byte)
                Case IFDType.[SHORT]
                    Return GetType(UShort)
                Case IFDType.[LONG]
                    Return GetType(UInteger)
                Case IFDType.RATIONAL
                    'BUG: This is incorrect!! Rational numbers use two Longs, the first is the
                    ' numerator, the second the denominator. Total bytes = 8.
                    Return GetType(ULong)
                Case IFDType.[SBYTE]
                    Return GetType(SByte)
                Case IFDType.UNDEFINED
                    Return GetType(Byte)
                Case IFDType.SSHORT
                    Return GetType(Short)
                Case IFDType.SLONG
                    Return GetType(Integer)
                Case IFDType.SRATIONAL
                    'BUG: This is incorrect!! SRational numbers use two SLongs, the first is the
                    ' numerator, the second the denominator. Total bytes = 8.
                    Return GetType(Long)
                Case IFDType.FLOAT
                    Return GetType(Single)
                Case IFDType.[DOUBLE]
                    Return GetType(Double)
                Case Else
                    Return GetType(Byte)
            End Select
        End Function


        Public Shared Function GetIfdTag(ByVal value As UShort) As IFDTag
            Dim result As IFDTag = 0
            Try
                result = value
            Catch generatedExceptionName As Exception
                result = IFDTag.Unknown
            End Try
            Return result
        End Function

        Public Shared Function GetIfdTag(ByVal value As Short) As IFDTag
            Dim result As IFDTag = 0
            Try
                result = value
            Catch generatedExceptionName As Exception
                result = IFDTag.Unknown
            End Try
            Return result
        End Function

        Public Shared Function GetIfdType(ByVal value As UShort) As IFDType
            Dim result As IFDType = 0
            Try
                result = value
            Catch generatedExceptionName As Exception
                result = IFDType.Unknown
            End Try
            Return result
        End Function

        Public Shared Function GetIfdType(ByVal value As Short) As IFDType
            Dim result As IFDType = 0
            Try
                result = value
            Catch generatedExceptionName As Exception
                result = IFDType.Unknown
            End Try
            Return result
        End Function

        Public Shared Function CustomByteArray(ByVal length As Integer) As Byte()
            Dim ms As New System.IO.MemoryStream(length)
            ms.SetLength(length)
            Dim result As Byte() = ms.ToArray()
            ms.Dispose()
            Return result
        End Function

        Public Shared Function ReverseByteArray(ByVal data As Byte()) As Byte()
            Array.Reverse(data, 0, data.Length)
            Return data
        End Function

        <Obsolete("Use the Array.Reverse() overload instead.", False)> _
        Public Shared Sub ReverseSubArray(ByRef data As Byte(), ByVal offset As Integer, ByVal length As Integer)
            'Array.Copy(ReverseByteArray(GetSubArray(data, offset, length)), offset, data, offset, length)
            Array.Reverse(data, offset, length)
        End Sub

        Public Shared Function ValidateEnumValue(ByVal value As Object) As Boolean
            If Not TypeOf (value) Is [Enum] Then Return False

            If value.ToString() = (DirectCast(value, Integer)).ToString() Then Return False

            Return True
        End Function


        Private Shared Function DecompressStrip(ByVal data As Byte, ByVal compression As CompressionValue) As Byte()
            Dim ret As Byte() = New Byte() {}

            Select Case compression
                Case CompressionValue.CCITT3_1D
                    Throw New NotSupportedException()
                Case CompressionValue.CCITT3_2D
                    Throw New NotSupportedException()
                Case CompressionValue.CCITT4
                    Throw New NotSupportedException()
                Case CompressionValue.JPEG
                    Throw New NotSupportedException()
                Case CompressionValue.LZW
                    Throw New NotSupportedException()
                Case CompressionValue.RLE
                    Throw New NotSupportedException()
                Case Else
                    Throw New NotSupportedException()
            End Select

            Return ret
        End Function

        Public Shared Function UnpackBits(ByVal data As Byte()) As Byte()
            Dim ms As New System.IO.MemoryStream
            Dim buffer As Byte()
            Dim count As Byte

            Dim index As Integer = 0
            Dim cont As Boolean = True
            While cont
                If data(index) = 128 Then
                    'Ignore this case.
                    index += 1
                ElseIf data(index) > 128 Then
                    'Expand repeated value
                    count = (ByteToSByte(data(index)) * -1) + 1
                    For i As Integer = 0 To count - 1
                        ms.WriteByte(data(index + 1))
                    Next
                    index += 2
                Else
                    'Copy literal array
                    count = CInt(data(index)) + 1
                    index += 1
                    buffer = GetSubArray(data, index, count)
                    ms.Write(buffer, 0, buffer.Length)
                    index += count
                End If

                If index >= (data.Length - 1) Then cont = False
            End While

            Return ms.ToArray()
        End Function

        Public Shared Function PackBits(ByVal data As Byte()) As Byte()
            Dim ms As New System.IO.MemoryStream

            Dim HomoCount As SByte = 0
            Dim currentByte As Byte = 0
            Dim HeteroCount As SByte = 0

            Dim StartIndex As Integer = 0
            Dim StopIndex As Integer = 0

            Dim buffer As Byte() = Nothing
            Dim buildBuffer As Boolean = True

            For i As Integer = 0 To data.Length - 1
                If buildBuffer Then
                    currentByte = data(i)

                    'Check the triplette
                    If data(i) = data(i + 1) AndAlso data(i + 1) = data(i + 2) Then
                        StopIndex = i - 1
                        buildBuffer = False
                        HeteroCount = StopIndex - StartIndex
                    End If

                    While HomoCount < 127
                        If data(i + HomoCount) = currentByte Then
                            HomoCount += 1
                        Else
                            Exit While
                        End If
                    End While

                    If HomoCount >= 3 Then
                        'Encode the compressed value
                    Else

                    End If
                    HomoCount = 0
                Else
                    If data(i) = currentByte Then
                        HomoCount -= 1
                    End If
                End If
            Next

            Return ms.ToArray()
        End Function

        Public Shared Function IntegerToByteArray(ByVal value As Integer, ByVal endian As EndianStyle) As Byte()
            Dim result As Byte() = New Byte(3) {}

            'Loop through each byte in the result array and set its value to the Byte cast of the last
            ' 8 bits in the integer, then right shift the integer by 8 bits.
            'This process produces an array in LittleEndian (Least significant byte first)
            For i As Integer = 0 To result.Length - 1
                result(i) = CByte(value Mod 256)
                value = value >> 8
            Next

            'If we want BigEndian, simply reverse the array.
            If endian = EndianStyle.BigEndian Then Array.Reverse(result)
            Return result
        End Function

        Public Shared Function IntegerToByte(ByVal value As Integer, ByVal endian As EndianStyle) As Byte
            Dim index As Integer = CInt(IIf(endian = EndianStyle.BigEndian, 3, 0))
            Return TiffUtilities.IntegerToByteArray(value, endian)(index)
        End Function

        Public Shared Function ShortToByteArray(ByVal value As Short, ByVal endian As EndianStyle) As Byte()
            Dim offset As Short = 0
            If endian = EndianStyle.BigEndian Then offset = 2
            Return TiffUtilities.GetSubArray(IntegerToByteArray(CInt(value), endian), offset, 2)
        End Function

        Public Shared Function WriteValueToArray(ByVal value As Integer, ByVal type As IFDType, ByVal offset As Integer, ByVal data As Byte(), ByVal endian As EndianStyle) As Byte()
            Dim size As Integer = TiffUtilities.GetSizeOfType(type)

            If size = 1 Then
                data(offset) = IntegerToByte(value, endian)
            ElseIf size = 2 Then
                Array.Copy(ShortToByteArray(value, endian), 0, data, offset, size)
            ElseIf size = 4 Then
                Array.Copy(IntegerToByteArray(value, endian), 0, data, offset, size)
            Else
                Throw New ArgumentOutOfRangeException("No support for types larger than 4 bytes (or less than 1)")
            End If

            Return data
        End Function

        ''' <summary>
        ''' Applies one Orientation Value to another, returning the resulting Orientation.
        ''' </summary>
        Public Shared Function UpdateOrientation(ByVal original As OrientationType, ByVal applied As OrientationType) As OrientationType
            Dim list As New List(Of Integer)(2)
            If CInt(original) < CInt(applied) Then
                list.Add(CInt(original))
                list.Add(CInt(applied))
            Else
                list.Add(CInt(applied))
                list.Add(CInt(original))
            End If

            'We only support values 1-8
            If list(0) < 1 OrElse list(1) >= 9 Then Return OrientationType.NotSupported
            'Any orientation applied to None/None equals that orientation
            If list(0) = 1 Then Return list(1)
            'Any orientation applied to itself results in None/None. Unless it's 90° or 270°. Then it's 180°
            If list(0) = list(1) Then
                If list(0) = 6 OrElse list(0) = 8 Then Return OrientationType.Rotate180FlipNone
                Return OrientationType.RotateNoneFlipNone
            End If

            'Add the values together to take advantage of certain patterns in the results table on my whiteboard
            Dim sum As Integer = list(0) + list(1)

            '
            Select Case (sum)
                Case 15, 3
                    Return 2
                Case 14, 2
                    Return 1
                Case 13, 5
                    Return 4
                Case 6, 4
                    Return 3
                Case 10
                    Return 7
                Case 9
                    Return 8
                Case 12
                    Select Case (list(0))
                        Case 4
                            Return 5
                        Case 5
                            Return 1
                        Case 6
                            Return 3
                    End Select
                Case 11
                    Select Case (list(0))
                        Case 3, 4
                            Return 6
                        Case 5
                            Return 2
                    End Select
                Case 8
                    Select Case (list(0))
                        Case 2
                            Return 5
                        Case 3
                            Return 8
                    End Select
                Case 7
                    Select Case (list(0))
                        Case 2
                            Return 6
                        Case 3
                            Return 2
                    End Select
            End Select

            Return OrientationType.NotSupported
        End Function

        Public Shared Function ByteArrayToIntegerArray(ByVal data As Byte(), ByVal type As IFDType, ByVal endian As EndianStyle) As Integer()
            Dim len As Integer = TiffUtilities.GetSizeOfType(type)
            If len > 4 Then Throw New NotSupportedException()

            Dim result() As Integer = New Integer((data.Length \ len)) {}

            For i As Integer = 0 To (data.Length \ len)
                result(i) = ByteToInteger(TiffUtilities.GetSubArray(data, i * len, len), endian)
            Next

            Return result
        End Function
    End Class
End NameSpace