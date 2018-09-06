Imports System
Imports System.IO

Public NotInheritable Class MimeSniffer

#Region " Shared Members "
    Public Shared ReadOnly Empty As New MimeTypeInfo()
#End Region

#Region " Public Methods "

    Public Shared Function GetMimeType(ByVal fileName As String, Optional ByVal lookIntoFile As Boolean = False) As MimeTypeInfo
        Dim file As New FileInfo(fileName)

        If Not lookIntoFile Then
            Return mPatterns.GetMimeTypeInfoByFileExtension(file.Extension)
        Else
            Using fs As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Return GetMimeType(fs)
            End Using
        End If
    End Function

    ''' <summary>
    ''' Gets the MIME type of the specified document object.
    ''' </summary>
    ''' <param name="obj">Document object whose MIME type to determine
    ''' (only Byte[] and FileStream are currently supported).</param>
    ''' <returns>MIME type information for the specified document object.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetMimeType(ByVal obj As Object) As MimeTypeInfo
        If obj Is Nothing Then
            Return Empty
        ElseIf Not TypeOf obj Is Byte() AndAlso Not TypeOf obj Is FileStream Then
            Throw New ArgumentException("Unsupported Object type")
        End If
        If TypeOf obj Is FileStream AndAlso Not DirectCast(obj, FileStream).CanRead AndAlso Not DirectCast(obj, FileStream).CanSeek Then
            Return Empty
        End If

        Dim match As Boolean = True
        For Each mti As MimeTypeInfo In mPatterns.MimeTypeInfos
            match = EvaluatePatternGroupCollection(obj, mti.PatternGroups)
            If match Then
                Return mti
            End If
        Next
        Return Empty
    End Function

    Public Shared Function GetMimeType(ByVal file As FileInfo, ByVal data As Byte()) As MimeTypeInfo
        Dim match As Boolean = True
        Dim MimeType As MimeTypeInfo = Nothing

        If Not file.Exists Then
            Return New MimeTypeInfo
        End If

        MimeType = GetMimeType(file.Name)
        If VerifyMimeType(MimeType, data) Then
            Return MimeType
        End If

        Return GetMimeType(data)
    End Function

    ''' <summary>
    ''' Verifies that the given document object has the specified MIME type.
    ''' </summary>
    ''' <param name="mimeType">Expected MIME type.</param>
    ''' <param name="obj"></param>
    ''' <returns>Boolean indicating whether the document object is of the expected type.</returns>
    ''' <remarks></remarks>
    Public Shared Function VerifyMimeType(ByVal mimeType As MimeTypeInfo, ByVal obj As Object) As Boolean
        If Not TypeOf obj Is Byte() AndAlso Not TypeOf obj Is FileStream Then
            Throw New ArgumentException("Unsupported Object type")
        End If
        If mimeType Is Nothing Then
            Return False
        End If
        Return EvaluatePatternGroupCollection(obj, mimeType.PatternGroups)
    End Function

#End Region

#Region " Private Methods "
    Private Shared Function IsText(ByVal data As Byte) As Boolean
        If data > 32 AndAlso data < 127 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Shared Function IsWhiteSpace(ByVal data As Byte) As Boolean
        Dim chars() As Byte = New Byte() {9, 10, 13, 32}

        For Each piece As Byte In chars
            If data = piece Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Shared Function IsTextFile(ByVal data As Byte()) As Boolean
        For Each piece As Byte In data
            If Not IsText(piece) OrElse Not IsWhiteSpace(piece) Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Shared Function EvaluatePatternGroup(ByVal obj As Object, ByVal patterns As MimeTypePatternGroup) As Boolean
        Dim match As Boolean = True

        For Each pattern As MimeTypePatternInfo In patterns.Patterns
            If pattern.PatternType = "String" Then
                match = EvaluateAsciiPattern(obj, pattern)
                If Not match Then
                    Return False
                End If
            ElseIf pattern.PatternType = "Decimal" Then
                match = EvaluateDecimalPattern(obj, pattern)
                If Not match Then
                    Return False
                End If
            End If
        Next

        Return True
    End Function

    Private Shared Function EvaluateAsciiPattern(ByVal obj As Object, ByVal pattern As MimeTypePatternInfo) As Boolean
        Dim dataLength As Integer = 0
        Dim IsByte As Boolean
        If TypeOf obj Is Byte() Then
            dataLength = DirectCast(obj, Byte()).Length
            IsByte = True
        ElseIf TypeOf obj Is FileStream Then
            dataLength = DirectCast(obj, FileStream).Length
            IsByte = False
        Else
            Throw New ArgumentException("Unsupported Object type")
        End If

        Dim length As Integer = pattern.Pattern.Length
        Dim offset As Integer = pattern.PatternOffset
        Dim range As Integer = pattern.PatternRange
        Dim sample As String = String.Empty

        If dataLength = 0 Then Return False

        If (offset + length) > dataLength Then
            Return False
        End If

        If pattern.PatternRange > 0 Then
            length += range
            If length > dataLength Then
                length = dataLength - offset
            End If
        End If

        If IsByte Then
            sample = System.Text.ASCIIEncoding.GetEncoding(1252).GetString(DirectCast(obj, Byte()), offset, length)
        Else
            Dim data() As Byte = New Byte(length - 1) {}
            Dim fs As FileStream = DirectCast(obj, FileStream)
            fs.Seek(offset, SeekOrigin.Begin)
            fs.Read(data, 0, length)
            sample = System.Text.ASCIIEncoding.GetEncoding(1252).GetString(data, 0, data.Length)
        End If

        If sample Is String.Empty Then
            Return False
        End If

        Return sample.Contains(pattern.Pattern)
    End Function

    Private Shared Function EvaluateDecimalPattern(ByVal obj As Object, ByVal pattern As MimeTypePatternInfo) As Boolean
        Dim array() As String = pattern.Pattern.Split(",")

        Dim dataLength As Integer = 0
        Dim IsByte As Boolean
        If TypeOf obj Is Byte() Then
            dataLength = DirectCast(obj, Byte()).Length
            IsByte = True
        ElseIf TypeOf obj Is FileStream Then
            dataLength = DirectCast(obj, FileStream).Length
            IsByte = False
        Else
            Throw New ArgumentException("Unsupported Object type")
        End If

        Dim length As Integer = array.Length
        Dim offset As Integer = pattern.PatternOffset
        Dim range As Integer = pattern.PatternRange
        Dim test() As Byte = New Byte(array.Length - 1) {}

        If (offset + length) > dataLength Then
            Return False
        ElseIf (offset + length + range) > dataLength Then
            range -= dataLength - (offset + length)
        End If

        Dim data As Byte()
        If IsByte Then
            data = DirectCast(obj, Byte())
        Else
            data = New Byte(length - 1) {}
            Dim fs As FileStream = DirectCast(obj, FileStream)
            fs.Seek(offset, SeekOrigin.Begin)
            DirectCast(obj, FileStream).Read(data, 0, length + range)
            offset = 0
        End If

        Dim match As Boolean = True
        For i As Integer = 0 To range
            match = True
            For j As Integer = 0 To length
                If data(i + offset) <> Convert.ToInt32(array(i)) Then
                    match = False
                    Exit For
                End If
            Next
            If match Then
                Exit For
            End If
            offset += 1
        Next

        Return match
    End Function

    Private Shared Function EvaluatePatternGroupCollection(ByVal obj As Object, ByVal groups As MimeTypePatternGroup()) As Boolean
        Dim match As Boolean = True

        For Each patternGroup As MimeTypePatternGroup In groups
            match = EvaluatePatternGroup(obj, patternGroup)
            If match Then
                Exit For
            End If
        Next

        Return match
    End Function

    Private Shared Function GetByteArrayFromFile(ByVal file As FileInfo, ByVal offset As Integer, ByVal length As Integer) As Byte()
        If Not file.Exists Then
            Return Nothing
        End If
        Dim fs As New FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
        Dim data() As Byte = GetByteArrayFromFile(fs, offset, length)
        fs.Close()
        fs.Dispose()
        Return data
    End Function

    Private Shared Function GetByteArrayFromFile(ByVal fs As FileStream, ByVal offset As Integer, ByVal length As Integer) As Byte()
        If Not fs Is Nothing OrElse _
           Not fs.CanRead OrElse _
           Not fs.CanSeek OrElse _
           offset < 0 OrElse _
           length < 0 OrElse _
           (offset + length) > fs.Length _
        Then
            Return Nothing
        End If

        Dim data() As Byte = New Byte(length - 1) {}
        fs.Seek(offset, SeekOrigin.Begin)
        fs.Read(data, 0, length)
        Return data
    End Function
#End Region

#Region " Constructors "
    Shared Sub New()
        mPatterns = New MimeTypesProfileReader().ReadProfile
    End Sub
#End Region

    Private Shared mPatterns As MimeTypesProfile
End Class
