Namespace Tiff
    Public Class TiffPageStructure

        Private mIfdOffset As Integer
        ''' <summary>
        ''' Returns the Local Offset of the IFD in the PageData
        ''' </summary>
        Public ReadOnly Property IfdOffset() As Integer
            Get
                Return mIfdOffset
            End Get
        End Property

        Private mPageData As System.IO.MemoryStream
        ''' <summary>
        ''' Returns the Raw Page data, ready for insertion.  This data includes
        ''' the image data, IFD and data referenced by the IFD.  
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property RawPageData() As System.IO.MemoryStream
            Get
                Return mPageData
            End Get
        End Property

        Public Sub New(ByVal offset As Integer, ByVal data As System.IO.MemoryStream)
            mIfdOffset = offset
            mPageData = data
        End Sub

    End Class
End NameSpace