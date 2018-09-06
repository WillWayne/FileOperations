Imports System
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Collections

''' <summary>
''' Manages single and multipage images.
''' </summary>
Public Class TiffManager

#Region "Properties"
	''' <summary>
	''' Gets the number of pages in the current image.
	''' </summary>
	Public ReadOnly Property Count() As Integer
		Get
			Return mPages
		End Get
	End Property

	''' <summary>
	''' Sets the TiffManager's current image.
	''' </summary>
	''' <value>A valid Image type (tiff, jpeg, bmp, etc)</value>
	Public WriteOnly Property Picture() As Image
		Set(ByVal value As Image)
			mImage = Nothing
			Try
				mImage = value
			Catch ex As Exception
			End Try
			GetPageNumber()
			mImages = SplitImage()
		End Set
	End Property

	''' <summary>
	''' 
	''' Gets an arraylist of MemoryStreams.  Each stream represents one page from the stored Image.
	''' </summary>
	Public ReadOnly Property ArrayList() As ArrayList
		Get
			Return mImages
		End Get
	End Property
#End Region

#Region "Public Methods"
	''' <summary>
	''' Resets the contents of the instance to default "empty" values.
	''' </summary>
	Public Sub Clear()
		mImage = Nothing
		mPages = 0
	End Sub

	''' <summary>
	''' Converts any supported image type into an <see cref="ArrayList"/> of <see cref="MemoryStream"/>s
	''' </summary>
	''' <param name="img">The image file you want to split.</param>
	''' <returns>Returns an ArrayList of MemoryStreams.
	''' Each MemoryStream in the array represents one page of the passed image.</returns>
	''' <remarks>Does not require an instance of the object.</remarks>
	Public Shared Function SplitImage(ByVal img As Image) As ArrayList
		Dim splitImages As ArrayList = New ArrayList
		Try
			Dim objGuid As Guid = img.FrameDimensionsList(0)
			Dim objDimension As FrameDimension = New FrameDimension(objGuid)

			Dim Pages As Integer = img.GetFrameCount(objDimension)

			Dim enc As Encoder = Encoder.Compression
			Dim int As Integer = 0

			Dim i As Integer
			For i = 0 To Pages - 1
				img.SelectActiveFrame(objDimension, i)
				Dim ep As EncoderParameters = New EncoderParameters(1)
				ep.Param(0) = New EncoderParameter(enc, EncoderValue.CompressionNone)
				Dim info As ImageCodecInfo = GetEncoderInfo("image/tiff")

				Dim ms As MemoryStream = New MemoryStream
				img.Save(ms, info, ep)
				splitImages.Add(ms)
			Next
		Catch ex As Exception
			Return Nothing
		End Try
		Return splitImages
	End Function

	''' <summary>
	''' Splits the image stored in the instance into an <see cref="ArrayList"/> of <see cref="MemoryStream"/>s
	''' </summary>
	''' <returns>Returns an ArrayList of MemoryStreams.
	''' Each MemoryStream in the array represents one page of the passed image.</returns>
	Public Function SplitImage() As ArrayList
		'Since an ArrayList of the original image is stored globally, is this function necessary as-is?
		' - Check to make sure it's not used internally as-is first, if so, consider making it Private
		Return SplitImage(mImage)
	End Function

	''' <summary>
	''' Extracts the specified page from the locally stored image.
	''' </summary>
	''' <param name="page">A positive <see cref="Integer"/> that indecates the page number to retrieve.</param>
	''' <returns>Returns the specified image as an <see cref="Image"/> object.</returns>
	''' <remarks>A 10 page image will have <c>[0-9]</c> as valid page numbers.
	''' Returns Nothing if an error occurs</remarks>
	Public Function GetPage(ByVal page As Integer) As Image
		If page >= 0 AndAlso page < Count Then
			Dim ms As MemoryStream = New MemoryStream
			Try
				'Dim objGuid As Guid = mImage.FrameDimensionsList(0)
				'Dim objDimension As FrameDimension = New FrameDimension(objGuid)

				'mImage.SelectActiveFrame(objDimension, page)
				'mImage.Save(ms, ImageFormat.Bmp)

				'Return Image.FromStream(ms)
				Return Image.FromStream(mImages(page))
			Catch ex As Exception
				Return Nothing
			End Try
		Else
			Return Nothing
		End If
	End Function

	''' <summary>
	''' Merges an <see cref="ArrayList"/> of <see cref="MemoryStream"/>s into a single MemoryStream using the <see cref="ImageFormat.Tiff"/> encoding.
	''' </summary>
	''' <param name="images">The <see cref="ArrayList"/> of <see cref="MemoryStream"/> encoded Images to merge.</param>
	''' <returns>Returns a single, multiframe Image as a MemoryStream.</returns>
	''' <remarks>Returns Nothing if an error occurs</remarks>
	Public Shared Function MergeImages(ByVal images As ArrayList) As MemoryStream
		If IsNothing(images) Then
			Return Nothing
		End If
		If images.Count = 1 Then
			Return images.Item(0)
		End If

		'Should this Try/Catch frame be removed?
		Try
			Dim ms As MemoryStream = New MemoryStream
			Dim img As Bitmap = Image.FromStream(images(0))
			Dim img2 As Bitmap
			Dim i As Integer = 1

			Dim myImageCodecInfo As ImageCodecInfo
			Dim myEncoder As Encoder
			Dim myEncoderParameter As EncoderParameter
			Dim myEncoderParameters As EncoderParameters

			myImageCodecInfo = GetEncoderInfo("image/tiff")
			myEncoder = Encoder.SaveFlag
			myEncoderParameters = New EncoderParameters(1)

			' Save the first page (frame).
			myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.MultiFrame))
			myEncoderParameters.Param(0) = myEncoderParameter
			img.Save(ms, myImageCodecInfo, myEncoderParameters)

			For i = 1 To images.Count - 1
				img2 = Image.FromStream(images(i))
				myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.FrameDimensionPage))
				myEncoderParameters.Param(0) = myEncoderParameter
				img.SaveAdd(img2, myEncoderParameters)
			Next

			myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.Flush))
			myEncoderParameters.Param(0) = myEncoderParameter
			img.SaveAdd(myEncoderParameters)

			Return ms
		Catch ex As Exception
			Return Nothing
		End Try
	End Function

	''' <summary>
	''' Merges the image stored in the current instance into a <see cref="MemoryStream"/>.
	''' </summary>
	''' <returns>Returns a single, multiframe Image as a MemoryStream.</returns>
	''' <remarks></remarks>
	Public Function MergeImages() As MemoryStream
		Return MergeImages(mImages)
	End Function

	''' <summary>
	''' 
	''' </summary>
	''' <param name="images"></param>
	''' <param name="location"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function MergeImages(ByVal images As ArrayList, ByVal location As String) As Boolean
		If IsNothing(images) Then
			Return False
		End If

		Dim img As Bitmap = Image.FromStream(images(0))
		Dim img2 As Bitmap
		Dim i As Integer = 1

		Dim myImageCodecInfo As ImageCodecInfo
		Dim myEncoder As Encoder
		Dim myEncoderParameter As EncoderParameter
		Dim myEncoderParameters As EncoderParameters

		myImageCodecInfo = GetEncoderInfo("image/tiff")
		myEncoder = Encoder.SaveFlag
		myEncoderParameters = New EncoderParameters(1)

		' Save the first page (frame).
		myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.MultiFrame))
		myEncoderParameters.Param(0) = myEncoderParameter
		img.Save(location, myImageCodecInfo, myEncoderParameters)

		myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.FrameDimensionPage))
		myEncoderParameters.Param(0) = myEncoderParameter

		For i = 1 To images.Count - 1
			img2 = Image.FromStream(images(i))
			img.SaveAdd(img2, myEncoderParameters)
		Next

		myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.Flush))
		myEncoderParameters.Param(0) = myEncoderParameter
		img.SaveAdd(myEncoderParameters)

		Return True
	End Function

	Public Sub MergeImages(ByVal location As String)
		MergeImages(mImages, location)
	End Sub

	Public Sub DeleteImage(ByVal index As Integer)
		If index >= 0 AndAlso index < Count Then
			Clear()
			Try
				mImages.RemoveAt(index)
				Picture = Image.FromStream(MergeImages(mImages))
			Catch ex As Exception
			End Try
		End If
	End Sub

	Public Sub InsertImage(ByVal index As Integer, ByVal img As Image)
		If index >= 0 AndAlso index < Count AndAlso Not IsNothing(img) Then
			Dim ms As MemoryStream = New MemoryStream
			img.Save(ms, ImageFormat.Tiff)
			Try
				mImages.Insert(index, ms)
				Picture = Image.FromStream(MergeImages(mImages))
			Catch ex As Exception
			End Try
		ElseIf index = Count Then
			Append(img)
		End If
	End Sub

	Public Sub ReplaceImage(ByVal index As Integer, ByVal img As Image)
		If index >= 0 AndAlso index < Count Then
			DeleteImage(index)
			InsertImage(index, img)
		End If
	End Sub

	Public Sub MoveImage(ByVal fromIndex As Integer, ByVal toIndex As Integer)
		If fromIndex >= 0 AndAlso toIndex >= 0 Then
			If Count > fromIndex AndAlso Count > toIndex Then
				Dim img As Image = Me.GetPage(fromIndex)
				DeleteImage(fromIndex)
				InsertImage(toIndex, img)
			End If
		End If
	End Sub

	Public Sub Append(ByVal img As Image)
		If Not IsNothing(mImage) Then
			Dim ms As MemoryStream = New MemoryStream
			img.Save(ms, ImageFormat.Tiff)
			mImages.Add(ms)
			Picture = Image.FromStream(MergeImages(mImages))
		Else
			Picture = img
		End If
	End Sub
#End Region

#Region "Private Methods"
	Private Sub GetPageNumber()
		Dim objGuid As Guid = mImage.FrameDimensionsList(0)
		Dim objDimension As FrameDimension = New FrameDimension(objGuid)

		mPages = mImage.GetFrameCount(objDimension)
	End Sub

	Private Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
		Dim encoders() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders
		Dim i As Integer
		For i = 0 To encoders.Length - 1
			If encoders(i).MimeType = mimeType Then
				Return encoders(i)
			End If
		Next
		Return Nothing
	End Function
#End Region

#Region "Constructors"
	Public Sub New()
		mImage = Nothing
	End Sub

	Public Sub New(ByVal img As Image)
		If Not IsNothing(img) Then
			'mImage = img
			'GetPageNumber()
			Picture = img
		End If
	End Sub

	Public Sub New(ByVal images As ArrayList)
		Try
			If Not IsNothing(images) Then
				mImages = images
				mImage = Image.FromStream(MergeImages(images))
				GetPageNumber()
			End If
		Catch ex As Exception
			mImage = Nothing
		End Try
	End Sub

	Public Sub New(ByVal value As String)
		If Not value = "" Then
			Try
				Picture = Image.FromFile(value)
			Catch ex As Exception
				mImage = Nothing
			End Try
		Else
			mImage = Nothing
		End If
	End Sub

#End Region

#Region "Member Variables"
	'Private mImageFileName As String
	Private mPages As Integer
	Private mImage As Bitmap
	Private mImages As ArrayList
#End Region
End Class
