Imports System
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Collections

#Region " Revision 3 "
''' <summary>
''' Manages single and multipage images.
''' </summary>
Public Class ImageManager
	Implements IDisposable


#Region " Properties "
	''' <summary>
	''' Gets the number of pages in the current image.
	''' </summary>
	Public ReadOnly Property Count() As Integer
		Get
			If IsDirty Then
				Return mImages.Count
			End If
			mPages = ImageManager.GetPageNumber(mImage)
			Return mPages
		End Get
	End Property

	''' <summary>
	''' Sets the ImageManager's current image.
	''' </summary>
	''' <value>A valid Image type (tiff, jpeg, bmp, etc)</value>
	Public Property Image() As Image
		Get
			If Me.IsDirty And Not Me.PreserveReference Then
				Me.Save()
			End If

			Return mImage
		End Get
		Set(ByVal value As Image)
			Clear()
			mImage = value
			If value Is Nothing Then
				mPages = 0
			Else
				mPages = ImageManager.GetPageNumber(mImage)
			End If
		End Set
	End Property

	''' <summary>
	'''
	''' Gets an arraylist of MemoryStreams.  Each stream represents one page from the stored Image.
	''' </summary>
	Public ReadOnly Property ToArrayList() As List(Of Image)
		Get
			If mImages.Count > 0 Then
				Return mImages
			Else
				Return ImageManager.SplitImage(mImage)
			End If
		End Get
	End Property

	Public ReadOnly Property ToMemoryStream() As MemoryStream
		Get
			If IsDirty Then
				If KeepCurrent Then
					Save()
				Else
					Return ImageManager.MergeToStream(mImages)
				End If
			End If
			Return ImageManager.ImageToStream(mImage)
		End Get
	End Property

	Public ReadOnly Property IsDirty() As Boolean
		Get
			Return Me.mIsDirty
		End Get
	End Property

	Public Property KeepCurrent() As Boolean
		Get
			Return mKeepCurrent
		End Get
		Set(ByVal value As Boolean)
			mKeepCurrent = value
		End Set
	End Property

	Public Property NotifyOnChange() As Boolean
		Get
			Return mNotifyOnChange
		End Get
		Set(ByVal value As Boolean)
			mNotifyOnChange = value
		End Set
	End Property

	Public Property PreserveReference() As Boolean
		Get
			Return mNotifyOnChange
		End Get
		Set(ByVal value As Boolean)
			mNotifyOnChange = value
		End Set
	End Property
#End Region

#Region " Shared Methods "
	''' <summary>
	''' Extracts an image from a Multipage image file, such as a TIFF or an animated GIF.
	''' </summary>
	''' <param name="originalimage">The Multipage image to extract a frame from.</param>
	''' <param name="index">The 0-based index at which to extract the image frame.</param>
	''' <returns>Returns a <see cref="Stream"/> object.</returns>
	''' <remarks>
	''' This function will return <c>Nothing</c> if an error occurs,
	''' or if the passed index is out of range.
	''' </remarks>
	Public Shared Function GetPageStream(ByVal originalImage As Image, ByVal index As Integer) As System.IO.Stream
		Dim pages As Integer = GetPageNumber(originalImage)
		If index >= 0 AndAlso index < pages Then
			Dim ms As MemoryStream = New MemoryStream
			Try
				Dim objGuid As Guid = originalImage.FrameDimensionsList(0)
				Dim objDimension As FrameDimension = New FrameDimension(objGuid)

				originalImage.SelectActiveFrame(objDimension, index)
				originalImage.Save(ms, originalImage.RawFormat)

				Return ms
			Catch ex As Exception
				Return Nothing
			End Try
		Else
			Return Nothing
		End If
	End Function

	''' <summary>
	''' Extracts an image from a Multipage image file, such as a TIFF or an animated GIF.
	''' </summary>
	''' <param name="originalimage">The Multipage image to extract a frame from.</param>
	''' <param name="index">The 0-based index at which to extract the image frame.</param>
	''' <returns>Returns an <see cref="Image"/> object.</returns>
	''' <remarks>
	''' This function will return <c>Nothing</c> if an error occurs,
	''' or if the passed index is out of range.
	''' </remarks>
	Public Shared Function GetPageImage(ByVal originalImage As Image, ByVal index As Integer) As Image
		Dim ms As MemoryStream = GetPageStream(originalImage, index)
		If ms Is Nothing Then
			Return Nothing
		Else
			Return Image.FromStream(ms)
		End If
	End Function

	''' <summary>
	''' Gets the number of pages contained within the passed <see cref="Image"/>.
	''' </summary>
	Public Shared Function GetPageNumber(ByVal originalImage As Image) As Integer
		If originalImage Is Nothing Then
			Return 0
		End If
		Dim objGuid As Guid = originalImage.FrameDimensionsList(0)
		Dim objDimension As FrameDimension = New FrameDimension(objGuid)

		Return originalImage.GetFrameCount(objDimension)
	End Function

	Public Shared Function GetPageNumber(ByVal filename As String) As Integer
		Dim file As FileInfo
		Dim result As Integer = 0
		Try
			file = New FileInfo(filename)
			If file.Exists Then
				Dim img As Image = Image.FromFile(file.FullName)
				result = ImageManager.GetPageNumber(img)
				img = Nothing
			End If
		Catch ex As Exception
			Return 0
		End Try
		Return result
	End Function

	''' <summary>
	''' Gets a range of single page images from a multipage image.
	''' </summary>
	''' <param name="originalImage">The image to extract the pages from.</param>
	''' <param name="startingIndex">The index at which to begin extraction; 0-based.</param>
	''' <param name="maxNumberOfPages">The total number of Images to extract.</param>
	''' <returns>Returns a <see cref="List(of Image)"/> that contains the extracted Images.</returns>
	''' <remarks>
	''' The number of actual images contained in the <see cref="List(of Image)"/> depends on
	''' the number of valid images within the specified range.
	''' </remarks>
	Public Shared Function GetPageWindow(ByVal originalImage As Image, ByVal startingIndex As Integer, ByVal maxNumberOfPages As Integer) As List(Of Image)
		If maxNumberOfPages < 1 Then
			Return New List(Of Image)
		End If
		If startingIndex < 0 Then
			startingIndex = 0
		End If

		Dim splitImages As New List(Of Image)
		Dim Pages As Integer = ImageManager.GetPageNumber(originalImage)
		If startingIndex = 0 AndAlso Pages = 1 Then
			splitImages.Add(originalImage)
			ImageManager.OnProgressUpdate(originalImage)
			Return splitImages
		End If
		Dim ImageType As String = ImageManager.GetImageMimeType(originalImage)
		Dim index As Integer = startingIndex
		Dim length As Integer = startingIndex + maxNumberOfPages

		If originalImage Is Nothing OrElse Pages < startingIndex Then
			Return New List(Of Image)
		ElseIf Pages < length Then
			length = Pages
		End If

		Try
			Dim img As Image = Nothing

			Dim i As Integer
			For i = startingIndex To length - 1
				img = ImageManager.GetPageImage(originalImage, i)
				ImageManager.OnProgressUpdate(img)
				splitImages.Add(img)
			Next
		Catch ex As Exception
			Throw New Exception("Unable to split this Image")
		End Try
		Return splitImages
	End Function

	''' <summary>
	''' Splits a multiframe image into a <see cref="List(of Image)"/>.
	''' </summary>
	''' <param name="originalimage">The multiframe image to split.</param>
	''' <returns>
	''' Returns a <see cref="List(of Image)"/> containing the frames of the original image in their original order.
	''' </returns>
	Public Shared Function SplitImage(ByVal originalImage As Image) As List(Of Image)
		Dim Pages As Integer = ImageManager.GetPageNumber(originalImage)
		Return ImageManager.GetPageWindow(originalImage, 0, Pages)
	End Function

	''' <summary>
	''' Splits a <see cref="MemoryStream"/> containing a multiframe image into a <see cref="List(of Image)"/>.
	''' </summary>
	''' <param name="ms"></param>
	''' <returns>
	''' Returns a <see cref="List(of Image)"/> that contains the
	''' </returns>
	Public Shared Function SplitImage(ByVal ms As MemoryStream) As List(Of Image)
		Return SplitImage(Image.FromStream(ms, True, True))
	End Function

	''' <summary>
	''' Merges a Uniform <see cref="List(Of T)"/> into a <see cref="Stream"/>.
	''' </summary>
	''' <typeparam name="T"></typeparam>
	''' <param name="list"></param>
	''' <returns></returns>
	''' <remarks>
	''' Accepted types for <paramref name="T"/> are:
	''' <list>
	''' <item><see cref="Image"/></item>
	''' <item><see cref="MemoryStream"/></item>
	''' <item><see cref="FileStream"/></item>
	''' </list>
	'''
	''' <c>Nothing</c> will be returned if the list is empty or if it is of an
	''' unsupported type.
	''' </remarks>
	Public Shared Function MergeToStream(Of T)(ByVal list As List(Of T)) As System.IO.Stream
		'Return Nothing if Nothing is passed, bypassing all other logic.
		If list Is Nothing OrElse list.Count = 0 Then
			Return Nothing
		End If

		'Return Nothing if
		If Not TypeOf list(0) Is Image AndAlso _
		   Not TypeOf list(0) Is Bitmap AndAlso _
		   Not TypeOf list(0) Is MemoryStream AndAlso _
		   Not TypeOf list(0) Is FileStream AndAlso _
		   Not TypeOf list(0) Is FileInfo AndAlso _
		   Not TypeOf list(0) Is String Then
			Return Nothing
		End If

		'' New code to support merging multi-page images 12/7/2006
		Dim ImageList As New List(Of Image)
		Dim tempList As New List(Of Image)

		' Create a temporary list of Images (compound or otherwise) from the passed list
		For Each obj As Object In list
			tempList.Add(ImageManager.ObjectToImage(obj))
		Next

		' Cycle through the temporary images, adding them to the ImageList list
		For Each img As Image In tempList
			If img IsNot Nothing Then
				If ImageManager.GetPageNumber(img) > 1 Then
					' If the image has multiple pages, split the image and add its pages
					'   one by one to the ImageList
					Dim newList As List(Of Image) = ImageManager.SplitImage(img)
					For Each x As Image In newList
						ImageList.Add(x)
					Next
				Else
					'Otherwise, we have a single image that should be added to the ImageList
					ImageList.Add(img)
				End If
			End If
		Next
		'' End new code

		'  ** Commented line segments are in support of new code 12/7/2006
		'       Reinstate them to nullify the new code

		'If the first image is nothing, keep trying until a valid image is found
		' or the entire list is searched
		Dim image As System.Drawing.Image = ImageList(0) 'ObjectToImage(list.Item(0))
		Dim j As Integer = 1
		While image Is Nothing
			If j >= ImageList.Count Then 'list.Count Then
				Exit While
			End If
			image = ImageList(j) 'ObjectToImage(list.Item(j))
			j += 1
		End While

		If image Is Nothing Then
			Return Nothing
		End If

		Dim ms As Stream = New MemoryStream

		If ImageList.Count - j > 0 Then	'list.Count - j > 0 Then
			Dim image2 As Image = Nothing
			Dim i As Integer

			Dim myImageCodecInfo As ImageCodecInfo
			Dim myEncoder As Encoder
			Dim myEncoderParameter As EncoderParameter
			Dim myEncoderParameters As EncoderParameters

			Try
				myImageCodecInfo = GetEncoderInfo("image/tiff")
				myEncoder = Encoder.SaveFlag
				myEncoderParameters = New EncoderParameters(1)

				' Save the first page (frame) into the MemoryStream.
				myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.MultiFrame))
				myEncoderParameters.Param(0) = myEncoderParameter
				image.Save(ms, myImageCodecInfo, myEncoderParameters)


				'Save the remaining pages into the MemoryStream.
				For i = j To ImageList.Count - 1 'list.Count - 1
					image2 = ImageList(i) 'ObjectToImage(list.Item(i))
					If Not image2 Is Nothing Then
						myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.FrameDimensionPage))
						myEncoderParameters.Param(0) = myEncoderParameter
						image.SaveAdd(image2, myEncoderParameters)
					End If
				Next

				'Save the EncoderParameters to the MemoryStream.
				myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.Flush))
				myEncoderParameters.Param(0) = myEncoderParameter
				image.SaveAdd(myEncoderParameters)

				Return ms
			Catch ex As Exception
				Return Nothing
			Finally

				For Each Item As Image In ImageList
					Item.Dispose()
				Next

				For Each Item As Image In tempList
					Item.Dispose()
				Next

				image.Dispose()
				image2.Dispose()
			End Try
		Else
			'Dim enc As EncoderParameter =
			image.Save(ms, ImageManager.GetImageFormat(image))
			Return ms
		End If
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <typeparam name="T"></typeparam>
	''' <param name="list"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function MergeToImage(Of T)(ByVal list As List(Of T)) As Image
		If list.Count > 0 Then
			Return Image.FromStream(MergeToStream(list), True, True)
		End If
		Return Nothing
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <typeparam name="T"></typeparam>
	''' <param name="list"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function MergeToByteArray(Of T)(ByVal list As List(Of T)) As Byte()
		Return CType(MergeToStream(list), MemoryStream).ToArray
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="list"></param>
	''' <param name="index"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function DeletePageList(ByVal list As List(Of Image), ByVal index As Integer) As List(Of Image)
		If Not list Is Nothing Then
			If index >= 0 Then
				If index < list.Count Then
					list.RemoveAt(index)
				End If
			End If
		End If
		Return list
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="originalimage"></param>
	''' <param name="index"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function DeletePageStream(ByVal originalImage As Image, ByVal index As Integer) As System.IO.Stream
		Dim list As List(Of Image) = ImageManager.SplitImage(originalImage)
		Return ImageManager.MergeToStream(ImageManager.DeletePageList(list, index))
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="originalImage"></param>
	''' <param name="index"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function DeletePageImage(ByVal originalImage As Image, ByVal index As Integer) As Image
		Return Image.FromStream(ImageManager.DeletePageStream(originalImage, index))
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="list"></param>
	''' <param name="index"></param>
	''' <param name="page"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function InsertPageList(ByVal list As List(Of Image), ByVal index As Integer, ByVal page As Image) As List(Of Image)
		If page Is Nothing Then
			Return list
		End If

		Dim pagenum As Integer = ImageManager.GetPageNumber(page)


		If list Is Nothing Then
			list = New List(Of Image)
			Return ImageManager.AppendPageList(list, page)
		End If

		If index >= 0 AndAlso index < list.Count Then
			If pagenum = 1 Then
				list.Insert(index, page)
			ElseIf pagenum > 1 Then
				Dim i As Integer = 0
				Dim pagelist As List(Of Image) = ImageManager.SplitImage(page)
				For Each img As Image In pagelist
					list.Insert(index, pagelist.Item(i))
					index += 1
					i += 1
				Next
			End If
		ElseIf index = list.Count Then
			list = ImageManager.AppendPageList(list, page)
		End If

		Return list
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="originalimage"></param>
	''' <param name="index"></param>
	''' <param name="page"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function InsertPageStream(ByVal originalImage As Image, ByVal index As Integer, ByVal page As Image) As System.IO.Stream
		Dim list As List(Of Image) = ImageManager.SplitImage(originalImage)
		ImageManager.InsertPageList(list, index, page)
		Return ImageManager.MergeToStream(list)
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="originalImage"></param>
	''' <param name="index"></param>
	''' <param name="page"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function InsertPageImage(ByVal originalImage As Image, ByVal index As Integer, ByVal page As Image) As Image
		Return Image.FromStream(ImageManager.InsertPageStream(originalImage, index, page))
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="list"></param>
	''' <param name="moveFrom"></param>
	''' <param name="moveTo"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function MovePageList(ByVal list As List(Of Image), ByVal moveFrom As Integer, ByVal moveTo As Integer) As List(Of Image)
		If moveFrom = moveTo Then
			Return list
		End If
		Dim image = list.Item(moveFrom)
		ImageManager.DeletePageList(list, moveFrom)
		Return ImageManager.InsertPageList(list, moveTo, image)
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="originalimage"></param>
	''' <param name="movefrom"></param>
	''' <param name="moveto"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function MovePageStream(ByVal originalImage As Image, ByVal moveFrom As Integer, ByVal moveTo As Integer) As System.IO.Stream
		If moveFrom = moveTo Then
			Return ImageManager.ImageToStream(originalImage)
		End If
		Dim list As List(Of Image) = ImageManager.SplitImage(originalImage)
		ImageManager.MovePageList(list, moveFrom, moveTo)
		Return ImageManager.MergeToStream(list)
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="originalImage"></param>
	''' <param name="moveFrom"></param>
	''' <param name="moveTo"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function MovePageImage(ByVal originalImage As Image, ByVal moveFrom As Integer, ByVal moveTo As Integer) As Image
		If moveFrom = moveTo Then
			Return originalImage
		End If
		Return Image.FromStream(ImageManager.MovePageStream(originalImage, moveFrom, moveTo))
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="list"></param>
	''' <param name="index"></param>
	''' <param name="page"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function ReplacePageList(ByVal list As List(Of Image), ByVal index As Integer, ByVal page As Image) As List(Of Image)
		ImageManager.DeletePageList(list, index)
		Return ImageManager.InsertPageList(list, index, page)
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="originalimage"></param>
	''' <param name="index"></param>
	''' <param name="page"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function ReplacePageStream(ByVal originalImage As Image, ByVal index As Integer, ByVal page As Image) As System.IO.Stream
		Dim list As List(Of Image) = ImageManager.SplitImage(originalImage)
		ImageManager.ReplacePageList(list, index, page)
		Return ImageManager.MergeToStream(list)
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="originalImage"></param>
	''' <param name="index"></param>
	''' <param name="page"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function ReplacePageImage(ByVal originalImage As Image, ByVal index As Integer, ByVal page As Image) As Image
		Return Image.FromStream(ReplacePageStream(originalImage, index, page))
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="list"></param>
	''' <param name="page"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function AppendPageList(ByVal list As List(Of Image), ByVal page As Image) As List(Of Image)
		If page Is Nothing Then
			Return list
		End If

		If list Is Nothing Then
			list = New List(Of Image)
		End If

		Dim pagenum As Integer = ImageManager.GetPageNumber(page)
		If pagenum = 1 Then
			list.Add(page)
		ElseIf pagenum > 1 Then
			Dim pagelist As List(Of Image) = ImageManager.SplitImage(page)
			For Each img As Image In pagelist
				list.Add(img)
			Next
		End If

		Return list
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="originalimage"></param>
	''' <param name="page"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function AppendPageStream(ByVal originalImage As Image, ByVal page As Image) As System.IO.Stream
		Dim list As List(Of Image) = ImageManager.SplitImage(originalImage)
		ImageManager.AppendPageList(list, page)
		Return ImageManager.MergeToStream(list)
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="originalImage"></param>
	''' <param name="page"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function AppendPageImage(ByVal originalImage As Image, ByVal page As Image) As Image
		Return Image.FromStream(ImageManager.AppendPageStream(originalImage, page))
	End Function

	''' <summary>
	''' Accepts an <see cref="Object"/> and converts it into an <see cref="Image"/>.
	''' </summary>
	''' <param name="obj">An image expressed as one of several different <see cref="Type"/>s.</param>
	''' <returns>Returns the passed <see cref="Object"/> expressed as an <see cref="Image"/></returns>
	''' <remarks>
	''' This method supports the following <see cref="Object"/> <see cref="Type"/>s:
	''' 1. <see cref="Bitmap"/>
	''' 2. <see cref="Image"/>
	''' 3. <see cref="MemoryStream"/>
	''' 4. <see cref="FileStream"/>
	''' 5. <see cref="FileInfo"/>
	''' 6. <see cref="String"/> (as a file name and path like a <see cref="FileInfo.FullName"/>)
	''' </remarks>
	Public Shared Function ObjectToImage(ByVal obj As Object) As Image
		Dim img As Image = Nothing
		Dim file As New FileInfo("C:\" + Guid.NewGuid.ToString + ".uno")

		If TypeOf obj Is Byte() Then
			Return Image.FromStream(New MemoryStream(DirectCast(obj, Byte())))
		ElseIf TypeOf obj Is System.Drawing.Bitmap Then
			Return CType(CType(obj, Bitmap), Image)
		ElseIf TypeOf obj Is System.Drawing.Image Then
			Return CType(obj, Image)
		ElseIf TypeOf obj Is System.IO.MemoryStream Then
			Try
				img = System.Drawing.Image.FromStream(CType(obj, MemoryStream), True, True)
			Catch ex As Exception
				Throw New FormatException("[ImageManager] The supplied MemoryStream is not a valid image.")
			End Try
		ElseIf TypeOf obj Is System.IO.FileStream Then
			Try
				img = System.Drawing.Image.FromStream(CType(obj, FileStream), True, True)
			Catch ex As Exception
				Throw New FormatException("[ImageManager] The supplied FileStream is not a valid image.")
			End Try
		ElseIf TypeOf obj Is FileInfo Then
			file = CType(obj, FileInfo)
		ElseIf TypeOf obj Is System.String Then
			' make sure file exists and is accessible
			If System.IO.File.Exists(CStr(obj)) Then
				file = New FileInfo(CType(obj, String))
			Else
				Throw New System.IO.FileNotFoundException("Specified file does not exist or is not accessible: " + obj)
			End If
		Else
			Return Nothing
		End If

		If file.Exists Then
			Try
				img = System.Drawing.Image.FromFile(file.FullName)
			Catch ex As Exception
				Throw New FormatException("[ImageManager] The specified file is not a valid image.")
			End Try
		Else
			Return Nothing
		End If

		Return img
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="obj"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function GetImageMimeType(ByVal obj As Object) As String
		Dim image As Image = ObjectToImage(obj)
		If image Is Nothing Then
			Return String.Empty
		End If

		Dim bmpFormat As ImageFormat = image.RawFormat
		Dim strFormat As String = String.Empty

		If (bmpFormat.Equals(ImageFormat.Bmp)) Then
			strFormat = "image/bmp"
		ElseIf (bmpFormat.Equals(ImageFormat.Emf)) Then
			strFormat = "windows/metafile"
		ElseIf (bmpFormat.Equals(ImageFormat.Exif)) Then
			'An Exif is a jpeg with extra information
			strFormat = "image/jpeg" 'System.Net.Mime.MediaTypeNames.Image.Jpeg
		ElseIf (bmpFormat.Equals(ImageFormat.Gif)) Then
			strFormat = "image/gif"	'System.Net.Mime.MediaTypeNames.Image.Gif
		ElseIf (bmpFormat.Equals(ImageFormat.Icon)) Then
			strFormat = "image/x-icon"
		ElseIf (bmpFormat.Equals(ImageFormat.Jpeg)) Then
			strFormat = "image/jpeg" 'System.Net.Mime.MediaTypeNames.Image.Jpeg
		ElseIf (bmpFormat.Equals(ImageFormat.MemoryBmp)) Then
			strFormat = "image/bmp"
		ElseIf (bmpFormat.Equals(ImageFormat.Png)) Then
			strFormat = "image/png"
		ElseIf (bmpFormat.Equals(ImageFormat.Tiff)) Then
			strFormat = "image/tiff" 'System.Net.Mime.MediaTypeNames.Image.Tiff
		ElseIf (bmpFormat.Equals(ImageFormat.Wmf)) Then
			strFormat = "windows/metafile"
		End If

		Return strFormat
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="image"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function ImageToStream(ByVal image As Image) As System.IO.Stream
		Dim ms As New MemoryStream
		image.Save(ms, GetImageFormat(image))
		Return ms
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="image"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function GetImageFormat(ByVal image As Image) As System.Drawing.Imaging.ImageFormat
		Dim bmpFormat As ImageFormat = image.RawFormat
		Dim strFormat As String = "unidentified format"

		If (bmpFormat.Equals(ImageFormat.Bmp)) Then
			Return ImageFormat.Bmp
		ElseIf (bmpFormat.Equals(ImageFormat.Emf)) Then
			Return ImageFormat.Emf
		ElseIf (bmpFormat.Equals(ImageFormat.Exif)) Then
			Return ImageFormat.Exif
		ElseIf (bmpFormat.Equals(ImageFormat.Gif)) Then
			Return ImageFormat.Gif
		ElseIf (bmpFormat.Equals(ImageFormat.Icon)) Then
			Return ImageFormat.Icon
		ElseIf (bmpFormat.Equals(ImageFormat.Jpeg)) Then
			Return ImageFormat.Jpeg
		ElseIf (bmpFormat.Equals(ImageFormat.MemoryBmp)) Then
			Return ImageFormat.Bmp
		ElseIf (bmpFormat.Equals(ImageFormat.Png)) Then
			Return ImageFormat.Png
		ElseIf (bmpFormat.Equals(ImageFormat.Tiff)) Then
			Return ImageFormat.Tiff
		ElseIf (bmpFormat.Equals(ImageFormat.Wmf)) Then
			Return ImageFormat.Wmf
		End If

		Return Nothing
	End Function

	Public Shared Function GetImageFormat(ByVal mimeType As String) As Imaging.ImageFormat
		If (mimeType.Equals("image/bmp")) Then
			Return ImageFormat.Bmp
		ElseIf (mimeType.Equals("windows/metafile")) Then
			Return ImageFormat.Emf
		ElseIf (mimeType.Equals("image/jpeg")) Then
			Return ImageFormat.Exif
		ElseIf (mimeType.Equals("image/gif")) Then
			Return ImageFormat.Gif
		ElseIf (mimeType.Equals("image/x-icon")) Then
			Return ImageFormat.Icon
		ElseIf (mimeType.Equals("image/jpeg")) Then
			Return ImageFormat.Jpeg
		ElseIf (mimeType.Equals("image/bmp")) Then
			Return ImageFormat.Bmp
		ElseIf (mimeType.Equals("image/png")) Then
			Return ImageFormat.Png
		ElseIf (mimeType.Equals("image/tiff")) Then
			Return ImageFormat.Tiff
		ElseIf (mimeType.Equals("windows/metafile")) Then
			Return ImageFormat.Wmf
		End If

		Return Nothing
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="mimeType"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
		Dim encoders() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders
		Dim i As Integer
		For i = 0 To encoders.Length - 1
			If encoders(i).MimeType = mimeType Then
				Return encoders(i)
			End If
		Next
		Return Nothing
	End Function

	''' <summary>
	'''
	''' </summary>
	''' <param name="image"></param>
	''' <param name="mimetype"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function TryGetImageMimeType(ByVal image As Image, ByRef mimetype As String, ByVal exOut As Exception) As Boolean
		mimetype = GetImageMimeType(image)
		If mimetype = String.Empty Then
			Return False
		Else
			Return True
		End If
	End Function

	Public Shared Function WriteImageToDisk(ByVal imageToSave As Image, ByVal filename As String, ByVal overwrite As Boolean, Optional ByVal compressionEncoder As EncoderValue = EncoderValue.CompressionCCITT4) As Boolean
		Dim file As FileInfo
		Dim fs As FileStream = Nothing
		Dim result As Boolean = False
		Try
			file = New FileInfo(filename)
			If Not file.Directory.Exists Then
				If file.Directory.Parent.Exists Then
					file.Directory.Create()
				Else
					Return False
				End If
			ElseIf file.Exists Then
				If Not overwrite Then
					Return False
				Else
					file.Delete()
				End If
			End If

			Dim pages As Integer = ImageManager.GetPageNumber(imageToSave)
			fs = New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None)

			'If pages = 1 Then
			'	imageToSave.Save(fs, ImageManager.GetImageFormat(imageToSave))
			'ElseIf pages > 1 Then

			Dim objGuid As Guid = imageToSave.FrameDimensionsList(0)
			Dim objDimension As FrameDimension = New FrameDimension(objGuid)
			imageToSave.SelectActiveFrame(objDimension, 0)

			Dim img As Image = imageToSave 'ImageManager.GetPageImage(imageToSave, 0)
			Dim i As Integer = 1

			Dim myImageCodecInfo As ImageCodecInfo
			Dim myEncoder As Encoder
			Dim myEncoderParameter As EncoderParameter
			Dim myEncoderParameters As EncoderParameters
			Dim myEncoderValue As EncoderValue

			myImageCodecInfo = GetEncoderInfo("image/tiff")
			myEncoder = Encoder.SaveFlag
			myEncoderParameters = New EncoderParameters(2)
			myEncoderValue = EncoderValue.MultiFrame

			Dim myCompressionEncoder As Encoder = Encoder.Compression
			Dim myCompressionEncoderValue As EncoderValue = compressionEncoder
			Dim myCompressionEncoderParameter As EncoderParameter = _
			  New EncoderParameter(myCompressionEncoder, myCompressionEncoderValue)

			' Save the first page (frame).
			myEncoderParameter = New EncoderParameter(myEncoder, myEncoderValue)
			myEncoderParameters.Param(0) = myEncoderParameter
			myEncoderParameters.Param(1) = myCompressionEncoderParameter

			img.Save(fs, myImageCodecInfo, myEncoderParameters)

			myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.FrameDimensionPage))
			myEncoderParameters.Param(0) = myEncoderParameter
			For i = 1 To pages - 1
				imageToSave.SelectActiveFrame(objDimension, i)
				img.SaveAdd(imageToSave, myEncoderParameters)
			Next

			myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.Flush))
			myEncoderParameters.Param(0) = myEncoderParameter
			img.SaveAdd(myEncoderParameters)

			img = Nothing
			result = True
			'End If
		Catch ex As Exception
			result = False
		Finally
			If fs IsNot Nothing Then
				fs.Close()
				fs.Dispose()
			End If
		End Try
		Return result
	End Function

	Public Shared Function SelectCompression(ByVal img As Image) As EncoderValue
		Return EncoderValue.CompressionCCITT3
		If img.PixelFormat = Imaging.PixelFormat.Format1bppIndexed Then
			Return EncoderValue.CompressionCCITT4
		Else
			Return EncoderValue.CompressionNone
		End If
	End Function

	Public Shared Function UpdateImageOnDisk(ByVal filename As String, ByVal changes As Dictionary(Of Integer, Image), ByVal overwrite As Boolean) As Boolean
		Dim file As FileInfo
		Dim fs As FileStream = Nothing
		Dim originalImage As Image = Nothing
		Dim result As Boolean = False
		Try
			file = New FileInfo(filename)
			If Not file.Directory.Exists Then
				Return False
			ElseIf file.Exists AndAlso Not overwrite Then
				Return False
			End If

			file.CopyTo(file.FullName + ".tmp")
			originalImage = Image.FromFile(filename + ".tmp", True)

			Dim pages As Integer = ImageManager.GetPageNumber(originalImage)
			fs = New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None)

			If pages = 1 AndAlso Not changes.Count > 1 Then
				If changes.Count = 1 AndAlso changes.Item(0) IsNot Nothing Then
					changes.Item(0).Save(fs, ImageManager.GetImageFormat(changes.Item(0)))
				Else
					originalImage.Save(fs, ImageManager.GetImageFormat(originalImage))
				End If
			ElseIf pages > 1 Then

				Dim img As Image = ImageManager.GetUpToDateImage(originalImage, changes, 0)
				Dim i As Integer = 1

				Dim myImageCodecInfo As ImageCodecInfo
				Dim myEncoder As Encoder
				Dim myEncoderParameter As EncoderParameter
				Dim myEncoderParameters As EncoderParameters
				Dim myEncoderValue As EncoderValue

				myImageCodecInfo = GetEncoderInfo("image/tiff")
				myEncoder = Encoder.SaveFlag
				myEncoderParameters = New EncoderParameters(1)
				myEncoderValue = EncoderValue.MultiFrame

				' Save the first page (frame).
				myEncoderParameter = New EncoderParameter(myEncoder, myEncoderValue)
				myEncoderParameters.Param(0) = myEncoderParameter

				img.Save(fs, myImageCodecInfo, myEncoderParameters)

				myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.FrameDimensionPage))
				myEncoderParameters.Param(0) = myEncoderParameter
				For i = 1 To pages - 1
					img.SaveAdd(ImageManager.GetUpToDateImage(originalImage, changes, i), myEncoderParameters)
					'img.fr()
				Next

				myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.Flush))
				myEncoderParameters.Param(0) = myEncoderParameter
				img.SaveAdd(myEncoderParameters)

				img = Nothing
				result = True
			End If
		Catch ex As Exception
			result = False
		Finally
			If fs IsNot Nothing Then
				originalImage.Dispose()
				originalImage = Nothing
				System.IO.File.Delete(filename + ".tmp")
				fs.Close()
				fs.Dispose()
			End If
		End Try
		Return result
	End Function

	Public Shared Function UpdateImageOnDisk(ByVal filename As String, ByVal images As List(Of FileInfo), _
	 ByVal overwrite As Boolean, _
	 Optional ByVal compressionEncoder As EncoderValue = EncoderValue.CompressionCCITT4) As Boolean
		Dim omissions As New List(Of Integer)
		Return ImageManager.UpdateImageOnDisk(filename, images, omissions, overwrite, compressionEncoder)
	End Function

	Public Shared Function UpdateImageOnDisk(ByVal filename As String, ByVal images As List(Of FileInfo), _
	 ByVal omissions As List(Of Integer), ByVal overwrite As Boolean, _
	 Optional ByVal compressionEncoder As EncoderValue = EncoderValue.CompressionCCITT4) As Boolean
		Dim fs As FileStream = Nothing
		Dim img As Image = Nothing
		Dim og As Image = Nothing
		Dim file As FileInfo = Nothing
		Dim originalFile As FileInfo = Nothing
		Dim result As Boolean = False
		Try
			file = New FileInfo(filename + ".tmp")

			If file.Exists AndAlso file.IsReadOnly Then Return False
			If Not file.Directory.Exists Then file.Directory.Create()
			'If Not file.Exists Then file.Create()

			If images.Count = 1 Then
				If images(0).Exists Then
					images(0).CopyTo(filename, overwrite)
					result = True
				Else
					result = False
				End If
			ElseIf images.Count > 1 Then
				Dim startPoint As Integer = images.Count
				For i As Integer = 0 To images.Count - 1
					If Not omissions.Contains(i) Then
						startPoint = i
						Exit For
					End If
				Next
				originalFile = New FileInfo(filename)

				If Not originalFile.Exists Then
					og = Nothing
				Else
					og = Image.FromFile(originalFile.FullName, True)
				End If

				Dim frame As Image = Nothing

				If images(startPoint).Exists Then
					img = Image.FromFile(images(startPoint).FullName, True)
				Else
					If og IsNot Nothing Then img = ImageManager.GetPageImage(og, startPoint)
				End If
				If img Is Nothing Then Return False

				fs = New FileStream(file.FullName, FileMode.Create, FileAccess.Write, FileShare.None)

				Dim myImageCodecInfo As ImageCodecInfo = GetEncoderInfo("image/tiff")

				Dim myEncoder As Encoder = Encoder.SaveFlag
				Dim myEncoderValue As EncoderValue = EncoderValue.MultiFrame
				Dim myEncoderParameter As EncoderParameter = New EncoderParameter(myEncoder, myEncoderValue)

				Dim myCompressionEncoder As Encoder = Encoder.Compression
				Dim myCompressionEncoderValue As EncoderValue = compressionEncoder
				Dim myCompressionEncoderParameter As EncoderParameter = _
				  New EncoderParameter(myCompressionEncoder, myCompressionEncoderValue)

				Dim myEncoderParameters As EncoderParameters = New EncoderParameters(2)
				Dim mySafeEncoderParameters As EncoderParameters = New EncoderParameters(1)

				' Save the first page (frame).
				mySafeEncoderParameters.Param(0) = myEncoderParameter
				myEncoderParameters.Param(0) = myEncoderParameter
				myEncoderParameters.Param(1) = myCompressionEncoderParameter
				Try
					img.Save(fs, myImageCodecInfo, myEncoderParameters)
				Catch ex As Exception
					img.Save(fs, myImageCodecInfo, mySafeEncoderParameters)
				End Try

				'Prepare and save the remaining pages.
				myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.FrameDimensionPage))
				myEncoderParameters.Param(0) = myEncoderParameter
				mySafeEncoderParameters.Param(0) = myEncoderParameter

				startPoint += 1
				For i As Integer = startPoint To images.Count - 1
					If omissions.Contains(i) Then
						frame = Nothing
					ElseIf images(i).Exists Then
						frame = Image.FromFile(images(i).FullName, True)
					Else
						If og IsNot Nothing Then
							frame = ImageManager.GetPageImage(og, i)
						Else
							frame = Nothing
						End If
					End If
					If frame IsNot Nothing Then
						Try
							img.SaveAdd(frame, myEncoderParameters)
						Catch ex As Exception
							img.SaveAdd(frame, mySafeEncoderParameters)
						End Try
					End If
				Next

				myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.Flush))
				myEncoderParameters.Param(0) = myEncoderParameter
				img.SaveAdd(myEncoderParameters)
				result = True
			End If
		Catch ex As Exception
			result = False
		Finally
			If img IsNot Nothing Then img.Dispose()
			If og IsNot Nothing Then og.Dispose()
			If fs IsNot Nothing Then
				fs.Close()
				fs.Dispose()
			End If
		End Try
		Try
			file.Refresh()
			If file.Exists Then
				If originalFile.Exists Then originalFile.Delete()
				file.CopyTo(originalFile.FullName, True)
				file.Delete()
			Else
				result = False
			End If
		Catch ex As Exception
			result = False
		End Try
		Return result
	End Function

	Public Shared Function RecompressExistingTiff(ByVal filename As String, ByVal compression As EncoderValue) As Boolean
		Dim file As FileInfo = Nothing
		Try
			file = New FileInfo(filename)
		Catch ex As Exception
		End Try
		If file Is Nothing Then Return False
		file.CopyTo(file.FullName + ".tmp")
		Dim img As Image = Image.FromFile(file.FullName + ".tmp")
		'Dim files As New List(Of FileInfo)
		'files.Add(file)
		Dim result As Boolean = False
		Try
			result = ImageManager.WriteImageToDisk(img, file.FullName, True, compression)
			'result = ImageManager.UpdateImageOnDisk(filename, files, True, compression)
		Catch ex As Exception
			result = False
		End Try

		img.Dispose()
		img = Nothing

		If Not result Then
			file.Delete()
			file = New FileInfo(filename + ".tmp")
			file.CopyTo(filename)
			file.Delete()
		Else
			file = New FileInfo(filename + ".tmp")
			file.Delete()
		End If

		Return result
	End Function

	Private Shared Function GetUpToDateImage(ByVal originalImage As Image, ByVal changes As Dictionary(Of Integer, Image), ByVal index As Integer) As Image
		Dim pages As Integer = ImageManager.GetPageNumber(originalImage)
		If pages <= index Then
			Return Nothing
		End If

		If changes.ContainsKey(index) Then
			Return changes(index)
		End If

		Dim objGuid As Guid = originalImage.FrameDimensionsList(0)
		Dim objDimension As FrameDimension = New FrameDimension(objGuid)
		originalImage.SelectActiveFrame(objDimension, index)
		Return originalImage
	End Function

	Public Shared Function SplitImage(ByVal originalImage As String, ByVal path As String) As List(Of Guid)
		Dim img As Image = Nothing
		Dim dir As DirectoryInfo

		path = path.TrimEnd("\")
		path += "\"

		Try
			dir = New DirectoryInfo(path)
			img = Image.FromFile(originalImage)
		Catch ex As Exception
			Return Nothing
		End Try

		If Not dir.Exists Then
			dir.Create()
		End If

		Dim paths As New List(Of Guid)
		Dim pages As Integer = ImageManager.GetPageNumber(img)
		Dim guid As Guid = guid.Empty

		For i As Integer = 0 To pages - 1
			guid = System.Guid.NewGuid
			paths.Add(guid)
			ImageManager.GetPageImage(img, i).Save(dir.FullName + guid.ToString + ".tiff", ImageFormat.Tiff)
		Next

		img.Dispose()
		img = Nothing
		Return paths
	End Function

	Public Shared Sub SelectFrame(ByRef img As Image, ByVal index As Integer)
		If img IsNot Nothing Then
			Dim objGuid As Guid = img.FrameDimensionsList(0)
			Dim objDimension As FrameDimension = New FrameDimension(objGuid)
			Dim pages As Integer = img.GetFrameCount(objDimension)
			If index < pages AndAlso index >= 0 Then img.SelectActiveFrame(objDimension, index)
		End If
	End Sub

	''' <summary>
	''' Performs a custom rotation on the supplied image.
	''' </summary>
	''' <param name="img">The Image whose pages are to be rotated.</param>
	''' <param name="rotations">
	''' A dictionary of which pages and how to rotate using a 0-based index; Page 1 = index 0.
	''' </param>
	''' <returns>Returns the merged image after the rotations have been made.</returns>
	''' <remarks>
	''' The passed image will not be disposed of - the calling method must handle this!
	'''  *Try to avoid the pitfall of assigning the result of this method to the original Image container.
	'''   the original image will not be disposed of and may remain resident in memory for some time.
	''' </remarks>
	Public Shared Function CustomRotation(ByVal img As Image, ByVal rotations As Dictionary(Of Integer, RotateFlipType)) As Image
		'Check for invalid data.
		If rotations.Count <= 0 Then Return img
		If img Is Nothing Then Return Nothing

		'Split the original image into a list of the pages.
		Dim pages As List(Of Image) = ImageManager.SplitImage(img)

		'Cycle through each of the index/rotation pairs.
		For Each pair As KeyValuePair(Of Integer, RotateFlipType) In rotations
			'If the intended page is a valid target in our collection, proceed with the rotation.
			If pages.Count > pair.Key Then pages(pair.Key).RotateFlip(pair.Value)
		Next

		'Return the result of the Merge operation on the collection.
		Return ImageManager.MergeToImage(pages)
	End Function

	''' <summary>
	''' Rotates all odd pages 180° and returns a new image with the changes.
	''' </summary>
	''' <param name="img">The image to make changes to.</param>
	''' <returns>Returns an image whose odd pages have been rotated 180°</returns>
	''' <remarks>
	''' The passed image will not be disposed of - the calling method must handle this!
	'''  *Try to avoid the pitfall of assigning the result of this method to the original Image container.
	'''   the original image will not be disposed of and may remain resident in memory for some time.
	''' </remarks>
	Public Shared Function RotateOddPages(ByVal img As Image, ByVal fliptype As RotateFlipType) As Image
		'Check for nothing.
		If img Is Nothing Then Return Nothing

		'Create working variables.
		Dim rotations As New Dictionary(Of Integer, RotateFlipType)
		Dim count As Integer = ImageManager.GetPageNumber(img)

		'Generate the index/rotation dictionary
		For i As Integer = 0 To count - 1
			If i Mod 2 = 0 Then rotations.Add(i, fliptype)
		Next

		'Return the result of the custom rotation.
		Return ImageManager.CustomRotation(img, rotations)
	End Function

	Public Shared Function RotateOddPages(ByVal filename As String, ByVal fliptype As RotateFlipType) As Boolean
		Dim file As New FileInfo(filename)
		Dim img As Image = Nothing
		Dim imgResult As Image = Nothing
		Dim saveResult As Boolean = False
		If Not file.Exists Then Return False

		Try
			img = Image.FromFile(filename)
			imgResult = ImageManager.RotateOddPages(img, fliptype)
			img.Dispose()
			img = Nothing
			saveResult = WriteImageToDisk(imgResult, filename, True, EncoderValue.CompressionCCITT4)
		Catch ex As Exception
			Return False
		Finally
			If img IsNot Nothing Then
				img.Dispose()
				img = Nothing
			End If

			imgResult.Dispose()
			imgResult = Nothing
		End Try

		Return saveResult
	End Function

	''' <summary>
	''' Rotates all even pages 180° and returns a new image with the changes.
	''' </summary>
	''' <param name="img">The image to make changes to.</param>
	''' <returns>Returns an image whose even pages have been rotated 180°</returns>
	''' <remarks>
	''' The passed image will not be disposed of - the calling method must handle this!
	'''  *Try to avoid the pitfall of assigning the result of this method to the original Image container.
	'''   the original image will not be disposed of and may remain resident in memory for some time.
	''' </remarks>
	Public Shared Function RotateEvenPages(ByVal img As Image, ByVal fliptype As RotateFlipType) As Image
		'Check for nothing.
		If img Is Nothing Then Return Nothing

		'Create working variables.
		Dim rotations As New Dictionary(Of Integer, RotateFlipType)
		Dim count As Integer = ImageManager.GetPageNumber(img)

		'Generate the index/rotation dictionary
		For i As Integer = 1 To count - 1
			If i Mod 2 = 1 Then rotations.Add(i, fliptype)
		Next

		'Return the result of the custom rotation.
		Return ImageManager.CustomRotation(img, rotations)
	End Function

	Public Shared Function RotateEvenPages(ByVal filename As String, ByVal fliptype As RotateFlipType) As Boolean
		Dim file As New FileInfo(filename)
		Dim img As Image = Nothing
		Dim imgResult As Image = Nothing
		Dim saveResult As Boolean = False
		If Not file.Exists Then Return False

		Try
			img = Image.FromFile(filename)
			imgResult = ImageManager.RotateEvenPages(img, fliptype)
			img.Dispose()
			img = Nothing
			saveResult = WriteImageToDisk(imgResult, filename, True, EncoderValue.CompressionCCITT4)
		Catch ex As Exception
			saveResult = False
		Finally
			If img IsNot Nothing Then
				img.Dispose()
				img = Nothing
			End If

			imgResult.Dispose()
			imgResult = Nothing
		End Try

		Return saveResult
	End Function
#End Region

#Region " Public Methods "
	''' <summary>
	''' Resets the contents of the instance to default "empty" values.
	''' </summary>
	Public Sub Clear()
		mImage = Nothing
		mImages.Clear()
		mPages = 0
		mIsDirty = False
	End Sub

	''' <summary>
	''' Splits the image stored in the instance into an <see cref="ArrayList"/> of <see cref="MemoryStream"/>s
	''' </summary>
	''' <returns>Returns an ArrayList of MemoryStreams.
	''' Each MemoryStream in the array represents one page of the passed image.</returns>
	Public Function SplitImage() As List(Of Image)
		'Since an ArrayList of the original image is stored globally, is this function necessary as-is?
		' - Check to make sure it's not used internally as-is first, if so, consider making it Private
		If Not IsDirty Then
			mImagesInit()
		End If
		Return mImages
	End Function

	''' <summary>
	''' Extracts the specified page from the locally stored image.
	''' </summary>
	''' <param name="page">A positive <see cref="Integer"/> that indecates the page number to retrieve.</param>
	''' <returns>Returns the specified image as an <see cref="Image"/> object.</returns>
	''' <remarks>A 10 page image will have <c>[0-9]</c> as valid page numbers.
	''' Returns Nothing if an error occurs</remarks>
	Public Function GetPage(ByVal page As Integer) As Image
		If mImages.Count > 0 Then
			If page < mImages.Count Then
				Return mImages(page)
			End If
		Else
			Return ImageManager.GetPageImage(mImage, page)
		End If
		Return Nothing
	End Function

	''' <summary>
	''' Merges the image stored in the current instance into a <see cref="MemoryStream"/>.
	''' </summary>
	''' <returns>Returns a single, multiframe Image as a MemoryStream.</returns>
	''' <remarks></remarks>
	Public Function MergeImages() As MemoryStream
		Return ImageManager.MergeToStream(mImages)
		'If IsDirty Then
		'    Return ImageManager.MergeToStream(mImages)
		'Else
		'    Return ImageManager.ImageToMemoryStream(mImage)
		'End If
	End Function

	Public Sub DeleteImage(ByVal index As Integer)
		mImagesInit()
		ImageManager.DeletePageList(mImages, index)
		Me.OnContentsChanged()
	End Sub

	Public Sub InsertImage(ByVal index As Integer, ByVal page As Image)
		If page Is Nothing Then
			Exit Sub
		End If

		mImagesInit()
		ImageManager.InsertPageList(mImages, index, page)
		Me.OnContentsChanged()
	End Sub

	Public Sub InsertImage(ByVal index As Integer, ByVal ms As MemoryStream)
		Me.InsertImage(index, ImageManager.ObjectToImage(ms))
	End Sub

	'Public Sub InsertImage(ByVal index As Integer, ByVal filename As String)
	'    Dim file As New FileInfo(filename)
	'    If Not file.Exists Then
	'        Throw New ArgumentException("The specified file does not exist.", filename)
	'    End If


	'End Sub

	Public Sub ReplaceImage(ByVal index As Integer, ByVal page As Image)
		If page Is Nothing Then
			Exit Sub
		End If

		mImagesInit()
		ImageManager.ReplacePageList(mImages, index, page)
		Me.OnContentsChanged()
	End Sub

	Public Sub MoveImage(ByVal fromIndex As Integer, ByVal toIndex As Integer)
		If fromIndex = toIndex Then
			Exit Sub
		End If

		mImagesInit()

		Dim img As Image = Me.GetPage(fromIndex)
		ImageManager.MovePageList(mImages, fromIndex, toIndex)
		Me.OnContentsChanged()
	End Sub

	Public Sub Add(ByVal page As Image)
		If page Is Nothing Then
			Exit Sub
		End If

		mImagesInit()
		ImageManager.AppendPageList(mImages, page)
		Me.OnContentsChanged()
	End Sub

	''' <summary>
	''' Adds an Image as the last page of the current instance.
	''' </summary>
	''' <param name="obj">The Image object</param>
	''' <remarks>This method will accept Images as Image, Bitmap, MemoryStream, FileStream or String(As a filename).
	''' If an object is not a valid image, nothing will be added to the instance and an Argument Exception
	''' will be thrown.</remarks>
	Public Sub Add(ByVal obj As Object)
		Dim image As Image = ImageManager.ObjectToImage(obj)
		If Not image Is Nothing Then
			Me.Add(image)
		Else
			Throw New ArgumentException("The passed object is not a valid image.")
		End If
	End Sub

	Public Sub Save()
		If IsDirty Then
			If Not mPreserveReference Then
				Image = ImageManager.MergeToImage(mImages)
				mIsDirty = False
			Else
				OnImageSaved(ImageManager.MergeToImage(mImages))
			End If
		End If
	End Sub

	''' <summary>
	''' Saves the image to the specified file location.
	''' </summary>
	''' <remarks>
	''' This method will Overwrite an existing file.
	''' This method will not propagate changes into the Image property.
	''' </remarks>
	Public Function Save(ByVal filename As String) As Boolean
		Return Me.Save(filename, True, False)
	End Function

	Public Function Save(ByVal filename As String, ByVal overwrite As Boolean, ByVal merge As Boolean) As Boolean
		Try
			'Make sure we can overwrite an existing file, or that the specified directory exists.
			Dim file As New FileInfo(filename)
			If file.Exists AndAlso Not overwrite Then
				Exit Function
			ElseIf Not file.Directory.Exists Then
				Exit Function
			End If

			If merge And IsDirty Then
				Me.Save()
			End If

			Return ImageManager.WriteImageToDisk(Me.Image, filename, overwrite)
		Catch ex As Exception
			Return False
		End Try
	End Function
#End Region

#Region " Events "
	Public Event ImageSaved(ByVal image As Image)

	Private Sub OnImageSaved(ByVal image As Image)
		RaiseEvent ImageSaved(image)
	End Sub

	Public Event ContentsChanged(ByVal sender As Object)

	Private Sub OnContentsChanged()
		mIsDirty = True
		If Not mPreserveReference AndAlso mKeepCurrent Then
			Save()
		End If
		If mNotifyOnChange Then
			RaiseEvent ContentsChanged(Me)
		End If
	End Sub


	Public Shared Event ProgressUpdate(ByVal img As Image)

	Private Shared Sub OnProgressUpdate(ByVal img As Image)
		RaiseEvent ProgressUpdate(img)
	End Sub
#End Region

#Region " Private Methods "
	''' <summary>
	''' Initializes the internal <see cref="List(of Image)"/> collection.
	''' </summary>
	''' <remarks>If the collection already exists, nothing will take place.</remarks>
	Private Sub mImagesInit()
		If Me.mImages Is Nothing Then
			Me.mImages = New List(Of Image)
		End If

		If Me.mImages.Count = 0 Then
			Me.mImages = ImageManager.SplitImage(mImage)
		End If
	End Sub
#End Region

#Region " Constructors "
	Public Sub New()
		Me.Image = Nothing
	End Sub

	Public Sub New(ByVal img As Image)
		If Not IsNothing(img) Then
			Me.Image = img
		End If
	End Sub

	Public Sub New(ByVal images As List(Of Image))
		Me.New(ImageManager.MergeToImage(images))
	End Sub

	Public Sub New(ByVal images As List(Of MemoryStream))
		Me.New(ImageManager.MergeToImage(images))
	End Sub

	Public Sub New(ByVal filename As String)
		If Not filename = String.Empty Then
			Try
				Me.Image = System.Drawing.Image.FromFile(filename)
			Catch ex As Exception
				mImage = Nothing
			End Try
		Else
			mImage = Nothing
		End If
	End Sub

	Public Sub New(ByVal stream As Stream)
		Me.New(ImageManager.ObjectToImage(stream))
	End Sub

	Public Sub New(ByVal data As Byte())
		Me.New(New MemoryStream(data))
	End Sub

#End Region

#Region " Member Variables "
	Private mPages As Integer = 0
	Private mImage As System.Drawing.Image = Nothing
	Private mImages As New List(Of System.Drawing.Image)
	Private mIsDirty As Boolean = False

	Private mKeepCurrent As Boolean = False
	Private mPreserveReference As Boolean = False
	Private mNotifyOnChange As Boolean = False
#End Region

#Region " IDisposable Support "
	Private disposedValue As Boolean = False		' To detect redundant calls

	' IDisposable
	Protected Overridable Sub Dispose(ByVal disposing As Boolean)
		If Not Me.disposedValue Then
			If disposing Then
				' TODO: free unmanaged resources when explicitly called
			End If

			' TODO: free shared unmanaged resources
		End If
		Me.disposedValue = True
	End Sub

	' This code added by Visual Basic to correctly implement the disposable pattern.
	Public Sub Dispose() Implements IDisposable.Dispose
		' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
		Dispose(True)
		GC.SuppressFinalize(Me)
	End Sub
#End Region

End Class

#End Region

#Region " Old Version "

#Region " Revision 2 "
'**********************************************************************************************************
'Changes for this revision:
'  * Name Changed to ImageManager
'  * Delay building of the internal ArrayList until an operation to alter the original image takes place.
'      + This will increase performance by allowing a multipage image to be stored and read from without
'          going through the intensive operation of splitting the image.
'  * Optionally delay replacing the stored image with the merging of the ArrayList until the full image
'      is requested.  This feature is be controlled through the KeepCurrent property.
'      + Enabling this feature will eliminate redundant or unneccessary operations until needed, thereby
'          increasing performance.
'      + When enabled, the ArrayList is built and will be the sole target of most operations, including
'          retreival, deletion and replacement.
'  * New Properties:
'      + NotifyOnChange; If enabled, raises an event when the contents of the object have changed.
'      + KeepCurrent; When enabled, causes the mImage object to be updated on every OnContentsChanged call.
'      + PreserveReference; Makes the mImage object "Read-Only" internally to preserve object references.
'          - Overrides the KeepCurrent property.
'      + IsDirty; indicates whether or not changes have been saved.
'          - When PreserveReference = True, "saved" refers to current data having been requested.
'  * Enhanced Shared methods for better instance-less functionality.
'**********************************************************************************************************

'Imports System
'Imports System.IO
'Imports System.Drawing
'Imports System.Drawing.Imaging
'Imports System.Collections

'''' <summary>
'''' Manages single and multipage images.
'''' </summary>
'Public Class ImageManager

'#Region "Properties"
'    ''' <summary>
'    ''' Gets the number of pages in the current image.
'    ''' </summary>
'    Public ReadOnly Property Count() As Integer
'        Get
'            Return mPages
'        End Get
'    End Property

'    ''' <summary>
'    ''' Sets the ImageManager's current image.
'    ''' </summary>
'    ''' <value>A valid Image type (tiff, jpeg, bmp, etc)</value>
'    Public WriteOnly Property Picture() As Image
'        Set(ByVal value As Image)
'            mImage = Nothing
'            Try
'                mImage = value
'            Catch ex As Exception
'            End Try
'            GetPageNumber()
'            mImages = SplitImage()
'        End Set
'    End Property

'    ''' <summary>
'    '''
'    ''' Gets an arraylist of MemoryStreams.  Each stream represents one page from the stored Image.
'    ''' </summary>
'    Public ReadOnly Property ArrayList() As ArrayList
'        Get
'            Return mImages
'        End Get
'    End Property
'#End Region

'#Region "Public Methods"
'    ''' <summary>
'    ''' Resets the contents of the instance to default "empty" values.
'    ''' </summary>
'    Public Sub Clear()
'        mImage = Nothing
'        mPages = 0
'    End Sub

'    ''' <summary>
'    ''' Converts any supported image type into an <see cref="ArrayList"/> of <see cref="MemoryStream"/>s
'    ''' </summary>
'    ''' <param name="img">The image file you want to split.</param>
'    ''' <returns>Returns an ArrayList of MemoryStreams.
'    ''' Each MemoryStream in the array represents one page of the passed image.</returns>
'    ''' <remarks>Does not require an instance of the object.</remarks>
'    Public Shared Function SplitImage(ByVal img As Image) As ArrayList
'        Dim splitImages As ArrayList = New ArrayList
'        Try
'            Dim objGuid As Guid = img.FrameDimensionsList(0)
'            Dim objDimension As FrameDimension = New FrameDimension(objGuid)

'            Dim Pages As Integer = img.GetFrameCount(objDimension)

'            Dim enc As Encoder = Encoder.Compression
'            Dim int As Integer = 0

'            Dim i As Integer
'            For i = 0 To Pages - 1
'                img.SelectActiveFrame(objDimension, i)
'                Dim ep As EncoderParameters = New EncoderParameters(1)
'                ep.Param(0) = New EncoderParameter(enc, EncoderValue.CompressionNone)
'                Dim info As ImageCodecInfo = GetEncoderInfo("image/tiff")

'                Dim ms As MemoryStream = New MemoryStream
'                img.Save(ms, info, ep)
'                splitImages.Add(ms)
'            Next
'        Catch ex As Exception
'            Return Nothing
'        End Try
'        Return splitImages
'    End Function

'    ''' <summary>
'    ''' Splits the image stored in the instance into an <see cref="ArrayList"/> of <see cref="MemoryStream"/>s
'    ''' </summary>
'    ''' <returns>Returns an ArrayList of MemoryStreams.
'    ''' Each MemoryStream in the array represents one page of the passed image.</returns>
'    Public Function SplitImage() As ArrayList
'        'Since an ArrayList of the original image is stored globally, is this function necessary as-is?
'        ' - Check to make sure it's not used internally as-is first, if so, consider making it Private
'        Return SplitImage(mImage)
'    End Function

'    ''' <summary>
'    ''' Extracts the specified page from the locally stored image.
'    ''' </summary>
'    ''' <param name="page">A positive <see cref="Integer"/> that indecates the page number to retrieve.</param>
'    ''' <returns>Returns the specified image as an <see cref="Image"/> object.</returns>
'    ''' <remarks>A 10 page image will have <c>[0-9]</c> as valid page numbers.
'    ''' Returns Nothing if an error occurs</remarks>
'    Public Function GetPage(ByVal page As Integer) As Image
'        If page >= 0 AndAlso page < Count Then
'            Dim ms As MemoryStream = New MemoryStream
'            Try
'                'Dim objGuid As Guid = mImage.FrameDimensionsList(0)
'                'Dim objDimension As FrameDimension = New FrameDimension(objGuid)

'                'mImage.SelectActiveFrame(objDimension, page)
'                'mImage.Save(ms, ImageFormat.Bmp)

'                'Return Image.FromStream(ms)
'                Return Image.FromStream(mImages(page))
'            Catch ex As Exception
'                Return Nothing
'            End Try
'        Else
'            Return Nothing
'        End If
'    End Function

'    ''' <summary>
'    ''' Merges an <see cref="ArrayList"/> of <see cref="MemoryStream"/>s into a single MemoryStream using the <see cref="ImageFormat.Tiff"/> encoding.
'    ''' </summary>
'    ''' <param name="images">The <see cref="ArrayList"/> of <see cref="MemoryStream"/> encoded Images to merge.</param>
'    ''' <returns>Returns a single, multiframe Image as a MemoryStream.</returns>
'    ''' <remarks>Returns Nothing if an error occurs</remarks>
'    Public Shared Function MergeImages(ByVal images As ArrayList) As MemoryStream
'        If IsNothing(images) Then
'            Return Nothing
'        End If
'        If images.Count = 1 Then
'            Return images.Item(0)
'        End If

'        'Should this Try/Catch frame be removed?
'        Try
'            Dim ms As MemoryStream = New MemoryStream
'            Dim img As Image = Image.FromStream(images(0))
'            Dim img2 As Image
'            Dim i As Integer = 1

'            Dim myImageCodecInfo As ImageCodecInfo
'            Dim myEncoder As Encoder
'            Dim myEncoderParameter As EncoderParameter
'            Dim myEncoderParameters As EncoderParameters

'            myImageCodecInfo = GetEncoderInfo("image/tiff")
'            myEncoder = Encoder.SaveFlag
'            myEncoderParameters = New EncoderParameters(1)

'            ' Save the first page (frame).
'            myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.MultiFrame))
'            myEncoderParameters.Param(0) = myEncoderParameter
'            img.Save(ms, myImageCodecInfo, myEncoderParameters)

'            For i = 1 To images.Count - 1
'                img2 = Image.FromStream(images(i))
'                myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.FrameDimensionPage))
'                myEncoderParameters.Param(0) = myEncoderParameter
'                img.SaveAdd(img2, myEncoderParameters)
'            Next

'            myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.Flush))
'            myEncoderParameters.Param(0) = myEncoderParameter
'            img.SaveAdd(myEncoderParameters)

'            Return ms
'        Catch ex As Exception
'            Return Nothing
'        End Try
'    End Function

'    ''' <summary>
'    ''' Merges the image stored in the current instance into a <see cref="MemoryStream"/>.
'    ''' </summary>
'    ''' <returns>Returns a single, multiframe Image as a MemoryStream.</returns>
'    ''' <remarks></remarks>
'    Public Function MergeImages() As MemoryStream
'        Return MergeImages(mImages)
'    End Function

'    ''' <summary>
'    '''
'    ''' </summary>
'    ''' <param name="images"></param>
'    ''' <param name="location"></param>
'    ''' <returns></returns>
'    ''' <remarks></remarks>
'    Public Shared Function MergeImages(ByVal images As ArrayList, ByVal location As String) As Boolean
'        If IsNothing(images) Then
'            Return False
'        End If

'        Dim img As Image = Image.FromStream(images(0))
'        Dim img2 As Image
'        Dim i As Integer = 1

'        Dim myImageCodecInfo As ImageCodecInfo
'        Dim myEncoder As Encoder
'        Dim myEncoderParameter As EncoderParameter
'        Dim myEncoderParameters As EncoderParameters

'        myImageCodecInfo = GetEncoderInfo("image/tiff")
'        myEncoder = Encoder.SaveFlag
'        myEncoderParameters = New EncoderParameters(1)

'        ' Save the first page (frame).
'        myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.MultiFrame))
'        myEncoderParameters.Param(0) = myEncoderParameter
'        img.Save(location, myImageCodecInfo, myEncoderParameters)

'        myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.FrameDimensionPage))
'        myEncoderParameters.Param(0) = myEncoderParameter

'        For i = 1 To images.Count - 1
'            img2 = Image.FromStream(images(i))
'            img.SaveAdd(img2, myEncoderParameters)
'        Next

'        myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.Flush))
'        myEncoderParameters.Param(0) = myEncoderParameter
'        img.SaveAdd(myEncoderParameters)

'        Return True
'    End Function

'    Public Sub MergeImages(ByVal location As String)
'        MergeImages(mImages, location)
'    End Sub

'    Public Sub DeleteImage(ByVal index As Integer)
'        If index >= 0 AndAlso index < Count Then
'            Clear()
'            Try
'                mImages.RemoveAt(index)
'                Picture = Image.FromStream(MergeImages(mImages))
'            Catch ex As Exception
'            End Try
'        End If
'    End Sub

'    Public Sub InsertImage(ByVal index As Integer, ByVal img As Image)
'        If index >= 0 AndAlso index < Count AndAlso Not IsNothing(img) Then
'            Dim ms As MemoryStream = New MemoryStream
'            img.Save(ms, ImageFormat.Tiff)
'            Try
'                mImages.Insert(index, ms)
'                Picture = Image.FromStream(MergeImages(mImages))
'            Catch ex As Exception
'            End Try
'        ElseIf index = Count Then
'            Append(img)
'        End If
'    End Sub

'    Public Sub ReplaceImage(ByVal index As Integer, ByVal img As Image)
'        If index >= 0 AndAlso index < Count Then
'            DeleteImage(index)
'            InsertImage(index, img)
'        End If
'    End Sub

'    Public Sub MoveImage(ByVal fromIndex As Integer, ByVal toIndex As Integer)
'        If fromIndex >= 0 AndAlso toIndex >= 0 Then
'            If Count > fromIndex AndAlso Count > toIndex Then
'                Dim img As Image = Me.GetPage(fromIndex)
'                DeleteImage(fromIndex)
'                InsertImage(toIndex, img)
'            End If
'        End If
'    End Sub

'    Public Sub Append(ByVal img As Image)
'        If Not IsNothing(mImage) Then
'            Dim ms As MemoryStream = New MemoryStream
'            img.Save(ms, ImageFormat.Tiff)
'            mImages.Add(ms)
'            Picture = Image.FromStream(MergeImages(mImages))
'        Else
'            Picture = img
'        End If
'    End Sub
'#End Region

'#Region "Private Methods"
'    Private Sub GetPageNumber()
'        Dim objGuid As Guid = mImage.FrameDimensionsList(0)
'        Dim objDimension As FrameDimension = New FrameDimension(objGuid)

'        mPages = mImage.GetFrameCount(objDimension)
'    End Sub

'    Private Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
'        Dim encoders() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders
'        Dim i As Integer
'        For i = 0 To encoders.Length - 1
'            If encoders(i).MimeType = mimeType Then
'                Return encoders(i)
'            End If
'        Next
'        Return Nothing
'    End Function
'#End Region

'#Region "Constructors"
'    Public Sub New()
'        mImage = Nothing
'    End Sub

'    Public Sub New(ByVal img As Image)
'        If Not IsNothing(img) Then
'            'mImage = img
'            'GetPageNumber()
'            Picture = img
'        End If
'    End Sub

'    Public Sub New(ByVal images As ArrayList)
'        Try
'            If Not IsNothing(images) Then
'                mImages = images
'                mImage = Image.FromStream(MergeImages(images))
'                GetPageNumber()
'            End If
'        Catch ex As Exception
'            mImage = Nothing
'        End Try
'    End Sub

'    Public Sub New(ByVal value As String)
'        If Not value = "" Then
'            Try
'                Picture = Image.FromFile(value)
'            Catch ex As Exception
'                mImage = Nothing
'            End Try
'        Else
'            mImage = Nothing
'        End If
'    End Sub

'#End Region

'#Region "Member Variables"
'    'Private mImageFileName As String
'    Private mPages As Integer
'    Private mImage As Image
'    Private mImages As ArrayList
'#End Region
'End Class
#End Region

#Region " Revision 1 "
'Imports system
'Imports System.IO
'Imports System.Drawing
'Imports System.Drawing.Imaging
'Imports System.Collections


'Public Class ImageManager

'#Region "Properties"
'    ''' <summary>
'    ''' Returns the number of pages in the current Tiff image
'    ''' </summary>
'    Public ReadOnly Property Count()
'        Get
'            Return mPages
'        End Get
'    End Property

'    ''' <summary>
'    ''' Get or set the ImageManager's current image
'    ''' </summary>
'    ''' <value>A valid Image type</value>
'    ''' <returns></returns>
'    ''' <remarks></remarks>
'    Public Property Picture() As Image
'        Get
'            Return mImage
'        End Get
'        Set(ByVal value As Image)
'            mImage = New ImageManager(value).Picture
'        End Set
'    End Property
'#End Region

'#Region "Public Methods"
'    Public Function SplitImage() As ArrayList
'        Dim splitImages As ArrayList = New ArrayList
'        Try
'            Dim objGuid As Guid = mImage.FrameDimensionsList(0)
'            Dim objDimension As FrameDimension = New FrameDimension(objGuid)

'            Dim enc As Encoder = Encoder.Compression
'            Dim int As Integer = 0

'            Dim i As Integer
'            For i = 0 To mPages - 1
'                mImage.SelectActiveFrame(objDimension, i)
'                Dim ep As EncoderParameters = New EncoderParameters(1)
'                ep.Param(0) = New EncoderParameter(enc, EncoderValue.CompressionNone)
'                Dim info As ImageCodecInfo = GetEncoderInfo("image/tiff")

'                Dim ms As MemoryStream = New MemoryStream
'                mImage.Save(ms, info, ep)
'                splitImages.Add(ms)
'            Next
'        Catch ex As Exception
'            Return Nothing
'        End Try
'        Return splitImages
'    End Function

'    Public Shared Function SplitImage(ByVal img As Image) As ArrayList
'        Dim tm As ImageManager = New ImageManager(img)
'        Return tm.SplitImage
'    End Function

'    Public Function GetPage(ByVal page As Integer) As Image
'        Dim ms As MemoryStream = New MemoryStream
'        'Dim retImage As Image = Nothing
'        Try
'            Dim objGuid As Guid = mImage.FrameDimensionsList(0)
'            Dim objDimension As FrameDimension = New FrameDimension(objGuid)

'            mImage.SelectActiveFrame(objDimension, page)
'            mImage.Save(ms, ImageFormat.Bmp)

'            Return Image.FromStream(ms)
'        Catch ex As Exception
'            Return Nothing
'        End Try
'    End Function

'    Public Shared Function MergeImages(ByVal images As ArrayList) As Image
'        Dim ms As MemoryStream = New MemoryStream
'        Dim img As Image = Image.FromStream(images(0))
'        Dim img2 As Image
'        Dim i As Integer = 1

'        Dim myImageCodecInfo As ImageCodecInfo
'        Dim myEncoder As Encoder
'        Dim myEncoderParameter As EncoderParameter
'        Dim myEncoderParameters As EncoderParameters

'        myImageCodecInfo = GetEncoderInfo("image/tiff")
'        myEncoder = Encoder.SaveFlag
'        myEncoderParameters = New EncoderParameters(1)

'        ' Save the first page (frame).
'        myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.MultiFrame))
'        myEncoderParameters.Param(0) = myEncoderParameter
'        img.Save(ms, myImageCodecInfo, myEncoderParameters)

'        For i = 1 To images.Count - 1
'            img2 = Image.FromStream(images(i))
'            myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.FrameDimensionPage))
'            myEncoderParameters.Param(0) = myEncoderParameter
'            img.SaveAdd(img2, myEncoderParameters)
'        Next

'        myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.Flush))
'        myEncoderParameters.Param(0) = myEncoderParameter
'        img.SaveAdd(myEncoderParameters)

'        Return Image.FromStream(ms)
'    End Function

'    Public Function MergeImages() As Image
'        Return MergeImages(SplitImage(mImage))
'    End Function

'Public Shared Function MergeImages(ByVal images As ArrayList, ByVal location As String) As Boolean
'	'Dim ms As MemoryStream = New MemoryStream
'	Dim img As Image = Image.FromStream(images(0))
'	Dim img2 As Image
'	Dim i As Integer = 1

'	Dim myImageCodecInfo As ImageCodecInfo
'	Dim myEncoder As Encoder
'	Dim myEncoderParameter As EncoderParameter
'	Dim myEncoderParameters As EncoderParameters

'	myImageCodecInfo = GetEncoderInfo("image/tiff")
'	myEncoder = Encoder.SaveFlag
'	myEncoderParameters = New EncoderParameters(1)

'	' Save the first page (frame).
'	myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.MultiFrame))
'	myEncoderParameters.Param(0) = myEncoderParameter
'	img.Save("test123.tiff", myImageCodecInfo, myEncoderParameters)

'	myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.FrameDimensionPage))
'	myEncoderParameters.Param(0) = myEncoderParameter

'	For i = 1 To images.Count - 1
'		img2 = Image.FromStream(images(i))
'		img.SaveAdd(img2, myEncoderParameters)
'	Next

'	myEncoderParameter = New EncoderParameter(myEncoder, Fix(EncoderValue.Flush))
'	myEncoderParameters.Param(0) = myEncoderParameter
'	img.SaveAdd(myEncoderParameters)

'	Return True
'End Function

'    Public Sub MergeImages(ByVal location As String)
'        MergeImages(SplitImage(mImage), location)
'    End Sub

'    Public Sub DeleteImage(ByVal index As Integer)
'        Dim alImages As ArrayList = New ArrayList
'        alImages = SplitImage()
'        alImages.RemoveAt(index)
'        mImage = MergeImages(alImages)
'    End Sub

'    Public Sub InsertImage(ByVal index As Integer, ByVal img As Image)
'        Dim alImages As ArrayList = New ArrayList
'        alImages = SplitImage()
'        Dim ms As MemoryStream = New MemoryStream
'        img.Save(ms, ImageFormat.Bmp)
'        alImages.Insert(index, ms)
'        mImage = MergeImages(alImages)
'    End Sub

'    Public Sub ReplaceImage(ByVal index As Integer, ByVal img As Image)
'        DeleteImage(index)
'        InsertImage(index, img)
'    End Sub

'    Public Sub MoveImage(ByVal currIndex As Integer, ByVal newIndex As Integer)
'        Dim img As Image = Me.GetPage(currIndex)
'        DeleteImage(currIndex)
'        InsertImage(newIndex, img)
'    End Sub

'    Public Sub Append(ByVal img As Image)
'        Dim alImages As ArrayList = New ArrayList
'        alImages = SplitImage()
'        Dim ms As MemoryStream = New MemoryStream
'        img.Save(ms, ImageFormat.Bmp)
'        alImages.Add(ms)
'        mImage = MergeImages(alImages)
'    End Sub
'#End Region

'#Region "Private Methods"
'    Private Sub GetPageNumber()
'        Dim objGuid As Guid = mImage.FrameDimensionsList(0)
'        Dim objDimension As FrameDimension = New FrameDimension(objGuid)

'        mPages = mImage.GetFrameCount(objDimension)
'    End Sub

'    Private Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
'        Dim encoders() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders
'        Dim i As Integer
'        For i = 0 To encoders.Length - 1
'            If encoders(i).MimeType = mimeType Then
'                Return encoders(i)
'            End If
'        Next
'        Return Nothing
'    End Function
'#End Region

'#Region "Constructors"
'    Public Sub New()
'        mImage = Nothing
'    End Sub

'    Public Sub New(ByVal img As Image)
'        If Not IsNothing(img) Then
'            mImage = img
'            GetPageNumber()
'        End If
'    End Sub

'    Public Sub New(ByVal images As ArrayList)
'        Try
'            mImage = MergeImages(images)
'            GetPageNumber()
'        Catch ex As Exception
'            mImage = Nothing
'        End Try
'    End Sub

'    Public Sub New(ByVal value As String)
'        If Not value = "" Then
'            Try
'                mImage = Image.FromFile(value)
'                GetPageNumber()
'            Catch ex As Exception
'                mImage = Nothing
'            End Try
'        Else
'            mImage = Nothing
'        End If
'    End Sub

'#End Region

'#Region "Member Variables"
'    Private mImageFileName As String
'    Private mPages As Integer
'    Private mImage As Image
'#End Region
'End Class
#End Region

#Region " Original "
'http://www.codeproject.com/csharp/ImageManager.asp
'Note that this is coded in C#.NET

'using System;
'using System.IO;
'using System.Drawing;
'using System.Drawing.Imaging;
'using System.Collections;

'Namespace AMA.Util
'{
'	/// <summary>
'	/// Summary description for ImageManager.
'	/// </summary>
'	public class ImageManager : IDisposable
'	{
'		private string _ImageFileName;
'		private int _PageNumber;
'		private Image image;
'		private string _TempWorkingDir;

'		public ImageManager(string imageFileName)
'		{
'			this._ImageFileName=imageFileName;
'			image=Image.FromFile(_ImageFileName);
'			GetPageNumber();
'		}

'		public ImageManager(){
'		}

'		/// <summary>
'		/// Read the image file for the page number.
'		/// </summary>
'		private void GetPageNumber(){
'			Guid objGuid=image.FrameDimensionsList[0];
'			FrameDimension objDimension=new FrameDimension(objGuid);

'			//Gets the total number of frames in the .tiff file
'			_PageNumber=image.GetFrameCount(objDimension);

'			return;
'		}

'		/// <summary>
'		/// Read the image base string,
'		/// Assert(GetFileNameStartString(@"c:\test\abc.tif"),"abc")
'		/// </summary>
'		/// <param name="strFullName"></param>
'		/// <returns></returns>
'		private string GetFileNameStartString(string strFullName){
'			int posDot=_ImageFileName.LastIndexOf(".");
'			int posSlash=_ImageFileName.LastIndexOf(@"\");
'			return _ImageFileName.Substring(posSlash+1,posDot-posSlash-1);
'		}

'		/// <summary>
'		/// This function will output the image to a TIFF file with specific compression format
'		/// </summary>
'		/// <param name="outPutDirectory">The splited images' directory</param>
'		/// <param name="format">The codec for compressing</param>
'		/// <returns>splited file name array list</returns>
'		public ArrayList SplitTiffImage(string outPutDirectory,EncoderValue format)
'		{
'			string fileStartString=outPutDirectory+"\\"+GetFileNameStartString(_ImageFileName);
'			ArrayList splitedFileNames=new ArrayList();
'			try{
'				Guid objGuid=image.FrameDimensionsList[0];
'				FrameDimension objDimension=new FrameDimension(objGuid);

'				//Saves every frame as a separate file.
'				Encoder enc=Encoder.Compression;
'				int curFrame=0;
'				for (int i=0;i<_PageNumber;i++)
'				{
'					image.SelectActiveFrame(objDimension,curFrame);
'					EncoderParameters ep=new EncoderParameters(1);
'					ep.Param[0]=new EncoderParameter(enc,(long)format);
'					ImageCodecInfo info=GetEncoderInfo("image/tiff");

'					//Save the master bitmap
'					string fileName=string.Format("{0}{1}.TIF",fileStartString,i.ToString());
'					image.Save(fileName,info,ep);
'					splitedFileNames.Add(fileName);

'					curFrame++;
'				}
'			}catch (Exception){
'				throw;
'			}

'			return splitedFileNames;
'		}

'		/// <summary>
'		/// This function will join the TIFF file with a specific compression format
'		/// </summary>
'		/// <param name="imageFiles">string array with source image files</param>
'		/// <param name="outFile">target TIFF file to be produced</param>
'		/// <param name="compressEncoder">compression codec enum</param>
'		public void JoinTiffImages(string[] imageFiles,string outFile,EncoderValue compressEncoder)
'		{
'			try{
'				//If only one page in the collection, copy it directly to the target file.
'				if (imageFiles.Length==1)
'				{
'					File.Copy(imageFiles[0],outFile,true);
'					return;
'				}

'				//use the save encoder
'				Encoder enc=Encoder.SaveFlag;

'				EncoderParameters ep=new EncoderParameters(2);
'				ep.Param[0]=new EncoderParameter(enc,(long)EncoderValue.MultiFrame);
'				ep.Param[1] = new EncoderParameter(Encoder.Compression,(long)compressEncoder);

'				Bitmap pages=null;
'				int frame=0;
'				ImageCodecInfo info=GetEncoderInfo("image/tiff");


'				foreach(string strImageFile in imageFiles)
'				{
'					if(frame==0)
'					{
'						pages=(Bitmap)Image.FromFile(strImageFile);

'						//save the first frame
'						pages.Save(outFile,info,ep);
'					}
'					else
'					{
'						//save the intermediate frames
'						ep.Param[0]=new EncoderParameter(enc,(long)EncoderValue.FrameDimensionPage);

'						Bitmap bm=(Bitmap)Image.FromFile(strImageFile);
'						pages.SaveAdd(bm,ep);
'					}

'					if(frame==imageFiles.Length-1)
'					{
'						//flush and close.
'						ep.Param[0]=new EncoderParameter(enc,(long)EncoderValue.Flush);
'						pages.SaveAdd(ep);
'					}

'					frame++;
'				}
'			}catch (Exception){
'				throw;
'			}

'			return;
'		}

'		/// <summary>
'		/// This function will join the TIFF file with a specific compression format
'		/// </summary>
'		/// <param name="imageFiles">array list with source image files</param>
'		/// <param name="outFile">target TIFF file to be produced</param>
'		/// <param name="compressEncoder">compression codec enum</param>
'		public void JoinTiffImages(ArrayList imageFiles,string outFile,EncoderValue compressEncoder)
'		{
'			try
'			{
'				//If only one page in the collection, copy it directly to the target file.
'				if (imageFiles.Count==1){
'					File.Copy((string)imageFiles[0],outFile,true);
'					return;
'				}

'				//use the save encoder
'				Encoder enc=Encoder.SaveFlag;

'				EncoderParameters ep=new EncoderParameters(2);
'				ep.Param[0]=new EncoderParameter(enc,(long)EncoderValue.MultiFrame);
'				ep.Param[1] = new EncoderParameter(Encoder.Compression,(long)compressEncoder);

'				Bitmap pages=null;
'				int frame=0;
'				ImageCodecInfo info=GetEncoderInfo("image/tiff");


'				foreach(string strImageFile in imageFiles)
'				{
'					if(frame==0)
'					{
'						pages=(Bitmap)Image.FromFile(strImageFile);

'						//save the first frame
'						pages.Save(outFile,info,ep);
'					}
'					else
'					{
'						//save the intermediate frames
'						ep.Param[0]=new EncoderParameter(enc,(long)EncoderValue.FrameDimensionPage);

'						Bitmap bm=(Bitmap)Image.FromFile(strImageFile);
'						pages.SaveAdd(bm,ep);
'						bm.Dispose();
'					}

'					if(frame==imageFiles.Count-1)
'					{
'						//flush and close.
'						ep.Param[0]=new EncoderParameter(enc,(long)EncoderValue.Flush);
'						pages.SaveAdd(ep);
'					}

'					frame++;
'				}
'			}
'			catch (Exception ex)
'			{
'#If DEBUG Then
'				Console.WriteLine(ex.Message);
'#End If
'				throw;
'			}

'			return;
'		}

'		/// <summary>
'		/// Remove a specific page within the image object and save the result to an output image file.
'		/// </summary>
'		/// <param name="pageNumber">page number to be removed</param>
'		/// <param name="compressEncoder">compress encoder after operation</param>
'		/// <param name="strFileName">filename to be outputed</param>
'		/// <returns></</returns>
'		public void RemoveAPage(int pageNumber,EncoderValue compressEncoder,string strFileName){
'			try
'			{
'				//Split the image files to single pages.
'				ArrayList arrSplited=SplitTiffImage(this._TempWorkingDir,compressEncoder);

'				//Remove the specific page from the collection
'				string strPageRemove=string.Format("{0}\\{1}{2}.TIF",_TempWorkingDir,GetFileNameStartString(this._ImageFileName),pageNumber);
'				arrSplited.Remove(strPageRemove);

'				JoinTiffImages(arrSplited,strFileName,compressEncoder);
'			}
'			catch(Exception)
'			{
'				throw;
'			}

'			return;
'		}

'		/// <summary>
'		/// Getting the supported codec info.
'		/// </summary>
'		/// <param name="mimeType">description of mime type</param>
'		/// <returns>image codec info</returns>
'		private ImageCodecInfo GetEncoderInfo(string mimeType){
'			ImageCodecInfo[] encoders=ImageCodecInfo.GetImageEncoders();
'			for (int j=0;j<encoders.Length;j++){
'				if (encoders[j].MimeType==mimeType)
'					return encoders[j];
'			}

'			throw new Exception( mimeType + " mime type not found in ImageCodecInfo" );
'		}

'		/// <summary>
'		/// Return the memory steam of a specific page
'		/// </summary>
'		/// <param name="pageNumber">page number to be extracted</param>
'		/// <returns>image object</returns>
'		public Image GetSpecificPage(int pageNumber)
'		{
'			MemoryStream ms=null;
'			Image retImage=null;
'			try
'			{
'                ms=new MemoryStream();
'				Guid objGuid=image.FrameDimensionsList[0];
'				FrameDimension objDimension=new FrameDimension(objGuid);

'				image.SelectActiveFrame(objDimension,pageNumber);
'				image.Save(ms,ImageFormat.Bmp);

'				retImage=Image.FromStream(ms);

'				return retImage;
'			}
'			catch (Exception)
'			{
'				ms.Close();
'				retImage.Dispose();
'				throw;
'			}
'		}

'		/// <summary>
'		/// Convert the existing TIFF to a different codec format
'		/// </summary>
'		/// <param name="compressEncoder"></param>
'		/// <returns></returns>
'		public void ConvertTiffFormat(string strNewImageFileName,EncoderValue compressEncoder)
'		{
'			//Split the image files to single pages.
'			ArrayList arrSplited=SplitTiffImage(this._TempWorkingDir,compressEncoder);
'			JoinTiffImages(arrSplited,strNewImageFileName,compressEncoder);

'			return;
'		}

'		/// <summary>
'		/// Image file to operate
'		/// </summary>
'		public string ImageFileName
'		{
'			get
'			{
'				return _ImageFileName;
'			}
'			set{
'				_ImageFileName=value;
'			}
'		}

'		/// <summary>
'		/// Buffering directory
'		/// </summary>
'		public string TempWorkingDir
'		{
'			get
'			{
'				return _TempWorkingDir;
'			}
'			set{
'				_TempWorkingDir=value;
'			}
'		}

'		/// <summary>
'		/// Image page number
'		/// </summary>
'		public int PageNumber
'		{
'			get
'			{
'				return _PageNumber;
'			}
'		}


'		#region IDisposable Members

'		public void Dispose()
'		{
'			image.Dispose();
'			System.GC.SuppressFinalize(this);
'		}

'		#endregion
'	}
'}
#End Region

#End Region

#Region " Largess "
'Dim index As Integer = 0
'Dim enc As Encoder = Encoder.SaveFlag
'Dim ep As EncoderParameters = New EncoderParameters(2)
'Dim info As ImageCodecInfo = ImageManager.GetEncoderInfo("image/tiff")

'Dim compressor As Long = 0
'If imageToSave.PixelFormat = Imaging.PixelFormat.Format1bppIndexed Then
'	compressor = EncoderValue.CompressionLZW
'Else
'	compressor = EncoderValue.CompressionNone
'End If

'ep.Param(0) = New EncoderParameter(enc, EncoderValue.MultiFrame)
'ep.Param(1) = New EncoderParameter(Encoder.Compression, EncoderValue.CompressionNone)

'For index = 0 To pages - 1
'	Dim objGuid As Guid = imageToSave.FrameDimensionsList(0)
'	Dim objDimension As FrameDimension = New FrameDimension(objGuid)

'	imageToSave.SelectActiveFrame(objDimension, index)
'	imageToSave.Save(fs, imageToSave.RawFormat)



'Next

'				//use the save encoder
'				Encoder enc=Encoder.SaveFlag;

'				EncoderParameters ep=new EncoderParameters(2);
'				ep.Param[0]=new EncoderParameter(enc,(long)EncoderValue.MultiFrame);
'				ep.Param[1] = new EncoderParameter(Encoder.Compression,(long)compressEncoder);

'				Bitmap pages=null;
'				int frame=0;
'				ImageCodecInfo info=GetEncoderInfo("image/tiff");


'				foreach(string strImageFile in imageFiles)
'				{
'					if(frame==0)
'					{
'						pages=(Bitmap)Image.FromFile(strImageFile);

'						//save the first frame
'						pages.Save(outFile,info,ep);
'					}
'					else
'					{
'						//save the intermediate frames
'						ep.Param[0]=new EncoderParameter(enc,(long)EncoderValue.FrameDimensionPage);

'						Bitmap bm=(Bitmap)Image.FromFile(strImageFile);
'						pages.SaveAdd(bm,ep);
'					}

'					if(frame==imageFiles.Length-1)
'					{
'						//flush and close.
'						ep.Param[0]=new EncoderParameter(enc,(long)EncoderValue.Flush);
'						pages.SaveAdd(ep);
'					}

'					frame++;
'				}
#End Region