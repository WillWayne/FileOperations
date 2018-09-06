Imports System.Runtime.InteropServices
Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports FileLibrary.Tiff.Enums

Public Class ImageConverter

#Region " Public Methods "
	''' <summary>
	''' Converts an Image to a byte array.
	''' </summary>
	''' <param name="image">The Image to convert</param>
	''' <returns>Returns a byte array </returns>
	''' <remarks></remarks>
	Public Shared Function ImageToByteArray(ByVal image As System.Drawing.Image) As Byte()
		If image Is Nothing Then
			Throw New ArgumentNullException("image")
		End If

		Dim ret As Byte()

		'For some reason, saving a "MemoryBitmap" into a MemoryStream causes problems.
		'  Evidently, the format has no "Encoder" causing the Save operation to fail.
		'  To prevent this, if the specified image is a MemoryBitmap, save it as a regular Bitmap.
		Dim sourceFormat As ImageFormat = image.RawFormat
		If sourceFormat.Equals(ImageFormat.MemoryBmp) OrElse sourceFormat.Equals(ImageFormat.Bmp) Then sourceFormat = ImageFormat.Png

		Using ms As New System.IO.MemoryStream
			image.Save(ms, sourceFormat)
			ret = ms.ToArray()
			ms.Close()
		End Using

		Return ret
	End Function

	Public Shared Function ByteArrayToImage(ByVal bytes As Byte()) As System.Drawing.Image
		Dim newImage As System.Drawing.Image

		Dim ms As New MemoryStream(bytes, 0, bytes.Length)
		ms.Write(bytes, 0, bytes.Length)
		newImage = System.Drawing.Image.FromStream(ms, True)

		Return newImage
	End Function

	Public Shared Function GetScaledThumbnailImage(ByVal fileName As String, ByVal size As Size) As System.Drawing.Image
		Return GetScaledThumbnailImage(fileName, size.Width, size.Height)
	End Function

	Public Shared Function GetScaledThumbnailImage(ByVal fileName As String, ByVal width As Integer, ByVal height As Integer) As System.Drawing.Image

		Dim tmb As Image = Nothing

		If Not System.IO.File.Exists(fileName) Then Return Nothing

		Try
			If File.Exists(fileName) Then
				Using fs As New FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
					tmb = GetScaledThumbnailImage(System.Drawing.Image.FromStream(fs), width, height)
				End Using
			End If
		Catch ex As Exception
			'Since we don't have a support for other file types, extract the icon associated with the file.
			'tmb = CType(System.Drawing.Icon.ExtractAssociatedIcon(fileName).ToBitmap(), Image)     'This is a temporary solution for creating thumbnail from other file types.

            ' Resource dictionary missing
            'tmb = My.Resources.NoThumbnailAvailable
		End Try

		Return tmb
	End Function

	Public Shared Function GetScaledThumbnailImage(ByVal dataContent As Byte(), ByVal size As Size) As System.Drawing.Image
		Return GetScaledThumbnailImage(dataContent, size.Width, size.Height)
	End Function

	Public Shared Function GetScaledThumbnailImage(ByVal dataContent As Byte(), ByVal width As Integer, ByVal height As Integer) As System.Drawing.Image
		If dataContent Is Nothing Then
			Throw New NullReferenceException("Null reference at GetScaledThumbnailImage for dataContent")
		End If
		Dim img As Image = Nothing

		Try
			img = ByteArrayToImage(dataContent)
			If img IsNot Nothing Then
				img = GetScaledThumbnailImage(img, width, height)
			End If
		Catch ex As Exception
			'Since we don't have a support for other file types, extract the icon associated with the file.
			'TOD0: Decide what to do for extracting thumbnails from bytes.
		End Try
		Return img
	End Function

	Public Shared Function GetScaledThumbnailImage(ByVal img As System.Drawing.Image, ByVal size As Size) As System.Drawing.Image
		Return GetScaledImage(img, size)
	End Function

	Public Shared Function GetScaledThumbnailImage(ByVal img As System.Drawing.Image, ByVal width As Integer, ByVal height As Integer) As System.Drawing.Image
		Return GetScaledImage(img, New Size(width, height))
	End Function

	Public Shared Function GetScaledThumbnailImage(ByVal originalImage As System.Drawing.Image, ByVal percentage As Integer, Optional ByVal interpolationMode As Drawing2D.InterpolationMode = Drawing2D.InterpolationMode.Default) As System.Drawing.Image

		If percentage < 1 Then Throw New Exception("The thumbnail size must be at least 1% from its original size")
		Dim tmpBitmap As New Bitmap(CInt(originalImage.Width * 0.01F * percentage), CInt(originalImage.Height * 0.01F * percentage))

		Dim tmpGraphic As Graphics = Graphics.FromImage(tmpBitmap)
		If tmpGraphic Is Nothing Then
			Throw New NullReferenceException("tmpGraphic")
		End If

		tmpGraphic.InterpolationMode = interpolationMode
		tmpGraphic.DrawImage(originalImage, New Rectangle(0, 0, tmpBitmap.Width, tmpBitmap.Height), 0, 0, originalImage.Width, originalImage.Height, GraphicsUnit.Pixel)

		If tmpGraphic IsNot Nothing Then
			tmpGraphic.Dispose()
		End If
		Return CType(tmpBitmap, Image)
	End Function

	''' <summary>
	''' Generates a high-quality scaled version of the original image based on the dimensions
	''' passed in.  Maintains original aspect ratio to fit within the new size.
	''' </summary>
	Public Shared Function GetScaledImage(ByVal image As Image, ByVal newSize As Size) As Image
		If image Is Nothing Then Throw New ArgumentNullException("The specified image is Null")
		If newSize.IsEmpty Then Throw New ArgumentException("The specified size is Empty")
		Dim scaledSize As Size = ScaleSizeToFit(image.Size, newSize)

		'Indexed images can not have thumbnails made from them, so we will leverage the Bitmap
		' constructor taking an image and then cast as an Image to use the GetThumbnailImage method.
		If ImageIsIndexed(image) Then Return CType(New Bitmap(image), System.Drawing.Image).GetThumbnailImage(scaledSize.Width, scaledSize.Height, Nothing, New IntPtr)
		Return image.GetThumbnailImage(scaledSize.Width, scaledSize.Height, Nothing, New IntPtr)
	End Function

	Public Shared Function GetScaledImage(ByVal image As Image, ByVal newSize As Size, ByVal preScaled As Boolean) As Image
		'5 is the magic number
		'
		'The pre-scaled image must be a factor of 5 times the size of the expected thumbnail
		'  to guarantee a quality image.
		'The original image must be a factor of 5 times the size of the target pre-scaled
		'  image for pre-scaling to have a positive benefit on execution time.
		'Therefore, to take advantage of pre-scaling, the image must be 25 times larger
		'  in its largest dimension to deliver a discernable performance increase.
		If preScaled AndAlso newSize.Height * 25 < Math.Max(Image.Width, Image.Height) Then
			Dim intermediateSize As Size = New Size(newSize.Width * 5, newSize.Height * 5)
			Return GetScaledImage(CType(New Bitmap(Image, ScaleSizeToFit(Image.Size, intermediateSize)), Image), newSize)
		End If

		'If the optimization is not to be performed, use the original image.
		Return GetScaledImage(Image, newSize)
	End Function

	Public Shared Function ConvertToRGB(ByVal original As Bitmap) As Bitmap
		Dim newImage As Bitmap = New Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb)
		newImage.SetResolution(original.HorizontalResolution, original.VerticalResolution)
		Dim g As Graphics = Graphics.FromImage(newImage)
		g.DrawImageUnscaled(original, 0, 0)
		g.Dispose()
		Return newImage
	End Function

	'Public Shared Function ConvertToOtherPixelFormat(ByVal original As Bitmap, ByVal newFormat As PixelFormat) As Bitmap
	'	If original.PixelFormat = newFormat Then Return original

	'	Dim newImage As Bitmap = New Bitmap(original.Width, original.Height, newFormat)
	'	newImage.SetResolution(original.HorizontalResolution, original.VerticalResolution)
	'	Dim g As Graphics = Graphics.FromImage(newImage)
	'	g.DrawImageUnscaled(original, 0, 0)
	'	g.Dispose()
	'	Return newImage
	'End Function

	''' <summary>
	''' Converts an image to 8 bits per pixel indexed; Grayscale.
	''' </summary>
	''' <param name="original">The Bitmap to convert</param>
	''' <returns>Returns a copy of the original image as an 8bpp grayscale image.</returns>
	''' <remarks>
	''' This image will only save to a stream as 24bpp if the JPEG format is selected.  (Maybe 16 bpp)
	''' Formats that support 8bpp indexed images include BMP, GIF and PNG, but file size savings is
	''' absent or infintesimal compared to saving the grayscale image as a 24bpp jpeg.
	''' The resulting 24bpp grayscale jpeg usually see some file size savings over the original, full-color jpeg.
	'''
	''' </remarks>
	Public Shared Function ConvertToGrayscale(ByVal original As Bitmap) As Bitmap
		Return ConvertImageTo1or8bpp(original, PixelFormat.Format8bppIndexed)
	End Function

	''' <summary>
	''' Converts an image to 1 bit per pixel indexed format.
	''' </summary>
	''' <param name="original">The Bitmap to convert</param>
	''' <returns>Returns a copy of the original image as an 1bpp bitonal image.</returns>
	''' <remarks>
	''' For best results, the result of this function should be saved as a TIFF using the
	''' <see cref="EncoderValue.CompressionCCITT4"/> compression.
	'''
	''' If converting a multi-frame image, the resulting Bitmap will be saved as a TIFF image
	''' using the the <see cref="EncoderValue.CompressionCCITT4"/> compression.
	''' </remarks>
	Public Shared Function ConvertToBitonal(ByVal original As Bitmap) As Bitmap
		Return ConvertImageTo1or8bpp(original, PixelFormat.Format1bppIndexed)
	End Function

	''' <summary>
	''' Inverts the colors of a Bitmap
	''' </summary>
	''' <param name="original">The Bitmap to Invert</param>
	''' <returns>Returns a new Bitmap identical to the original, but for an inversion of the colors</returns>
	Public Shared Function InvertBitmap(ByVal original As Bitmap) As Bitmap
		'Create a temporary bitmap for the BitBlt operation
		Using temp As Bitmap = New Bitmap(original.Width, original.Height, original.PixelFormat)
			Return PerformBitBlt(temp, original, TernaryRasterOperations.DSTINVERT)
		End Using
	End Function

	''' <summary>
	''' Replaces the specified color in a given bitmap with a new color (user defined)
	''' </summary>
	''' <param name="original">Bitmap to modify</param>
	''' <param name="targetColor">The Color to replace</param>
	''' <param name="replacementColor">The Color to introduce</param>
	''' <returns>Returns a new Bitmap as a result of the operation</returns>
	'''
	Public Shared Function ReplaceColor(ByRef original As Bitmap, ByVal targetColor As Color, ByVal replacementColor As Color) As Bitmap
		Return gdiColorChange(original, targetColor, replacementColor, True)
	End Function

	''' <summary>
	''' Replaces all colors in the Bitmap with the replacement color Except for the ignored color.
	''' </summary>
	''' <param name="original">Bitmap to modify</param>
	''' <param name="ignoredColor">The Color to not replace</param>
	''' <param name="replacementColor">The Color to replace other colors with</param>
	''' <returns>Returns a new Bitmap as a result of the operation</returns>
	''' <remarks>
	''' This method specifically designed to enhance the contrast between text and
	''' background colors on terminal screenshots.
	''' </remarks>
	Public Shared Function NotReplaceColor(ByVal original As Bitmap, ByVal ignoredColor As Color, ByVal replacementColor As Color) As Bitmap
		Return gdiColorChange(original, ignoredColor, replacementColor, False)
	End Function
#End Region

#Region " Private Methods "
	Private Shared Function CreateMask(ByVal original As Bitmap, ByVal backColor As Color) As Bitmap
		Dim w As Integer = original.Width
		Dim h As Integer = original.Height
		Dim target As Bitmap = New Bitmap(original.Size.Width, original.Size.Height, original.PixelFormat)
		'
		' Step (3): use GDI's BitBlt function to copy from original hbitmap into monocrhome bitmap
		' GDI programming is kind of confusing... nb. The GDI equivalent of "Graphics" is called a "DC".
		Dim pTargetDC As IntPtr = CreateCompatibleDC(GetDC(0)) 'target.GetHbitmap(Color.White)
		Dim pTargetBmp As IntPtr = target.GetHbitmap()
		SelectObject(pTargetDC, pTargetBmp)

		Dim screenDC As IntPtr = GetDC(IntPtr.Zero)
		Dim sourceDC As IntPtr = CreateCompatibleDC(original.GetHbitmap)
		SelectObject(sourceDC, original.GetHbitmap)

		Dim maskDC As IntPtr = CreateCompatibleDC(screenDC)
		Dim maskBmp As IntPtr = CreateCompatibleBitmap(maskDC, w, h)
		SelectObject(maskDC, maskBmp)

		SetBkColor(sourceDC, backColor.ToArgb() And &HFFFFFF)
		BitBlt(maskDC, 0, 0, w, h, sourceDC, 0, 0, TernaryRasterOperations.SRCCOPY)
		BitBlt(pTargetDC, 0, 0, w, h, sourceDC, 0, 0, TernaryRasterOperations.SRCINVERT)
		BitBlt(pTargetDC, 0, 0, w, h, maskDC, 0, 0, TernaryRasterOperations.SRCAND)

		Dim result As Bitmap = Bitmap.FromHbitmap(pTargetBmp)

		DeleteDC(sourceDC)
		ReleaseDC(IntPtr.Zero, screenDC)
		DeleteDC(sourceDC)
		DeleteDC(pTargetDC)
		DeleteObject(pTargetBmp)
		DeleteDC(maskDC)
		DeleteObject(maskBmp)

		Return result
	End Function

	Private Shared Function BlitReplaceColor(ByVal original As Bitmap, ByVal cOldColor As Color, ByVal cNewColor As Color) As Bitmap
		Dim w As Integer = original.Width
		Dim h As Integer = original.Height

		'Dim screenDC As IntPtr = GetDC(IntPtr.Zero)
		'Dim target As New Bitmap(w, h, original.PixelFormat)

		'Dim pTargetDC As IntPtr = CreateCompatibleDC(screenDC)
		'Dim pTargetBmp As IntPtr = target.GetHbitmap()
		'SelectObject(pTargetDC, pTargetBmp)

		Dim pSourceDC As IntPtr = CreateCompatibleDC(original.GetHbitmap)
		Dim pSourceBmp As IntPtr = original.GetHbitmap()
		SelectObject(pSourceDC, pSourceBmp)

		Dim pMaskDc As IntPtr '= CreateCompatibleDC(pTargetDC)
		Dim pMaskBmp As IntPtr '= original.GetHbitmap()

		'If Not CreateMask(original, cOldColor) Then
		'	DeleteDC(pSourceDC)
		'	DeleteDC(pMaskDc)
		'	DeleteObject(pMaskBmp)
		'	DeleteObject(pSourceBmp)
		'	Return Nothing
		'End If
		'SelectObject(pMaskDc, pMaskBmp)

		BitBlt(pMaskDc, 0, 0, w, h, pSourceDC, 0, 0, TernaryRasterOperations.SRCCOPY)

		Dim result As Bitmap = Bitmap.FromHbitmap(pMaskBmp)

		'SetBkColor(pSourceDC, cOldColor.ToArgb() And &HFFFFFF)
		'Dim sb As New SolidBrush(cNewColor)
		'Dim g As Graphics = Graphics.FromHdc(pTargetDC)
		'g.FillRectangle(sb, New Rectangle(0, 0, w, h))
		'sb.Dispose()

		'BitBlt(pMaskDc, 0, 0, w, h, pSourceDC, 0, 0, TernaryRasterOperations.SRCCOPY)
		'BitBlt(pTargetDC, 0, 0, w, h, pSourceDC, 0, 0, TernaryRasterOperations.SRCINVERT)
		'BitBlt(pTargetDC, 0, 0, w, h, pMaskDc, 0, 0, TernaryRasterOperations.SRCAND)
		'BitBlt(pTargetDC, 0, 0, w, h, pSourceDC, 0, 0, TernaryRasterOperations.SRCINVERT)

		'Dim result As Bitmap = Bitmap.FromHbitmap(pTargetBmp)

		DeleteDC(pSourceDC)
		DeleteDC(pMaskDc)
		DeleteObject(pMaskBmp)
		DeleteObject(pSourceBmp)

		Return result
	End Function


	''' <summary>
	''' Makes a change to an image by replacing colors
	''' </summary>
	''' <param name="original">The image to make color changes to</param>
	''' <param name="targetColor">The color used as a condition</param>
	''' <param name="replacementColor">The color to write into the Bitmap as a result of passing the condition.</param>
	''' <param name="ReplaceTargetColor">
	''' True to replace found targetColors with the replacementColor.
	''' False to replace all colors except the targetColor with the replacementColor.
	''' </param>
	''' <returns>Returns a new 32bpp bitmap that reflects the requested changes.</returns>
	Private Shared Function ColorChange(ByVal original As Bitmap, ByVal targetColor As Color, ByVal replacementColor As Color, ByVal ReplaceTargetColor As Boolean) As Bitmap
		Dim w As Integer = original.Width
		Dim h As Integer = original.Height
		Dim result As Bitmap = New Bitmap(original)
		Dim fromColor As Integer = targetColor.ToArgb()	' And &HFFFFFF
		Dim toColor As Integer = replacementColor.ToArgb() ' And &HFFFFFF

		Dim x As Integer
		Dim y As Integer
		Dim PixelSize As Integer = IIf(GetImageBitDepth(result) >= 8, GetImageBitDepth(result) / 8, 0)
		If PixelSize <> 4 Then Return Nothing 'Not ready to deal with sub-byte per pixel formats

		Dim sourceData As BitmapData = result.LockBits(New Rectangle(0, 0, w, h), _
		System.Drawing.Imaging.ImageLockMode.ReadWrite, result.PixelFormat)

		For y = 0 To sourceData.Height - 1
			For x = 0 To sourceData.Width - 1
				If (Marshal.ReadInt32(sourceData.Scan0, (sourceData.Stride * y) + (4 * x)) = fromColor) Xor (Not ReplaceTargetColor) Then
					Marshal.WriteInt32(sourceData.Scan0, (sourceData.Stride * y) + (4 * x), toColor)
				End If
			Next
		Next

		result.UnlockBits(sourceData)

		Return result
	End Function

	Private Shared Function gdiColorChange(ByVal original As Bitmap, ByVal targetColor As Color, ByVal replacementColor As Color, ByVal ReplaceTargetColor As Boolean) As Bitmap
		'Grab these for convenience
		Dim w As Integer = original.Width
		Dim h As Integer = original.Height

		'Short-circuit some easy solutions
		If targetColor.Equals(replacementColor) Then
			If ReplaceTargetColor Then
				'We're not making any changes, but go ahead and clone since we expect a new image reference.
				Return original.Clone
			Else
				'The user has just asked for an image of size w*h, purely of color replacementColor
				Dim ret As New Bitmap(w, h, original.PixelFormat)
				Dim g As Graphics = Graphics.FromImage(ret)
				Dim b As New SolidBrush(replacementColor)
				g.FillRectangle(b, New Rectangle(0, 0, w, h))
				g.Dispose()
				Return ret
			End If
		End If

		Dim targetBmp As Bitmap = New Bitmap(w, h, original.PixelFormat)

		Dim pTargetDC As IntPtr = CreateCompatibleDC(GetDC(0)) 'target.GetHbitmap(Color.White)
		Dim pTargetBmp As IntPtr = targetBmp.GetHbitmap()
		SelectObject(pTargetDC, pTargetBmp)

		Dim pSourceDC As IntPtr = CreateCompatibleDC(IntPtr.Zero)
		Dim pSourceBmp As IntPtr = original.GetHbitmap()
		SelectObject(pSourceDC, pSourceBmp)

		BitBlt(pTargetDC, 0, 0, w, h, pSourceDC, 0, 0, TernaryRasterOperations.SRCCOPY)

		Dim x As Integer
		Dim y As Integer

		'The high-order byte for the GetPixel and SetPixel operations must be '00'
		'Because the high-order byte for the ToArgb method is often 'FF', AND these
		' results with 0x00FFFFFF to change all the high-order bits to 0s for compatibility.
		Dim target As Int32 = targetColor.ToArgb() And &HFFFFFF
		Dim replacement As Int32 = ReverseRGB(replacementColor)	'.ToArgb() And &HFFFFFF

		'BUG: When writing this replacement color value into SetPixelV, the result is not
		'		translated into the correct color.  This could be an issue with conversion to 24bpp
		'		formats, or setting the pixel value itself.
		'UPDATE: As it turns out, the byte ordering seems to be reversing as the color is assigned to the pixel.
		'		Correction of this phenomenon should resolve the color issues.

		For y = 0 To h - 1
			For x = 0 To w - 1
				'The condition is met if the targetColor is found and ReplaceTargetColor is True
				' and when the targetColor is not found and ReplaceTargetColor is False
				If (GetPixel(pTargetDC, x, y) = target) Xor (Not ReplaceTargetColor) Then
					SetPixelV(pTargetDC, x, y, replacement)	'Set the new pixel color
				End If
			Next
		Next

		'Create the resulting Bitmap from the HBitmap.
		Dim result As Bitmap = Bitmap.FromHbitmap(pTargetBmp)

		'Cleanup the pointers
		DeleteDC(pSourceDC)
		DeleteDC(pTargetDC)
		DeleteObject(pTargetBmp)
		DeleteObject(pSourceBmp)

		Return result
	End Function

	Private Shared Function GetImageBitDepth(ByRef bmp As Bitmap) As Integer
		If bmp Is Nothing Then Throw New ArgumentNullException("The passed Bitmap is Null.")

		Dim pixelF As String = bmp.PixelFormat.ToString()
		If pixelF.StartsWith("Format1bpp") Then Return 1
		If pixelF.StartsWith("Format4bpp") Then Return 4
		If pixelF.StartsWith("Format8bpp") Then Return 8
		If pixelF.StartsWith("Format16bpp") Then Return 16
		If pixelF.StartsWith("Format24bpp") Then Return 24
		If pixelF.StartsWith("Format32bpp") Then Return 32
		If pixelF.StartsWith("Format48bpp") Then Return 48
		If pixelF.StartsWith("Format64bpp") Then Return 64

		Throw New ArgumentException("The PixelFormat of the Bitmap was of an unexpected type.")
	End Function

	Private Shared Function ConvertImageTo1or8bpp(ByVal original As Bitmap, ByVal pixelformat As PixelFormat) As Bitmap
		Dim pages As Integer = ImageHelper.GetPageCount(original)
		If pages = 1 Then Return ImageConverter.ConvertPageTo1or8bpp(original, pixelformat.Format1bppIndexed)

		'If we get here, we are converting a multi-frame image.
		Dim results() As Bitmap = New Bitmap(pages) {}
		Dim result As Bitmap = Nothing

		For i As Integer = 0 To pages - 1
			ImageHelper.SelectFrame(original, i)
			results(i) = ImageConverter.ConvertPageTo1or8bpp(original, pixelformat)
		Next

		Dim ms As New MemoryStream
		If pixelformat = Imaging.PixelFormat.Format1bppIndexed Then
			ImageHelper.WriteBitmapsToStream(results, ms, CompressionType.CCITT4, ImageFormat.Tiff)
		Else
			ImageHelper.WriteBitmapsToStream(results, ms, CompressionType.LZW, ImageFormat.Tiff)
		End If
		result = Bitmap.FromStream(ms, True, True)
		Return result
	End Function

	''' <summary>
	''' Copies a bitmap into a 1bpp/8bpp bitmap of the same dimensions, fast
	''' </summary>
	''' <param name="bmp">original bitmap</param>
	''' <param name="pixelFormat">The Pixel format to convert to.</param>
	''' <returns>
	''' Returns a copy of the Original Bitmap, converted to the new PixelFormat.
	''' </returns>
	Private Shared Function ConvertPageTo1or8bpp(ByVal bmp As System.Drawing.Bitmap, ByVal pixelFormat As PixelFormat) As System.Drawing.Bitmap
		Dim bpp As Integer = 0

		Select Case pixelFormat
			Case Imaging.PixelFormat.Format1bppIndexed
				bpp = 1
			Case Imaging.PixelFormat.Format8bppIndexed
				bpp = 8
			Case Else
				Throw New NotSupportedException("The specified Pixel Format is not supported by this method.")
		End Select

		' Plan: built into Windows GDI is the ability to convert
		' bitmaps from one format to another. Most of the time, this
		' job is actually done by the graphics hardware accelerator card
		' and so is extremely fast. The rest of the time, the job is done by
		' very fast native code.
		' We will call into this GDI functionality from VB. Our plan:
		' (1) Convert our Bitmap into a GDI hbitmap (ie. copy managed->unmanaged)
		' (2) Create a GDI monochrome hbitmap
		' (3) Use GDI "BitBlt" function to copy from hbitmap into monochrome (as above)
		' (4) Convert the monochrone hbitmap into a Bitmap (ie. copy unmanaged->managed)

		Dim w As Integer = bmp.Width
		Dim h As Integer = bmp.Height
		Dim sourceBmp As IntPtr = bmp.GetHbitmap()
		' this is step (1)
		'
		' Step (2): create the monochrome bitmap.
		' "BITMAPINFO" is an interop-struct which we define below.
		' In GDI terms, it's a BITMAPHEADERINFO followed by an array of two RGBQUADs
		Dim bmi As New BitmapInfoHeader()
		bmi.Size = 40
		' the size of the BITMAPHEADERINFO struct
		bmi.Width = w
		bmi.Height = h
		bmi.Planes = 1
		' "planes" are confusing. We always use just 1. Read MSDN for more info.
		bmi.BitCount = CShort(bpp)
		' ie. 1bpp or 8bpp
		bmi.Compression = BI_RGB
		' ie. the pixels in our RGBQUAD table are stored as RGBs, not palette indexes
		bmi.SizeImage = CInt((((w + 7) And 4294967288) * h / 8))
		bmi.XPelsPerMeter = 1000000
		' not really important
		bmi.YPelsPerMeter = 1000000
		' not really important
		' Now for the colour table.
		Dim ncols As UInteger = CInt(1) << bpp
		' 2 colours for 1bpp; 256 colours for 8bpp
		bmi.ClrUsed = ncols
		bmi.ClrImportant = ncols
		bmi.cols = New UInteger(255) {}
		' The structure always has fixed size 256, even if we end up using fewer colours
		If bpp = 1 Then
			bmi.cols(0) = MakeRGB(0, 0, 0)
			bmi.cols(1) = MakeRGB(255, 255, 255)
		Else
			For i As Integer = 0 To ncols - 1
				bmi.cols(i) = MakeRGB(i, i, i)
			Next
		End If
		' For 8bpp we've created an palette with just greyscale colours.
		' You can set up any palette you want here. Here are some possibilities:
		' greyscale: for (int i=0; i<256; i++) bmi.cols[i]=MAKERGB(i,i,i);
		' rainbow: bmi.biClrUsed=216; bmi.biClrImportant=216; int[] colv=new int[6]{0,51,102,153,204,255};
		' for (int i=0; i<216; i++) bmi.cols[i]=MAKERGB(colv[i/36],colv[(i/6)%6],colv[i%6]);
		' optimal: a difficult topic: http://en.wikipedia.org/wiki/Color_quantization
		'
		' Now create the indexed bitmap "hbm0"
		Dim bits0 As IntPtr
		' not used for our purposes. It returns a pointer to the raw bits that make up the bitmap.
		Dim targetBmp As IntPtr = CreateDIBSection(IntPtr.Zero, bmi, DIB_RGB_COLORS, bits0, IntPtr.Zero, 0)
		'
		' Step (3): use GDI's BitBlt function to copy from original hbitmap into monocrhome bitmap
		' GDI programming is kind of confusing... nb. The GDI equivalent of "Graphics" is called a "DC".
		Dim screenDC As IntPtr = GetDC(IntPtr.Zero)
		Dim sourceDC As IntPtr = CreateCompatibleDC(screenDC)
		SelectObject(sourceDC, sourceBmp)
		Dim targetDC As IntPtr = CreateCompatibleDC(screenDC)
		SelectObject(targetDC, targetBmp)

		BitBlt(targetDC, 0, 0, w, h, sourceDC, 0, 0, TernaryRasterOperations.SRCCOPY)
		Dim result As System.Drawing.Bitmap = System.Drawing.Bitmap.FromHbitmap(targetBmp)

		DeleteDC(sourceDC)
		DeleteDC(targetDC)
		ReleaseDC(IntPtr.Zero, screenDC)
		DeleteObject(sourceBmp)
		DeleteObject(targetBmp)

		Return result
	End Function

	Private Shared Function PerformBitBlt(ByVal original As Bitmap, ByRef target As Bitmap, ByVal command As TernaryRasterOperations) As Bitmap
		'Dim target As Bitmap = New Bitmap(original.Size.Width, original.Size.Height, original.PixelFormat)

		Dim pTargetDC As IntPtr = CreateCompatibleDC(GetDC(0)) 'target.GetHbitmap(Color.White)
		Dim pTargetBmp As IntPtr = target.GetHbitmap()
		SelectObject(pTargetDC, pTargetBmp)

		Dim pSourceDC As IntPtr = CreateCompatibleDC(pTargetDC)
		Dim pSourceBmp As IntPtr = original.GetHbitmap()
		SelectObject(pSourceDC, pSourceBmp)

		BitBlt(pTargetDC, 0, 0, original.Width, original.Height, pSourceDC, 0, 0, command)
		Dim result As Bitmap = Bitmap.FromHbitmap(pTargetBmp)

		DeleteDC(pSourceDC)
		DeleteDC(pTargetDC)
		DeleteObject(pTargetBmp)
		DeleteObject(pSourceBmp)
		Return result
	End Function

	Private Shared Function MakeRGB(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As UInteger
		Return CInt((b And 255)) Or CInt(((r And 255) << 8)) Or CInt(((g And 255) << 16))
	End Function

	''' <summary>
	''' Reverses the RGB component value of a specified Color object.
	''' </summary>
	''' <param name="color">The color whose RGB components are to be reversed</param>
	''' <returns>
	''' Returns an unsigned integer whose high and low-order bytes (excluding the highest)
	''' have been swapped.
	''' </returns>
	''' <remarks>
	''' This method was created specifically for compatibility with the methods employed by the
	''' gdiColorChange method.
	'''
	''' Essentially, the Red and Blue values are swapped.
	''' </remarks>
	Private Shared Function ReverseRGB(ByVal color As Color) As UInteger
		Dim cAlpha As Integer = color.ToArgb()
		Dim cRed As Integer
		Dim cGreen As Integer
		Dim cBlue As Integer

		cBlue = cAlpha Mod 256
		cAlpha /= 256
		cGreen = cAlpha Mod 256
		cAlpha /= 256
		cRed = cAlpha Mod 256
		cAlpha /= 256

		Return CInt((cRed And 255)) Or CInt(((cGreen And 255) << 8)) Or CInt(((cBlue And 255) << 16) Or CInt(((cAlpha And 255) << 32)))
		'Return MakeRGB(cRed, cGreen, cBlue)
	End Function

	Private Shared Function MakeIndexedBitmapinfo() As BitmapInfoHeader
		Dim notImplemented As Boolean = True
		If notImplemented Then
			Throw New NotImplementedException("This method is not yet Implemented.")
		Else
			'Create the BMI
			Dim result As New BitmapInfoHeader

			Return result
		End If
	End Function

	Private Shared Function MakeNonIndexedBitmapinfo() As BitmapInfoHeader
		Dim notImplemented As Boolean = True
		If notImplemented Then
			Throw New NotImplementedException("This method is not yet Implemented.")
		Else
			'Create the BMI
			Dim result As New BitmapInfoHeader

			Return result
		End If
	End Function

	''' <summary>
	''' Scales the given Size to fit snugly within the target size, etiher up or down.
	''' </summary>
	''' <param name="currentSize">The size of the object to scale.</param>
	''' <param name="targetBounds">
	''' The new constraints within which to scale the <paramref name="currentSize"/> up or down.
	''' </param>
	''' <returns>
	''' Returns a new size that fits within the constraints and maintains the original aspect ratio.
	''' </returns>
	Private Shared Function ScaleSizeToFit(ByVal currentSize As Size, ByVal targetBounds As Size) As Size
		'If either Size object is empty, return the targetBounds.
		If targetBounds.IsEmpty OrElse currentSize.IsEmpty Then Return targetBounds

		'Get the ratio of the original and the target size
		Dim DiffX As Double = currentSize.Width / CInt(targetBounds.Width)
		Dim DiffY As Double = currentSize.Height / CInt(targetBounds.Height)

		'Set the multiplier to use to the largest of the results
		Dim multiplier As Double = CDbl(Math.Max(DiffX, DiffY))
		Return New Size(CInt(currentSize.Width / multiplier), CInt(currentSize.Height / multiplier))
	End Function

	Private Shared Function ImageIsIndexed(ByVal image As Image) As Boolean
		Select Case image.PixelFormat
			Case PixelFormat.Format1bppIndexed
				Exit Select
			Case PixelFormat.Format4bppIndexed
				Exit Select
			Case PixelFormat.Format8bppIndexed
				Exit Select
			Case Else
				Return False
		End Select

		Return True
	End Function

#Region " GDI+ Methods "
	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	Private Shared Function CreateCompatibleDC(ByVal hdc As IntPtr) As IntPtr
	End Function

	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	  Private Shared Function CreateBitmap(ByVal width As Long, ByVal height As Long, ByVal planes As Integer, ByVal bpp As Integer, ByVal lvpBits As IntPtr) As IntPtr
	End Function

	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	Private Shared Function CreateCompatibleBitmap(ByVal hdc As IntPtr, ByVal width As Integer, ByVal height As Integer) As IntPtr
	End Function

	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	Private Shared Function CreateDIBSection(ByVal hdc As IntPtr, ByRef bmi As BitmapInfoHeader, ByVal Usage As UInteger, ByRef bits As IntPtr, ByVal hSection As IntPtr, ByVal dwOffset As UInteger) As IntPtr
	End Function



	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	Private Shared Function SelectObject(ByVal hdc As IntPtr, ByVal hgdiobj As IntPtr) As IntPtr
	End Function

	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	Private Shared Function BitBlt(ByVal hdcDst As IntPtr, ByVal xDst As Integer, ByVal yDst As Integer, ByVal w As Integer, ByVal h As Integer, ByVal hdcSrc As IntPtr, _
	ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal rop As Integer) As Integer
	End Function

	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	Private Shared Function GetPixel(ByVal hdc As IntPtr, ByVal xPos As Integer, ByVal yPos As Integer) As Int32
	End Function

	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	  Private Shared Function SetPixel(ByVal hdc As IntPtr, ByVal xPos As Integer, ByVal yPos As Integer, ByVal newColor As Integer) As Int32
	End Function

	''' <summary>
	''' Sets the color of a pixel to approximately the color specified. (As close as possible)
	''' </summary>
	''' <param name="hdc">The pointer to the HBitmap to change</param>
	''' <param name="xPos">The x-coordinate of the color</param>
	''' <param name="yPos">The y-coordinate of the color</param>
	''' <param name="newColor">The new color to set to the pixel</param>
	''' <returns>Returns a boolean indicating the success of the operation</returns>
	''' <remarks>
	''' This method is faster than SetPixel because it only returns a boolean.
	''' </remarks>
	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	 Private Shared Function SetPixelV(ByVal hdc As IntPtr, ByVal xPos As Integer, ByVal yPos As Integer, ByVal newColor As Integer) As Boolean
	End Function

	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	Private Shared Function SetBkColor(ByVal hdc As IntPtr, ByVal color As Integer) As Integer
	End Function



	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	Private Shared Function DeleteObject(ByVal hObject As IntPtr) As Boolean
	End Function

	<System.Runtime.InteropServices.DllImport("gdi32.dll")> _
	Private Shared Function DeleteDC(ByVal hdc As IntPtr) As Integer
	End Function
#End Region

#Region " user32 Methods "
	<System.Runtime.InteropServices.DllImport("user32.dll")> _
	Private Shared Function InvalidateRect(ByVal hwnd As IntPtr, ByVal rect As IntPtr, ByVal bErase As Integer) As Integer
	End Function

	<System.Runtime.InteropServices.DllImport("user32.dll")> _
	Private Shared Function GetDC(ByVal hwnd As IntPtr) As IntPtr
	End Function

	<System.Runtime.InteropServices.DllImport("user32.dll")> _
	Private Shared Function ReleaseDC(ByVal hwnd As IntPtr, ByVal hdc As IntPtr) As Integer
	End Function
#End Region

#End Region

#Region " Private Structures "
	<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)> _
	Private Structure BitmapInfoHeader
		Public Size As UInteger
		Public Width As Integer
		Public Height As Integer
		Public Planes As Short
		Public BitCount As Short
		Public Compression As UInteger
		Public SizeImage As UInteger
		Public XPelsPerMeter As Integer
		Public YPelsPerMeter As Integer
		Public ClrUsed As UInteger
		Public ClrImportant As UInteger
		<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=256)> _
		Public cols As UInteger()
	End Structure
#End Region

#Region " Member Variables "
	Private Shared BI_RGB As UInteger = 0
	Private Shared DIB_RGB_COLORS As UInteger = 0
#End Region
End Class


'
'   Generated missing classes
'

Public Class ImageHelper
    Public Shared Function GetPageCount(original As Bitmap) As Integer
        Throw New NotImplementedException
    End Function

    Public Shared Sub SelectFrame(original As Bitmap, i As Integer)
        Throw New NotImplementedException
    End Sub

    Public Shared Sub WriteBitmapsToStream(results As Bitmap(), memoryStream As MemoryStream, compressionType As CompressionType, imageFormat As ImageFormat)
        Throw New NotImplementedException
    End Sub
End Class
