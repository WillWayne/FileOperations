Namespace Tiff.Enums
    ''' <summary>
    '''     Specifies a raster-operation code. These codes define how the color data for the
    '''     source rectangle is to be combined with the color data for the destination
    '''     rectangle to achieve the final color.
    ''' </summary>
    Public Enum TernaryRasterOperations As UInteger
        ''' <summary>dest = source</summary>
        SRCCOPY = &Hcc0020
        ''' <summary>dest = source OR dest</summary>
        SRCPAINT = &Hee0086
        ''' <summary>dest = source AND dest</summary>
        SRCAND = &H8800c6
        ''' <summary>dest = source XOR dest</summary>
        SRCINVERT = &H660046
        ''' <summary>dest = source AND (NOT dest)</summary>
        SRCERASE = &H440328
        ''' <summary>dest = (NOT source)</summary>
        NOTSRCCOPY = &H330008
        ''' <summary>dest = (NOT src) AND (NOT dest)</summary>
        NOTSRCERASE = &H1100a6
        ''' <summary>dest = (source AND pattern)</summary>
        MERGECOPY = &Hc000ca
        ''' <summary>dest = (NOT source) OR dest</summary>
        MERGEPAINT = &Hbb0226
        ''' <summary>dest = pattern</summary>
        PATCOPY = &Hf00021
        ''' <summary>dest = DPSnoo</summary>
        PATPAINT = &Hfb0a09
        ''' <summary>dest = pattern XOR dest</summary>
        PATINVERT = &H5a0049
        ''' <summary>dest = (NOT dest)</summary>
        DSTINVERT = &H550009
        ''' <summary>dest = BLACK</summary>
        BLACKNESS = &H42
        ''' <summary>dest = WHITE</summary>
        WHITENESS = &Hff0062
        ''' <summary>
        ''' Capture window as seen on screen.  This includes layered windows
        ''' such as WPF windows with AllowsTransparency="true"
        ''' </summary>
        CAPTUREBLT = &H40000000
    End Enum
End NameSpace