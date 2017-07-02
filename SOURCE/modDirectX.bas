Attribute VB_Name = "modDirectX"
Public DX As DirectX7
Public DD As DirectDraw7
Public DD_Clip As DirectDrawClipper

Public DD_PrimarySurf As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

Public DD_PlayerIconSurf As DirectDrawSurface7
Public DDSD_PlayerIcon As DDSURFACEDESC2

Public rec As RECT
Public rec_pos As RECT

Sub InitDirectX()
    On Error GoTo DXErr

    ' Initialize DirextX
    Set DX = New DirectX7

    ' Initialize DirectDraw
    Set DD = DX.DirectDrawCreate(vbNullString)

    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmWorldMap.hWnd, DDSCL_NORMAL)

    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_SYSTEMMEMORY
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)

    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)



    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmWorldMap.Picture1.hWnd

    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip

    ' Initialize all surfaces
    Call InitSurfaces
    Exit Sub

    ' Error handling
DXErr:
    Call MsgBox("Error initializing DirectDraw! Make sure you have DirectX 7 or higher installed and a compatible graphics device. Err: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    Call GameDestroy
    End
End Sub



Sub InitSurfaces()
    


    ' Init Player Icon ddsd type and load the bitmap
    DDSD_PlayerIcon.lFlags = DDSD_CAPS
    DDSD_PlayerIcon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_PlayerIconSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\MapGFX.bmp", DDSD_PlayerIcon)
    SetMaskColorFromPixel DD_PlayerIconSurf, 0, 0


End Sub

Sub DestroyDirectX()

    Set DX = Nothing
    Set DD = Nothing

    Set DD_PlayerIconSurf = Nothing

    Set DD_PrimarySurf = Nothing
    Set DD_BackBuffer = Nothing
    
    
End Sub

Public Sub GameDestroy()

Call DestroyDirectX

End

End Sub

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal y As Long)
    Dim TmpR As RECT
    Dim TmpDDSD As DDSURFACEDESC2
    Dim TmpColorKey As DDCOLORKEY

    With TmpR
        .Left = X
        .Top = y
        .Right = X
        .Bottom = y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .Low = TheSurface.GetLockedPixel(X, y)
        .High = .Low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey

    TheSurface.Unlock TmpR
End Sub
