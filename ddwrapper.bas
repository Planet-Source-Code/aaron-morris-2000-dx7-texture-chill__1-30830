Attribute VB_Name = "modGFX"
Option Explicit

Public DX As New DirectX7
Public DDraw As DirectDraw7

Public bFS As Boolean 'full screen flag
Public hWnd As Long 'Handle to Window

'The DD7 Surfaces
Private Type DDSurf
    ddSurface As DirectDrawSurface7
    ddDescription As DDSURFACEDESC2
    ddClipper As DirectDrawClipper
End Type

Public DDSurfFileCount As Integer
Public LoadSurface() As DDSurf

Public ddPrimary As DDSurf 'Primary Buffer - We always have a Primary buffer
Public ddBuffer As DDSurf 'Back Buffer


'    Example of Full screen
'    1. Set the full screen flag to true for a direct draw full screen application
'    DX_Draw_SetUp Me.hWnd, 800, 600, 16, True
'    DDCreateSurface "c:\1.bmp"
'    Draw 100, 10


'    Example of Windowed
'    1. Set the full screen flag to false for a direct draw windows application
'    DX_Draw_SetUp Me.hWnd, 0, 0, 16, False
'    DDCreateSurface "c:\1.bmp"
'    Draw 100, 10



'---Function------
'DX_Draw_SetUp
'-----------------
'mHandle = Handle to DC - E.G. form.hWnd
'mHeight = Screen Res height - E.G.600, 480
'mWidth  = Screen Res Width - E.G. 800,640
'mCDepth = Color Depth - E.G 16,24,32

Public Function DX_Draw_SetUp(mHandle As Long, mWidth As Integer, mHeight As Integer, mCDepth As Integer, Optional mFullScreen As Boolean)
    On Error GoTo DDraw_SetUp_Error
    
    bFS = mFullScreen
    hWnd = mHandle
    'Create DD Obj...
    Set DDraw = DX.DirectDrawCreate("")
    
    
     If mFullScreen Then    'Full Screen Flag Set to true ... Set DD as full screen!
            
            
            
            
            Call DDraw.SetCooperativeLevel(mHandle, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
            DDraw.SetDisplayMode mWidth, mHeight, mCDepth, 0, DDSDM_DEFAULT
                   
            ddPrimary.ddDescription.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
            ddPrimary.ddDescription.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
            ddPrimary.ddDescription.lBackBufferCount = 1
        
            Set ddPrimary.ddSurface = DDraw.CreateSurface(ddPrimary.ddDescription)
    
            Dim caps As DDSCAPS2
            
            caps.lCaps = DDSCAPS_BACKBUFFER
    
            Set ddBuffer.ddSurface = ddPrimary.ddSurface.GetAttachedSurface(caps)
            ddBuffer.ddSurface.GetSurfaceDesc ddBuffer.ddDescription
     
            'Use black for transparent color key (&h0)
            Dim key As DDCOLORKEY
            
            key.low = 0
            key.high = 0
            
            ddBuffer.ddSurface.SetColorKey DDCKEY_SRCBLT, key
        Else 'For DDraw Windowed
            
            
            Call DDraw.SetCooperativeLevel(mHandle, DDSCL_NORMAL)
        
            ddPrimary.ddDescription.lFlags = DDSD_CAPS
            ddPrimary.ddDescription.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    
            Set ddPrimary.ddSurface = DDraw.CreateSurface(ddPrimary.ddDescription)
        
            'DD7 Clipper --- So we don't draw outside the window or hWnd we select as mHandle
            Set ddPrimary.ddClipper = DDraw.CreateClipper(0)
            ddPrimary.ddClipper.SetHWnd mHandle
            ddPrimary.ddSurface.SetClipper ddPrimary.ddClipper
            
        
        
        End If
        
        Exit Function

'Handles any errors
DDraw_SetUp_Error:
    MsgBox "An Error Occurred While Setting Up DirectDraw 7", 0, "Error"
    End
    
End Function

Public Function DDCreateSurface(sFileName As String)
On Error GoTo Err_DDCreateSurface

    ReDim Preserve LoadSurface(DDSurfFileCount)
        
    Set LoadSurface(DDSurfFileCount).ddSurface = Nothing
    
  
    LoadSurface(DDSurfFileCount).ddDescription.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    LoadSurface(DDSurfFileCount).ddDescription.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    LoadSurface(DDSurfFileCount).ddDescription.lWidth = 600
    LoadSurface(DDSurfFileCount).ddDescription.lHeight = 600
    
    Set LoadSurface(DDSurfFileCount).ddSurface = DDraw.CreateSurfaceFromFile(sFileName, LoadSurface(DDSurfFileCount).ddDescription)
    
    DDSurfFileCount = DDSurfFileCount + 1

    Exit Function
    
Err_DDCreateSurface:
    If Err = 91 Then
        MsgBox "Direct Draw Must Be Set-Up Before Loading A Bitmap!", 0, "Error"
        End
    Else
        MsgBox "" + Error$ + "   =  " & Err, 0, "Error"
        End
    End If
    
    
End Function

Public Function Draw(x As Integer, y As Integer)
    Dim srcRect As RECT
    Dim dstRect As RECT
    
    srcRect.Left = 0
    srcRect.Right = LoadSurface(DDSurfFileCount - 1).ddDescription.lWidth
    srcRect.Top = 0
    srcRect.Bottom = LoadSurface(DDSurfFileCount - 1).ddDescription.lHeight
    
    
    If bFS = False Then 'Windowed
        DX.GetWindowRect hWnd, dstRect
    End If
        
    dstRect.Left = x
    dstRect.Right = x + srcRect.Right
    dstRect.Top = y
    dstRect.Bottom = y + srcRect.Bottom
    
    If bFS = False Then 'Windowed
        Call ddPrimary.ddSurface.Blt(dstRect, LoadSurface(DDSurfFileCount - 1).ddSurface, srcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    Else
        'Use the back buffer
        DDClear 'Clear the back buffer
        Call ddBuffer.ddSurface.Blt(dstRect, LoadSurface(DDSurfFileCount - 1).ddSurface, srcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        Call ddPrimary.ddSurface.Flip(ddBuffer.ddSurface, DDFLIP_WAIT)
    End If
    
End Function

Public Function DDRestore()
    DDraw.RestoreAllSurfaces
End Function


Public Sub DDClear()

    Dim dstRect As RECT

    With dstRect
        .Top = 0
        .Bottom = Screen.Height
        .Left = 0
        .Right = Screen.Width
    End With
    
    'Fill the entire backbuffer
    ddBuffer.ddSurface.BltColorFill dstRect, 0
    

End Sub



