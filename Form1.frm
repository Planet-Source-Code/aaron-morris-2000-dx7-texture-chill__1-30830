VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Trance Tunnel .. By Aaron"
   ClientHeight    =   8550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   664
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const ayX As Integer = 640
Const ayY As Integer = 480
Dim tunn_angle(ayX, ayY) As Long
Dim tunn_dist(ayX, ayY) As Long
Const nSize As Long = 6

Function main()
    Dim done As Integer
    Dim x As Integer
    Dim y As Integer
    Dim i As Integer
    Dim DX As Long
    Dim dy As Long
    Dim ubase As Integer
    Dim vbase As Integer
    Dim rotation_u As Integer
    Dim rotation_v As Integer
    Dim u As Long
    Dim v As Long
    Dim rc As RECT
    Dim dst As Double
    Dim size As Long
    size = 1024 * nSize '512 * 1
    
  

    For y = 0 To ayY
        For x = 0 To ayX
           
            DX = x - ayX / 2
            dy = y - ayY / 2
            tunn_angle(x, y) = (Sin(dy) + Cos(DX)) '* 5
    
            dst = Sqr((DX * DX) + (dy * dy)) ' 50
        
            If (dst > 2) Then
                tunn_dist(x, y) = (size / dst)
    
            Else
                tunn_dist(x, y) = dst
            End If
            
            rotation_u = 20
            rotation_v = 0
            ubase = 0
            vbase = 0
        Next x
    Next y
        
        Do While (done = 0)
 
           
    
            For y = 0 To ayY Step nSize
                For x = 0 To ayX Step nSize
                    
                    u = (tunn_angle(x, y) + rotation_u)
                    v = (tunn_dist(x, y) + rotation_v)
                    u = u And 255
                    v = v And 255
                    xx = x
                    
                    yy = y
                    
                    rc.Left = u
                    rc.Right = u + nSize / 2
                    rc.Top = v
                    rc.Bottom = v + nSize / 2
                    
                    Call ddBuffer.ddSurface.BltFast(xx, yy, LoadSurface(0).ddSurface, rc, DDBLTFAST_WAIT)
                Next x
            Next y
            
            
            Call ddPrimary.ddSurface.Flip(ddBuffer.ddSurface, DDFLIP_WAIT)
           
            DoEvents
            
            ubase = ubase + nSize
            vbase = vbase + nSize
            
            rotation_u = ubase
            rotation_v = vbase
            
           If ubase > ayX Then ubase = 0
           If vbase > ayY Then vbase = 0
          
        Loop
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then End

End Sub

Private Sub Form_Load()

DX_Draw_SetUp Me.hWnd, 640, 480, 16, True
DDCreateSurface App.Path + "\texture.bmp"
main
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
