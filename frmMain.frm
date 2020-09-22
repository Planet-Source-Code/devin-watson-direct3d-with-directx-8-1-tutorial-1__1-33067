VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "My First Direct3D Program"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRender 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   60
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRender 
         Caption         =   "&Render"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Author: Devin Watson
'Original C++ code from "Special Effects Game Programming with DirectX"
'by Mason McCuskey

'If you have any questions about this, you can e-mail me
'at dwatson@erols.com

'Some form-level globals
Private mDX As DirectX8                     'DirectX object (ALWAYS NEED THIS)
Private mDX3D As Direct3D8                  'Direct3D object (need this for 3D)
Private mDX3DDevice As Direct3DDevice8      'Direct3D Device object (for output to video card)
Private mVertBuff As Direct3DVertexBuffer8  'Vertex Buffer object (need this to hold our triangle)
Private CanRun As Boolean                   'Flag that prevents some bad things from happening
Private Const PI = 3.14159                  'For the rotation calculations
Private Const D3DADAPTER_DEFAULT = 0        'Taken from one of the C++ headers. It isn't
                                            'defined in the Type Library for VB.

'Another custom constant, which tells Direct3D
'that we're using the Flexible Vertext Format (FVF) with some
'diffuse colors so the triangle will be "self-illuminating"
Private Const D3DFVF_CUSTOMVERTEX As Long = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

'Our custom vertex for the FVF
Private Type CustomVertex
    X As Single
    Y As Single
    Z As Single
    Color As Long
End Type

'Need this to calculate size when creating the Vertex
'Buffer and also when rendering.
Private testVert As CustomVertex

'Our triangle, which is composed of
'3 custom vertices (0->2)
Private MyTriangle(2) As CustomVertex

'We're going to use this for some arbitrary angle calculation.
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Cleanup()
    'This takes care of cleaning up
    'everything and returning the system
    'to normal.
    On Local Error Resume Next
    
    'Destroys the vertex buffer,
    'so it doesn't hang around
    'in memory. Generally, it
    'doesn't hurt you to check for
    'these things on exit explicitly.
    If Not mVertBuff Is Nothing Then
        Set mVertBuff = Nothing
    End If
    
    'Destroys the Direct3D device,
    'relinquishing control back
    'to Windows for its window.
    If Not mDX3DDevice Is Nothing Then
        Set mDX3DDevice = Nothing
    End If
    
    'And last but not least, we
    'take out Direct3D itself.
    If Not mDX3D Is Nothing Then
        Set mDX3D = Nothing
    End If
End Sub


Private Sub InitD3D()
    'Initializes Direct3D and
    'gathers some information. If
    'you wanted to be paranoid, you
    'could check for video card features
    'here and manually set things
    'according to what you need.
    
    On Local Error Resume Next
    
    Dim DisplayMode As D3DDISPLAYMODE
    Dim d3dpp As D3DPRESENT_PARAMETERS
    
    'First, we create a Direct3D object.
    'This provides interfaces to all
    'of our other objects we can use
    'to generate our scene.
    Set mDX3D = mDX.Direct3DCreate
    
    'Since there is no such thing as the FAILED() macro
    'like in the C++ library, we have to be a little
    'more verbose in our error checking.
    If mDX3D Is Nothing Then
        MsgBox "Could not create base Direct3D system.", vbOKOnly + vbCritical, "Error"
        CanRun = False
        Exit Sub
    End If
    
    'Get the default video card device information
    mDX3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DisplayMode
    
    'We'll run in Windowed mode for now, thank you
    'Try it with Windowed = False for fullscreen
    d3dpp.Windowed = True
    
    'I set this just to make sure that
    'even if I muck up Present(), DirectX
    'has something to fall back on.
    d3dpp.hDeviceWindow = Me.hWnd
    
    'We're not really going to do anything
    'advanced, so let's just just get rid of
    'what we swap out of the back buffer.
    d3dpp.SwapEffect = D3DSWAPEFFECT_DISCARD
    'And make sure the back buffer pixel format is compatible
    'with the video card's screen pixel format.
    d3dpp.BackBufferFormat = DisplayMode.Format
    
    'Now, let's create our Direct3D device, now that
    'we've got all of the information!
    Set mDX3DDevice = mDX3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, _
            D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    
    'If it failed, we have to exit the subroutine, and set a flag,
    'so that future processing can decide if it needs
    'to run or not.
    If mDX3DDevice Is Nothing Then
        MsgBox "Could not create Direct3D Device from DirectX 8!", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    
    'Turn off culling. This allows us
    'to see both front and back of objects.
    mDX3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
    'Turn off lighting as well. This
    'object will be "self-luminous", since
    'all of our vertexes will be colored.
    mDX3DDevice.SetRenderState D3DRS_LIGHTING, False
    CanRun = True
End Sub


Private Sub InitGeometry()
    'Creates the vertex buffer we
    'will be displaying. Basically,
    'the easiest way to communicate
    'with DirectX as to what
    'you want it to display is
    'to talk in terms of triangles, otherwise
    'known as an array or collection of
    '3 vertices (3-D coordinates)
    
    On Local Error Resume Next
    'If InitD3D failed, we can just exit now.
    If CanRun = False Then Exit Sub
    Dim RC As Long
    
    'NOTE: I am using the helper function D3DColorRGBA
    'to produce color values for each vertex, but,
    'you could easily use the regular RGB() function
    'built into VB. I don't because I like to use
    'the "DX-native functions". If this doesn't work
    'on your video card, try using D3DColorXRGB or
    'D3DColorARGB.
    
    'Also, try changing the Z component
    'to a different value for some interesting warping
    'during the rotation.
    
    
    'First vertex: Lower left-hand corner
    MyTriangle(0).X = -1
    MyTriangle(0).Y = -1
    MyTriangle(0).Z = 0
    MyTriangle(0).Color = D3DColorRGBA(0, 255, 0, 0)
    
   
    
    'Second vertex: Lower right-hand corner
    MyTriangle(1).X = 1
    MyTriangle(1).Y = -1
    MyTriangle(1).Z = 0
    MyTriangle(1).Color = D3DColorRGBA(255, 0, 255, 0)
    
    'Third vertex: Top
    MyTriangle(2).X = 0
    MyTriangle(2).Y = 1
    MyTriangle(2).Z = 0
    MyTriangle(2).Color = D3DColorRGBA(255, 255, 255, 0)
    
    'Now that we've defined our vertex, we need to
    'set up the vertex processing pipeline
    'to accept this custom format, using the
    'Flexible Vertex Format (FVF)
    Set mVertBuff = mDX3DDevice.CreateVertexBuffer(3 * Len(testVert), 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error creating Vertex Buffer: Invalid Call", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    
    'Well, now that we've created the vertex buffer,
    'based on our own custom FVF, we
    'need to fill it.
    'We can use this helper function,
    'provided by Direct3D, to lock,
    'fill, and unlock the vertex
    'buffer all in one line. Neat, eh?
    RC = D3DVertexBuffer8SetData(mVertBuff, 0, 3 * LenB(testVert), 0, MyTriangle(0))
    
    'Since this function does not set Err.Number,
    'we need to check against this known constant
    'to make sure it executed correctly.
    If RC = D3DERR_INVALIDCALL Then
        MsgBox "Invalid call to D3DVertextBuffer8SetData()!", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    
End Sub

Private Sub Render()
    'The main render routine.
    On Local Error Resume Next
    If CanRun = False Then
        MsgBox "Cannot run: Failure in CanRun flag!", vbOKOnly + vbCritical, "ERROR"
        Exit Sub
    End If
    'Clear the back buffer to a light blue
    'color.
    mDX3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorXRGB(0, 0, 128), 1, 0
    
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error calling D3DDevice8.Clear(): " & Err.Description, vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    'Begin rendering the scene.
    mDX3DDevice.BeginScene
    SetupMatrices
    
    'Render the vertex buffer contents
    mDX3DDevice.SetStreamSource 0, mVertBuff, LenB(testVert)
    mDX3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
    mDX3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 1
    
    mDX3DDevice.EndScene
    
    'This is the equivalent to a Blt()
    'operation in the good ol' days of
    'DX 7.0
    mDX3DDevice.Present ByVal 0, ByVal 0, ByVal 0, ByVal 0
    
End Sub

Private Sub SetupMatrices()
    'This sets up the World, View,
    'and Projection transform matrices.
    On Local Error Resume Next
    
    'For the World, we'll rotate along
    'the Y-axis. We're using timeGetTime()
    'from the Win32 API to derive an
    'arbitrary angle from which to rotate to.
    Dim matWorld As D3DMATRIX
    Dim matView As D3DMATRIX
    Dim matProj As D3DMATRIX
    
    'These vectors are needed to calculate a view,
    'as they show where we are, what we
    'are looking at, and which way is up.
    Dim vecEye As D3DVECTOR
    Dim vecAt As D3DVECTOR
    Dim vecUp As D3DVECTOR
    
    D3DXMatrixRotationY matWorld, (timeGetTime / 150#)
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error rotating World Matrix: " & Err.Description, vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    'Now that we've got it rotated,
    'we apply the transformation to the World
    mDX3DDevice.SetTransform D3DTS_WORLD, matWorld
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error setting transform on world: Invalid procedure call.", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    'Now we set up the View Matrix. This one
    'is a little trickier. First, we need
    'to set the 3 Vectors for positioning everything.
    
    'The first one defines where our position is.
    vecEye.X = 0#
    vecEye.Y = 3#
    vecEye.Z = -5#
    
    'The second defines what we are looking at
    'in the 3D World. In this case, it is (0,0,0),
    'or, the origin of the entire 3D World.
    vecAt.X = 0#
    vecAt.Y = 0#
    vecAt.Z = 0#
    
    'The third vector is our normal, which tells
    'us which way is up.
    vecUp.X = 0#
    vecUp.Y = 1#
    vecUp.Z = 0#
    
    'And we make the View Matrix!
    D3DXMatrixLookAtLH matView, vecEye, vecAt, vecUp
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error calling D3DXMatrixLookAtLH()!", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    
    mDX3DDevice.SetTransform D3DTS_VIEW, matView
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error setting transform on view: Invalid procedure call.", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    'Whew! Now for the Projection Matrix. This one
    'isn't nearly as rough.
    D3DXMatrixPerspectiveFovLH matProj, PI / 4, 1#, 1#, 100#
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error calling D3DXMatrixPerspectiveFovLH()!", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    
    
    mDX3DDevice.SetTransform D3DTS_PROJECTION, matProj
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error setting transform on projection: Invalid procedure call.", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    
End Sub

Private Sub Form_Load()
    'Make sure we start up DirectX itself.
    'Otherwise, we can't start up
    'any other component inside of
    'DX.
    Set mDX = New DirectX8
    
    CanRun = True
    
    'Start up Direct3D. This starts up
    'the Direct3D Device, and also
    'gets us some parameters and info
    'about the video board.
    InitD3D
    InitGeometry
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'We're exiting, so we should
    'gracefully leave by cleaning
    'up memory.
    Cleanup
End Sub


Private Sub mnuExit_Click()
    'Calls Form_Unload()
    Unload Me
End Sub


Private Sub mnuRender_Click()
    'This just turns the (very)
    'primitive rendering system
    'on and off. In a more refined
    'system, this would consist
    'of a framelocked loop. Maybe
    'if anyone asks, I'll put that
    'into another tutorial.
    tmrRender.Enabled = Not tmrRender.Enabled
    mnuRender.Checked = tmrRender.Enabled
End Sub


Private Sub tmrRender_Timer()
    'Calls the Render routine.
    Render
End Sub


