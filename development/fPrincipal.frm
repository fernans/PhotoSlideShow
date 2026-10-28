VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form fPrincipal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Dibujos"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   840
      Top             =   6840
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   5295
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   6855
      _cx             =   12091
      _cy             =   9340
      FlashVars       =   ""
      Movie           =   "flashA.swf"
      Src             =   "flashA.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
End
Attribute VB_Name = "fPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim rsNumMax As Long
Dim esperaCorta As Long
Dim esperaLarga As Long
Dim inicioRuta As String

Private Sub Flash_FlashCall(ByVal request As String)
    Dim objParser As MSXML.DOMDocument
    Dim x As MSXML.IXMLDOMNodeList
    Set objParser = New MSXML.DOMDocument
    objParser.async = "false"
    objParser.loadXML request
    
    Dim parametro As Collection
    Set parametro = New Collection
    
    Set x = objParser.getElementsByTagName("invoke")
    funcion = x(0).Attributes.getNamedItem("name").nodeValue
    Set x = objParser.getElementsByTagName("arguments")
    
    For i = 0 To x.length - 1
       If x(i).hasChildNodes Then
        parametro.Add x(i).childNodes(0).Text
       End If
    Next i
    
    '// -- ejecutar llamada
    
    '// ejemplo
    'If (funcion = "rutavideo") Then
    '    valor = CInt(parametro.Item(1))
    '    If (valor <= ruta.Count) Then
    '        t = ruta.Item(valor)
    '        Flash.SetReturnValue t
    '    End If
    'End If
    

End Sub

Private Sub Flash_FSCommand(ByVal command As String, ByVal args As String)
    If command = "alert" Then MsgBox args
    If command = "ESC" Then End
    If command = "buscar" Then
        Timer1.Enabled = False
        anterior = rs.AbsolutePosition
        rs.FindFirst "Id=" & "'" & args & "'"
        If rs.NoMatch Then
            rs.MoveFirst
            rs.Move anterior
            callFlashEscribirMensaje "Código incorrecto: " & args
            Timer1.Interval = esperaCorta
        Else
            ruta = inicioRuta & rs.Fields("ruta").Value
            callFlashCargarImagen ruta
            callFlashEscribirMensaje "Mostrando: " & args
            Timer1.Interval = esperaLarga
        End If
        Flash.SetVariable "_root.codigo", ""
        Timer1.Enabled = True
    End If
End Sub


Private Sub Form_Load()
    Math.Randomize
    Me.Flash.Movie = App.Path + "\dibujos.swf"
    
    Set db = OpenDatabase(App.Path & "\dibujos.mdb")
    
' leer configuracion
    Set rs = db.OpenRecordset("SELECT * FROM configuracion")
    rs.MoveFirst
    rs.FindFirst "clave='esperaCorta'"
    If Not rs.NoMatch Then
        esperaCorta = CLng(rs.Fields("valor").Value)
    Else
        esperaCorta = 1000
    End If
    rs.MoveFirst
    rs.FindFirst "clave='esperaLarga'"
    If Not rs.NoMatch Then
        esperaLarga = CLng(rs.Fields("valor").Value)
    Else
        esperaLarga = 5000
    End If
    rs.MoveFirst
    rs.FindFirst "clave='inicioRuta'"
    If Not rs.NoMatch Then
        inicioRuta = rs.Fields("valor").Value
    Else
        inicioRuta = "images\"
    End If

' configuracion de monitores
    rs.MoveFirst
    rs.FindFirst "clave='alto'"
    If Not rs.NoMatch Then
        alto = CLng(rs.Fields("valor").Value)
    Else
        alto = 1
    End If
    rs.MoveFirst
    rs.FindFirst "clave='ancho'"
    If Not rs.NoMatch Then
        ancho = CLng(rs.Fields("valor").Value)
    Else
        ancho = 1
    End If
    rs.MoveFirst
    rs.FindFirst "clave='top'"
    If Not rs.NoMatch Then
        Top = CLng(rs.Fields("valor").Value)
    Else
        Top = 0
    End If
    rs.MoveFirst
    rs.FindFirst "clave='left'"
    If Not rs.NoMatch Then
        Left = CLng(rs.Fields("valor").Value)
    Else
        Left = 0
    End If
' fin leer configuracion

    rs.Close
    Set rs = Nothing
    Set rs = db.OpenRecordset("SELECT * FROM dibujos")
    rs.MoveLast
    rs.MoveFirst
    rsNumMax = rs.RecordCount
     
    ' tamaño... en multiplos del monitor principal
    Me.Width = Screen.Width * ancho
    Me.Height = Screen.Height * alto
    ' posicion... respecto al (0,0) del monitor principal
    Me.Top = Screen.Height * Top
    Me.Left = Screen.Width * Left
    
End Sub

Private Sub Form_Resize()
    With Flash
    .Top = 0
    .Left = 0
    .Width = Me.Width
    .Height = Me.Height
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs.Close
    Set rs = Nothing
    db.Close
    Set db = Nothing
End Sub

Private Sub Timer1_Timer()

    Timer1.Enabled = False
    
    Do
        num = Int(rsNumMax * Rnd)
    Loop While ((rs.AbsolutePosition = num) And (rsNumMax > 1))
    
    rs.MoveFirst
    rs.Move (num)
    ruta = inicioRuta & rs.Fields("ruta").Value
    
    callFlashCargarImagen (ruta)
        
    Timer1.Interval = esperaCorta
    Timer1.Enabled = True
End Sub

Private Sub callFlashCargarImagen(ruta)
    com = Chr$(34)
    request = "<invoke name=" & com & "cargarImagen" & com & "returntype=" & com & "xml" & com & ">"
    request = request & "<arguments>"
    request = request & "<string>"
    request = request & ruta
    request = request & "</string>"
    request = request & "</arguments>"
    request = request & "</invoke>"
    Flash.CallFunction (request)
End Sub

Private Sub callFlashEscribirMensaje(mensaje)
    com = Chr$(34)
    request = "<invoke name=" & com & "escribeMensaje" & com & "returntype=" & com & "xml" & com & ">"
    request = request & "<arguments>"
    request = request & "<string>"
    request = request & mensaje
    request = request & "</string>"
    request = request & "</arguments>"
    request = request & "</invoke>"
    Flash.CallFunction (request)
End Sub

