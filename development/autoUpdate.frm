VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Auto-Uptade"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Updating..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const rutaImages = "\images"
Const rutaUpdate = "\update"
Const extension = ".jpg"
Const basedatos = "\dibujos.mdb"

Dim db As Database
Dim rs As Recordset

Dim Cancelar As Boolean

Private Sub bCancelar_Click()
    If (Cancelar) Then
        End
    Else
        Cancelar = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        rs.Close
        Set rs = Nothing
        db.Close
        Set db = Nothing
End Sub

Private Sub Form_Activate()
    Cancelar = False

    Dim fso As New FileSystemObject
    
    Dim sImages As String
    sImages = App.Path + rutaImages
    Dim fImages As Folder
    Set fImages = fso.GetFolder(sImages)
    
    Dim sUpdate As String
    sUpdate = App.Path + rutaUpdate
    Dim fUpdate As Folder
    Set fUpdate = fso.GetFolder(sUpdate)
    
    Dim cFiles As Files
    Set cFiles = fUpdate.Files
    Dim filename As String
    
    Dim codigo As String
    Dim fFile As File
    
    Set db = OpenDatabase(App.Path & basedatos)
    Set rs = db.OpenRecordset("SELECT * FROM dibujos")
    rs.MoveFirst
    
    Dim agregados As Long
    agreagados = 0
    PBar.Max = cFiles.Count
    PBar.Min = 0
    PBar.Value = 0
    
    For Each fFile In cFiles
        If Not Cancelar Then
            filename = fFile.Name
            posjpg = InStr(filename, extension)
            If (posjpg > 0) Then
                codigo = Left(filename, posjpg - 1)
                If (IsNumeric(codigo)) Then
                    Label2.Caption = codigo & " - " & filename
                    'buscamos en la bd
                    rs.FindFirst CStr("id='" & codigo & "'")
                    If (rs.NoMatch) Then
                        'el codigo no está repetido
                        'introducir en bd
                        rs.AddNew
                        rs.Fields("id").Value = codigo
                        rs.Fields("ruta").Value = filename
                        rs.Update
                        'eliminar archivo en destino si ya existe (cosa rara)
                        If (fso.FileExists(App.Path & rutaImages & "\" & filename)) Then
                            fso.DeleteFile (App.Path & rutaImages & "\" & filename), True
                        End If
                        'y mover el archivo
                        fFile.Move App.Path & rutaImages & "\"
                        
                        agregados = agregados + 1
                        Debug.Print "codigo OK: (" & codigo & ")"
                    Else
                        'codigo repetido...
                        Debug.Print "codigo repetido: (" & codigo & ")"
                    End If 'finde NoMatch
                End If 'finde isNumeric
            End If 'finde encontrado jpg
            PBar.Value = PBar.Value + 1
        End If 'finde cancelar
        
    Next 'finde each in cFiles
    
    Label1.Caption = "Completado"
    Label2.Caption = "Total archivos agregados: " & agregados
    rs.Close
    
    If (Cancelar) Then
        End
    Else
        bCancelar.Caption = "Cerrar"
        Cancelar = True
    End If
   
End Sub

