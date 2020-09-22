VERSION 5.00
Object = "{5ECE973D-0ADB-11D4-ACD6-0080C878CC01}#1.0#0"; "TIMECHART.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows-Standard
   Begin timechart.Urlaub Urlaub1 
      Height          =   4995
      Left            =   450
      TabIndex        =   5
      Top             =   660
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   8811
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Farbe_Hintergrund=   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Farbe_Urlaub    =   8388608
      Schriftfarbe_urlaub=   16777215
      Farbe_HalberTag =   8421631
      Farbe_einzelner_Tag=   49152
   End
   Begin VB.OptionButton Option_Quartal 
      Caption         =   "4. Quartal"
      Height          =   285
      Index           =   3
      Left            =   8160
      TabIndex        =   4
      Top             =   180
      Width           =   1695
   End
   Begin VB.OptionButton Option_Quartal 
      Caption         =   "3. Quartal"
      Height          =   285
      Index           =   2
      Left            =   6420
      TabIndex        =   3
      Top             =   180
      Width           =   1695
   End
   Begin VB.OptionButton Option_Quartal 
      Caption         =   "2. Quartal"
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   180
      Width           =   1695
   End
   Begin VB.OptionButton Option_Quartal 
      Caption         =   "1. Quartal"
      Height          =   285
      Index           =   0
      Left            =   2940
      TabIndex        =   1
      Top             =   180
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   450
      TabIndex        =   0
      Top             =   90
      Width           =   2235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim i As Integer
   Dim bla As Boolean
   Dim qua As Integer
   For i = 0 To 3
        If Option_Quartal(i).Value = True Then qua = i + 1
   Next i
    With Urlaub1
        .Datenbankpfad = App.Path & "\sample.mdb"
        .Tabelle_DATUM = "URLAUB"
        .Tabelle_NAMEN = "MITARBEITER"
        .TabFeld_ANZAHL_URLAUBSTAGE = "ANZAHL_URLAUBSTAGE"
        .TabFeld_Datum_bis = "URLAUB_BIS"
        .TabFeld_Datum_von = "URLAUB_VON"
        .TabFeld_NAMEN = "NAME"
        .TabFeld_NUMMER = "MA_NR"
        .TabFeld_Sonderurlaub = "SONDERURLAUB"
        .Quartal = qua
    End With
    bla = Urlaub1.show_chart
End Sub


