VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "v"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Paid"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5280
      TabIndex        =   54
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtprevbill 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3960
      TabIndex        =   51
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Add Previous Bill"
      Height          =   375
      Left            =   3960
      TabIndex        =   50
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox txtdue 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3960
      TabIndex        =   48
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtgrand 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   38
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Get Total"
      Height          =   615
      Left            =   360
      MaskColor       =   &H0080FF80&
      TabIndex        =   37
      Top             =   9240
      Width           =   6255
   End
   Begin VB.TextBox txttotalamount 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   35
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox txtothers 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   32
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox txtmaterials 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   31
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox txtsruch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   30
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtmort 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox txtmedical 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtidfee 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txtpca 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtcpc 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtcbu 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtbill 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtcubic 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtprev 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtpresent 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txt_prev_date 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txt_current_date 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   5055
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   5055
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txt_current_date2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Previous Bill"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3960
      TabIndex        =   53
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Previous"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   52
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Due Date"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3960
      TabIndex        =   49
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Billing Date"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      TabIndex        =   47
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consumer Type"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   46
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Barangay"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   45
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   44
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   43
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consumer ID"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   42
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Billing No."
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   41
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Meter Reading"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      TabIndex        =   40
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grand Total  >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   39
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Amount  >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   36
      Top             =   8760
      Width           =   2415
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Others  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   34
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Label Label114 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Materials-Fittings  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   33
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Surcharge  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   29
      Top             =   7680
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mortuary  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Medical  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ID Fee  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PCA  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CPC  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CBU  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Bill  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cubic Meter Consumed  >>"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Present"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Service Period"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prevbill As Double

Private Sub Command1_Click()
prevbill = txtprevbill.text
End Sub

Private Sub Command3_Click()
Dim prev As Double
Dim current As Double
Dim cubic As Double
Dim cbu As Double
Dim cpc As Double
Dim pca As Double
Dim idfee As Double
Dim medical As Double
Dim mortuary As Double
Dim surcharge As Double
Dim materials As Double
Dim others As Double
Dim totalamount As Double
Dim total As Double
Dim grandtotal As Double


If Trim(txtprevbill.text) = "" Then txtprevbill.text = "0.00"
If Trim(txtprev.text) = "" Then txtprev.text = "0.00"
If Trim(txtpresent.text) = "" Then txtpresent.text = "0.00"
If Trim(txtcbu.text) = "" Then txtcbu.text = "0.00"
If Trim(txtcpc.text) = "" Then txtcpc.text = "0.00"
If Trim(txtpca.text) = "" Then txtpca.text = "0.00"
If Trim(txtidfee.text) = "" Then txtidfee.text = "0.00"
If Trim(txtmedical.text) = "" Then txtmedical.text = "0.00"
If Trim(txtmort.text) = "" Then txtmort.text = "0.00"
If Trim(txtsruch.text) = "" Then txtsruch.text = "0.00"
If Trim(txtmaterials.text) = "" Then txtmaterials.text = "0.00"
If Trim(txtothers.text) = "" Then txtothers.text = "0.00"

prev = Val(txtprev.text)
current = Val(txtpresent.text)
cbu = Val(txtcbu.text)
cpc = Val(txtcpc.text)
pca = Val(txtpca.text)
idfee = Val(txtidfee.text)
medical = Val(txtmedical.text)
mortuary = Val(txtmort.text)
surcharge = Val(txtsruch.text)
materials = Val(txtmaterials.text)
others = Val(txtothers.text)

cubic = current - prev
total = cubic * 29.42
totalamount = total + cbu + cpc + pca + idfee + medical + mortuary + surcharge + materials + others
grandtotal = totalamount + prevbill
txtcubic.text = Format(cubic, "0.00")
txtbill.text = Format(total, "0.00")
txttotalamount.text = Format(totalamount, "0.00")
txtgrand.text = Format(grandtotal, "0.00")
End Sub

Private Sub Form_Load()
Dim previousDate As Date
Dim currentDate As Date
Dim due As Date
due = DateAdd("m", 6, currentDate)
previousDate = DateAdd("m", -8, currentDate)
txt_current_date.text = Format(currentDate, "5/24/2024")
txt_current_date2.text = Format(currentDate, "5/24/2024")
txt_prev_date.text = Format(previousDate, "mm/24/2024")
txtdue.text = Format(due, "mm/24/2024")
End Sub


