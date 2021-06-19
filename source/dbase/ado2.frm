VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDataEnv 
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCustomersTotal 
      DataField       =   "CustomersTotal"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   25
      Top             =   4815
      Width           =   660
   End
   Begin VB.TextBox txtTotalOrders 
      DataField       =   "TotalOrders"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   23
      Top             =   4435
      Width           =   660
   End
   Begin VB.TextBox txtFax 
      DataField       =   "Fax"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   21
      Top             =   4055
      Width           =   3375
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "Phone"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   19
      Top             =   3675
      Width           =   3375
   End
   Begin VB.TextBox txtCountry 
      DataField       =   "Country"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   17
      Top             =   3295
      Width           =   2475
   End
   Begin VB.TextBox txtPostalCode 
      DataField       =   "PostalCode"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   15
      Top             =   2915
      Width           =   1650
   End
   Begin VB.TextBox txtRegion 
      DataField       =   "Region"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   13
      Top             =   2535
      Width           =   2475
   End
   Begin VB.TextBox txtCity 
      DataField       =   "City"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   11
      Top             =   2155
      Width           =   2475
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   9
      Top             =   1775
      Width           =   3375
   End
   Begin VB.TextBox txtContactTitle 
      DataField       =   "ContactTitle"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   7
      Top             =   1395
      Width           =   3375
   End
   Begin VB.TextBox txtContactName 
      DataField       =   "ContactName"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   5
      Top             =   1015
      Width           =   3375
   End
   Begin VB.TextBox txtCompanyName 
      DataField       =   "CompanyName"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   3
      Top             =   635
      Width           =   3375
   End
   Begin VB.TextBox txtCustomerID 
      DataField       =   "CustomerID"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1995
      TabIndex        =   1
      Top             =   255
      Width           =   825
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "ado2.frx":0000
      Height          =   2640
      Left            =   150
      TabIndex        =   26
      Top             =   5250
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   4657
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      DataMember      =   "Command2"
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CustomersTotal:"
      Height          =   255
      Index           =   12
      Left            =   150
      TabIndex        =   24
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "TotalOrders:"
      Height          =   255
      Index           =   11
      Left            =   150
      TabIndex        =   22
      Top             =   4485
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   255
      Index           =   10
      Left            =   150
      TabIndex        =   20
      Top             =   4095
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Phone:"
      Height          =   255
      Index           =   9
      Left            =   150
      TabIndex        =   18
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Country:"
      Height          =   255
      Index           =   8
      Left            =   150
      TabIndex        =   16
      Top             =   3345
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PostalCode:"
      Height          =   255
      Index           =   7
      Left            =   150
      TabIndex        =   14
      Top             =   2955
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Region:"
      Height          =   255
      Index           =   6
      Left            =   150
      TabIndex        =   12
      Top             =   2580
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   255
      Index           =   5
      Left            =   150
      TabIndex        =   10
      Top             =   2205
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   255
      Index           =   4
      Left            =   150
      TabIndex        =   8
      Top             =   1815
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ContactTitle:"
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ContactName:"
      Height          =   255
      Index           =   2
      Left            =   150
      TabIndex        =   4
      Top             =   1065
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CompanyName:"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   675
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CustomerID:"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   300
      Width           =   1815
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
