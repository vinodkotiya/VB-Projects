VERSION 5.00
Begin VB.Form frmDataEnv 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "last"
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "next"
      Height          =   375
      Left            =   3600
      TabIndex        =   24
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "previous"
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "first"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtFax 
      DataField       =   "Fax"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   21
      Top             =   4400
      Width           =   3375
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "Phone"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   4020
      Width           =   3375
   End
   Begin VB.TextBox txtCountry 
      DataField       =   "Country"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Top             =   3640
      Width           =   2475
   End
   Begin VB.TextBox txtPostalCode 
      DataField       =   "PostalCode"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   3260
      Width           =   1650
   End
   Begin VB.TextBox txtRegion 
      DataField       =   "Region"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   2880
      Width           =   2475
   End
   Begin VB.TextBox txtCity 
      DataField       =   "City"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   2500
      Width           =   2475
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   2120
      Width           =   3375
   End
   Begin VB.TextBox txtContactTitle 
      DataField       =   "ContactTitle"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   1740
      Width           =   3375
   End
   Begin VB.TextBox txtContactName 
      DataField       =   "ContactName"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   1360
      Width           =   3375
   End
   Begin VB.TextBox txtCompanyName 
      DataField       =   "CompanyName"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   980
      Width           =   3375
   End
   Begin VB.TextBox txtCustomerID 
      DataField       =   "CustomerID"
      DataMember      =   "Customers"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   255
      Index           =   10
      Left            =   -765
      TabIndex        =   20
      Top             =   4445
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Phone:"
      Height          =   255
      Index           =   9
      Left            =   -765
      TabIndex        =   18
      Top             =   4065
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Country:"
      Height          =   255
      Index           =   8
      Left            =   -765
      TabIndex        =   16
      Top             =   3685
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PostalCode:"
      Height          =   255
      Index           =   7
      Left            =   -765
      TabIndex        =   14
      Top             =   3305
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Region:"
      Height          =   255
      Index           =   6
      Left            =   -765
      TabIndex        =   12
      Top             =   2925
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   255
      Index           =   5
      Left            =   -765
      TabIndex        =   10
      Top             =   2545
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   255
      Index           =   4
      Left            =   -765
      TabIndex        =   8
      Top             =   2165
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ContactTitle:"
      Height          =   255
      Index           =   3
      Left            =   -765
      TabIndex        =   6
      Top             =   1785
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ContactName:"
      Height          =   255
      Index           =   2
      Left            =   -765
      TabIndex        =   4
      Top             =   1405
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CompanyName:"
      Height          =   255
      Index           =   1
      Left            =   -765
      TabIndex        =   2
      Top             =   1025
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CustomerID:"
      Height          =   255
      Index           =   0
      Left            =   -765
      TabIndex        =   0
      Top             =   645
      Width           =   1815
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataEnvironment1.rsCustomers.MoveFirst
End Sub

Private Sub Command2_Click()
If DataEnvironment1.rsCustomers.BOF Then
 Beep
Else
DataEnvironment1.rsCustomers.MovePrevious
 If DataEnvironment1.rsCustomers.BOF Then
 DataEnvironment1.rsCustomers.MoveFirst
 End If
End If
 End Sub

Private Sub Command3_Click()
If DataEnvironment1.rsCustomers.EOF Then
 Beep
 Else
  DataEnvironment1.rsCustomers.MoveNext
  If DataEnvironment1.rsCustomers.EOF Then
   DataEnvironment1.rsCustomers.MoveLast
  End If
 End If
 

End Sub

Private Sub Command4_Click()
DataEnvironment1.rsCustomers.MoveLast
End Sub
