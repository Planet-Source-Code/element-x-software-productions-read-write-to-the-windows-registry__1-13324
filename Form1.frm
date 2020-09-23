VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Read/Write Registry"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read Registry"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write Registry"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Your Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' This is a what each thing does
' SaveSetting: Saves the following data to the registry
' Element-X Software: This creates a new Group in the registry under the path; HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Element-X Software
' Main is the Sub Folder under the Group Element-X Software
' Entry is the name of the registry file that contains your text
' Text1.Text is the text that will be entered into the file Entry under Main in the Group Element-X Software

SaveSetting "Element-X Software", "Main", "Entry", Text1.Text
End Sub

Private Sub Command2_Click()
' This is what each thing does
' GetSetting: Gets the registry entry
' Element-X Software: This is the group it is to look for
' Main: This is the folder to look under in the Element-X Software Group
' Entry: This is the file to display the contents of

MsgBox GetSetting("Element-X Software", "Main", "Entry")
End Sub

