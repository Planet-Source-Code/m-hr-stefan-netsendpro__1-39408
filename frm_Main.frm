VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frm_Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "NetSend ver. 2.0"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txt_ToSend 
      Height          =   4545
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   540
      Width           =   3195
   End
   Begin MSComctlLib.Toolbar tbr_Main 
      Height          =   390
      Left            =   3000
      TabIndex        =   1
      Top             =   90
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "img_Toolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Send"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "beenden"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img_Toolbar 
      Left            =   4860
      Top             =   5250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":0724
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lst_Computers 
      Height          =   5520
      Left            =   60
      MultiSelect     =   2  'Erweitert
      TabIndex        =   0
      Top             =   90
      Width           =   2835
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Form_Load()

  Dim intIDX As Integer
  Dim ServerList As ListOfServer
  
  MousePointer = vbHourglass
  Me.lst_Computers.Clear
  'Get List of Computers on Network (Type all: Server and Workstation
  ServerList = EnumServer(SRV_TYPE_ALL)
  If ServerList.Init Then
      For intIDX = 1 To UBound(ServerList.List)
          Me.lst_Computers.AddItem ServerList.List(intIDX).ServerName
        End If
      Next
  End If
  
  MousePointer = vbDefault

End Sub

Private Sub tbr_Main_ButtonClick(ByVal Button As MSComctlLib.Button)

  Select Case Button.Key
    Case "Send":    Call SendMessage
    Case "beenden": Call ProgBeenden
  End Select

End Sub

Private Sub SendMessage()

On Error GoTo errHandler
Dim intCounter As Integer
Dim arrComputers() As String
  'Shell Netsend in the Windows Directory
Dim intIDX As Integer
  Me.MousePointer = vbHourglass
  DoEvents
  For intIDX = 0 To Me.lst_Computers.ListCount - 1
    If Me.lst_Computers.Selected(intIDX) = True Then
      Call Shell("net.exe send " & Me.lst_Computers.List(intIDX) & " " & Me.txt_ToSend.Text)
      'Call a short Programm Sleep that net.exe can execute without problems
      Call Sleep(500)
      'check counter if i can get all Pc's in List
      intCounter = intCounter + 1
    End If
  Next
  Me.txt_ToSend.Text = ""
  Me.MousePointer = vbDefault
  
  If intCounter = Me.lst_Computers.ListCount Then
    MsgBox ("Transmission successfull!"), vbOKOnly
  Else
    


  Exit Sub
  
  
  
errHandler:

  MsgBox "Error transmitting!" & vbCrLf & _
         "Recipient: " & Me.lst_Computers.List(intIDX) & vbCrLf & _
         "Fehler: " & Err.Description
         
End Sub

Private Sub ProgBeenden()

Dim intAnswer As Integer

  intAnswer = MsgBox("Wollen Sie das Programm wirklich beenden?", vbOKCancel)
  Select Case intAnswer
    Case 1: Unload Me
    Case 2: Exit Sub
  End Select

End Sub
