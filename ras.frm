VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "PASSWORD RECOVERY BY John Papadakhs  PLeaZE VoTe "
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this prog is made by john papadakhs. I am using windows API'S to access the phone
'book of windows and get informations about the RAS CONNECTIONS available
'the ras api is RasGetCredentials. this programm shows the connection name of eatch dialup
'connection the username and the password. Only in windows XP the password is
'shown with "*". I'm working on an other api that solves that problem with XP
'and soon will have and XP passwords.also i'm working on getting the dialup number.
'all these will be send soon.that's a promise.



Dim llist As ListItem
Private Sub Form_Load()
'set the listview
ListView1.ColumnHeaders.Add , , "Connection Name", ListView1.Width / 3
ListView1.ColumnHeaders.Add , , "Username", ListView1.Width / 3
ListView1.ColumnHeaders.Add , , "Password", ListView1.Width / 3
'declarations for the use of the api
Dim rdp As VBRasDialParams

Dim b() As Byte
Dim rtn As Long


Dim sArray() As String
Dim iCtr As Integer
DUN_Services sArray 'here the connections names are stored in the sArray

For iCtr = 0 To UBound(sArray) 'here we take every connection name and use it to get
                               'get more infos about this connection by calling the
                               'VBRasGetEntryDialParams function
   rtn = VBRasGetEntryDialParams(b, vbNullString, sArray(iCtr))
   Call BytesToVBRasDialParams(b, rdp)
   'store the infos in the listview
     Set llist = ListView1.ListItems.Add(, , rdp.EntryName)
llist.ListSubItems.Add , , rdp.UserName
llist.ListSubItems.Add , , rdp.Password
Next
End Sub

