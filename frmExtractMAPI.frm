VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExractMAPI 
   Caption         =   "MAPI Email Address Extractor  www.mewsoft.com"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save List As"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Selected"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6480
      Top             =   4440
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7435
      View            =   3
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Email (Click to sort)"
         Object.Width           =   8820
      EndProperty
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   6480
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   6360
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Label lblTotalEmails 
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Total Emails:"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lblFilteredEmails 
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Filtered Emails:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   5880
      Width           =   1215
   End
End
Attribute VB_Name = "frmExractMAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================
'Author: Mewsoft Corporation
'Web site: http://www.mewsoft.com
'Copyright (c) Mewsoft Corporation
'Mission:
'   We Provide Full Featured:
'   Free Forum Software open source code,
'   Auction Software open source code,
'   Classifieds Software open source code,
'   Directory and Pay Per Click Search Engine Software open source code
'   For Every Business Size With The Lowest Prices on The Earth To
'   Make Everyone Dreams Comes True. Open Source Code in Perl SQL based
'   database to provide the biggest and fastest solutions ever for our products.

'========================================================
Option Explicit

Private Sub cmdDelete_Click()

Dim intIndex As Integer
'For loop counter
Dim intLoop As Integer
'Number of items in ListView
Dim intCount As Integer
'Number of selected items
Dim intSelected As Integer
'Array to hold selected items
Dim arrItems() As ListItem

intCount = ListView1.ListItems.Count
intSelected = 0

'Loop through and retrieve the selected items
For intLoop = 1 To intCount
    If ListView1.ListItems(intLoop).Selected Then
        intSelected = intSelected + 1
        ReDim Preserve arrItems(1 To intSelected) As ListItem
        Set arrItems(intSelected) = ListView1.ListItems(intLoop)
    End If
Next

'Loop through in reverse and remove the selected items
For intLoop = UBound(arrItems) To LBound(arrItems) Step -1
    ListView1.ListItems.Remove arrItems(intLoop).Index
Next

lblFilteredEmails.Caption = ListView1.ListItems.Count
End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdExtract_Click()

Call ProcessEmails

End Sub

Sub ProcessEmails()
    
    Dim TotalEmails As Long
    Dim FilteredEmails As Long
    
    On Error GoTo Err_Handler
    
    
    Screen.MousePointer = vbHourglass
    
    MAPISession1.DownLoadMail = False
    MAPISession1.SignOn
    
    If Err <> 0 Then
        MsgBox "Logon Failure: " + Error$
    Else
        MAPIMessages1.SessionID = MAPISession1.SessionID
    End If

    ' Go through the mail list one by one
    Dim i As Integer
    
    MAPIMessages1.FetchUnreadOnly = False
    MAPIMessages1.FetchMsgType = ""
    
    MAPIMessages1.Fetch
     
    'Debug.Print MAPIMessages1.MsgCount & "Messages Found." & vbCrLf
    TotalEmails = MAPIMessages1.MsgCount
    lblTotalEmails.Caption = TotalEmails
    
    Dim ListItem As ListItem
    Dim itmFound As ListItem
    Dim Email As String
    
    FilteredEmails = 0
    
    For i = 0 To MAPIMessages1.MsgCount - 1
        MAPIMessages1.MsgIndex = i
        Email = MAPIMessages1.MsgOrigAddress
        If Email <> "" Then
            'Debug.Print MAPIMessages1.MsgSubject
            'attempt to locate the corresponding item in the listview
            Set itmFound = ListView1.FindItem(Email, lvwSubItem, , 0)
        
            If itmFound Is Nothing Then
                FilteredEmails = FilteredEmails + 1
                Set ListItem = ListView1.ListItems.Add()
                ListItem.Text = CStr(FilteredEmails)
                ListView1.ListItems(FilteredEmails).ListSubItems.Add , , Email
                lblFilteredEmails.Caption = FilteredEmails
            End If
        End If
    Next i
    
    lblFilteredEmails.Caption = FilteredEmails
    
    MAPISession1.SignOff
    
    Screen.MousePointer = vbDefault
    Exit Sub
Err_Handler:
    Resume
End Sub

Private Sub List1_Click()

End Sub

Private Sub cmdSaveAs_Click()
    Dim File As String
    Dim F As Integer
    Dim intLoop As Integer
        
    On Error GoTo Cancel
    
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    CommonDialog1.FileName = "OutlookList.txt"
    CommonDialog1.DialogTitle = "Save Mail List As"
    
    CommonDialog1.ShowSave
    File = CommonDialog1.FileName
    If File = "" Then Exit Sub
    'Debug.Print File
    
    Dim itmFound As ListItem
   
    F = FreeFile
    
    Open File For Output As #F
    

    'Loop through and retrieve the selected items
    For intLoop = 1 To ListView1.ListItems.Count
    Set itmFound = ListView1.ListItems(intLoop)
    Print #F, itmFound.ListSubItems(1)
    Next
    
    Close #F
    
Cancel:
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Static sorder As Integer
   
    sorder = Not sorder
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.SortOrder = Abs(sorder)
    ListView1.Sorted = True

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

'Item.Index

End Sub

Private Sub Timer1_Timer()
DoEvents
End Sub
