VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSetParent 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdo1 
      Left            =   960
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   9720
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   615
      Left            =   9240
      TabIndex        =   8
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tab1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tab2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tab3"
            ImageVarType    =   2
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2566
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10800
      Top             =   6840
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Play/Stop"
      Height          =   195
      Left            =   8280
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9600
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   3525
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra1 
      Height          =   1455
      Left            =   6960
      TabIndex        =   1
      Top             =   840
      Width           =   2415
      Begin WMPLibCtl.WindowsMediaPlayer wmp1 
         Height          =   1320
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   2400
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "none"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   4233
         _cy             =   2328
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Parent"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Path to video clip...."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmSetParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hwndMP As Long 'Variable used to contain an individual handle

'Api Declarations:
'Findwindow enables us to do as the name implies...find a particular window
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'SetParent allows us to set the parent/container of an object
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


Dim nOffSet As Integer 'This is used to hold the offset used to adjust the controls appearance/sizing when reparented to the statusbar

Dim strFilePath As String 'This variable is used to hold the path string for the video clip




Private Sub Check1_Click() 'Event fires when the user clicks the checkbox
     Debug.Print wmp1.playState
   If wmp1.playState = 3 Then  'Conditional: See what the players state is
      
      wmp1.Close 'Demonstrates that the control continues to work
   
   Else
      
      wmp1.URL = strFilePath 'Media not present or path not set or playback was stopped...reset
   
   End If 'End conditional statement
   
   
End Sub

Private Sub cmdBrowse_Click() 'Event fires when the user clicks the browse button

  On Error GoTo BrowseError 'This lets the app no what to do if an error occurs in within this routine
  
  cdo1.ShowOpen 'Show the open file dialog using the MS-Common Dialog control
  
BrowseError:  'Label identifies browse error section
   
   'An error occured... so we'll see what happened
    
    
   If Err.Number = 32755 Then 'This error number when returned by the common dialog control indicates that the user canceled the open file action
   
     'We know the user canceled so we can safely ignore this error
     
   ElseIf Err.Number = 0 Then  'The operation was successful
     
     'We know everything was ok so we can assign the path to the textbox
     Text1.Text = cdo1.FileName
     
   Else
     'A different error occured... so we'll notify the user
     'In this istance i use a messagebox
     'vbcrlf = Carriage Return Line Feed or more simply put: New Line
     'The & character is used to concatenate the string displayed to the user
     'The error properties accessed here are rather self explanatory
      MsgBox "An Error occured while attempting to use the common dialog: " & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description
   End If
  
End Sub

Private Sub Command1_Click() 'Event fires when the user clicks command one

            If LenB(Text1.Text) Then 'Conditional: Make sure that the media path is present before proceeding
                strFilePath = Text1.Text
            Else
               
               'Path missing so we notify the user
               MsgBox "Please provide a path for a video clip"
               
               'Exit the sub
               Exit Sub
               
            End If 'End conditional
            
            
            
            SetParent fra1.hWnd, sb1.hWnd 'We Set the parent of the object here (fra1-Video)
            
            'Position the mediaplayer control within the panel
            wmp1.URL = strFilePath 'Assign the file path to the url property of the mediaplayer control
            
            fra1.Left = sb1.Panels(1).Left - -nOffSet 'Set the left most position of the frame
            
            fra1.Top = nOffSet 'Set the top position of the frame
            
            fra1.Width = sb1.Panels(1).Width - nOffSet * 2 'Set the width of the frame
            
            fra1.Height = sb1.Height - nOffSet * 2 'Set the height of the frame
            
            wmp1.Top = fra1.Top  'Set the top of the MediaPlayer
            
            wmp1.Left = fra1.Left - nOffSet 'Set the left of the  MediaPlayer
            
            wmp1.Height = fra1.Height  'Set the height of the MediaPlayer
            
            wmp1.Width = fra1.Width  'Set the width of the MediaPlayer
            
            
            
            SetParent Check1.hWnd, sb1.hWnd  'We Set the parent of the object here (CheckBox)
            
            'Position the Checkbox control within the panel
            Check1.Left = sb1.Panels(2).Left + nOffSet 'Set the left position of the CheckBox
            
            Check1.Width = sb1.Panels(2).Width - nOffSet * 2 'Set the width of the CheckBox
            
            Check1.Top = nOffSet  'Set the Top of the CheckBox
            
            
            SetParent ProgressBar1.hWnd, sb1.hWnd  'We Set the parent of the object here (ProgressBar)
            
            'Position the progress bar
            ProgressBar1.Left = sb1.Panels(3).Left + nOffSet 'Set the left position of the ProgressBar
            
            ProgressBar1.Width = sb1.Panels(3).Width - nOffSet * 2 'Set the width of the ProgressBar
            
            ProgressBar1.Top = nOffSet  'Set the Top of the ProgressBar
            
            Timer1.Enabled = True 'Turn the timer on
            
            
            
            SetParent ListView1.hWnd, sb1.hWnd  'We Set the parent of the object here (ProgressBar)
            
            'position the Listbox
            ListView1.Left = sb1.Panels(4).Left + nOffSet 'Set the left position of the ProgressBar
            
            ListView1.Width = sb1.Panels(4).Width - nOffSet * 2 'Set the width of the ProgressBar
            
            ListView1.Top = nOffSet  'Set the Top of the ProgressBar
            
            ListView1.Height = sb1.Height - nOffSet * 2 'Set the height of the listview minus the offset times 2
            
            ListView1.ListItems.Add 1, , "Test1" 'Add an entry to the listview
            
            ListView1.ListItems.Add 2, , "Test2" 'Add an entry to the listview
            
            ListView1.ListItems.Add 3, , "Test3" 'Add an entry to the listview
            
            ListView1.ListItems.Add 4, , "Test4" 'Add an entry to the listview
            
            ListView1.ListItems.Add 5, , "Test5" 'Add an entry to the listview
            
            ListView1.Enabled = True 'Enable the list view
            
            SetParent Command2.hWnd, sb1.hWnd  'We Set the parent of the object here (Command button)
            
            'Position the progress bar
            Command2.Left = sb1.Panels(5).Left + nOffSet 'Set the left position of the Command button
            
            Command2.Width = sb1.Panels(5).Width - nOffSet * 2 'Set the width of the Command button
            
            Command2.Height = sb1.Height - nOffSet * 2 'set the height of the Command button
            
            Command2.Top = nOffSet  'Set the Top of the Command button
            
            Command2.Enabled = True 'Enable the control
            
            SetParent TabStrip1.hWnd, sb1.hWnd  'We Set the parent of the object here (Command button)
            
            'Position the progress bar
            TabStrip1.Left = sb1.Panels(6).Left + nOffSet 'Set the left position of the Command button
            
            sb1.Panels(6).Width = TabStrip1.Width + nOffSet * 2 ' 'Set the width of the Command button
            
            TabStrip1.Height = sb1.Height - nOffSet * 2 'set the height of the Command button
            
            TabStrip1.Top = nOffSet  'Set the Top of the Command button
            
            TabStrip1.Enabled = True
            
            
            SetParent Drive1.hWnd, sb1.hWnd  'We Set the parent of the object here (Command button)
            'Position the progress bar
            
            Drive1.Left = sb1.Panels(7).Left + nOffSet 'Set the left position of the Command button
            
            Drive1.Width = sb1.Panels(7).Width - nOffSet * 2 'Set the width of the Command button
            
            Drive1.Top = nOffSet  'Set the Top of the Command button
            
            Drive1.Enabled = True 'Enable the Drivelist
            


End Sub

Private Sub Form_Load() 'Event fires when the form loads


        nOffSet = 3 * Screen.TwipsPerPixelX ' Get the value that we should use to adjust control positions
        
       ' strFilePath = "Your file path" 'Set a default file here if you want

End Sub

Private Sub Timer1_Timer() 'Event fires continually when the Timer is enabled the frequency of which is determined by the interval property


        'Demonstrate that the ProgressBar still works
        If ProgressBar1.Value >= ProgressBar1.Max Then 'See if the max value has been reached
           
           ProgressBar1.Value = 0 'The max was reached so we reset the ProgressBar to zero
        
        Else
           
           ProgressBar1.Value = ProgressBar1.Value + 10 'Max was not yet reached so we increment the ProgressBar by a value 10
        
        End If 'End condition
        
        
    
End Sub
