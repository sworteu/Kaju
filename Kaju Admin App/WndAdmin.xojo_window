#tag DesktopWindow
Begin DesktopWindow WndAdmin
   Backdrop        =   0
   BackgroundColor =   &cFFFFFF00
   Composite       =   True
   DefaultLocation =   0
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   False
   HasMaximizeButton=   True
   HasMinimizeButton=   True
   HasTitleBar     =   True
   Height          =   680
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   1253638004
   MenuBarVisible  =   True
   MinimumHeight   =   680
   MinimumWidth    =   932
   Resizeable      =   True
   Title           =   "Untitled"
   Type            =   0
   Visible         =   True
   Width           =   932
   Begin DesktopListBox lbVersions
      AllowAutoDeactivate=   True
      AllowAutoHideScrollbars=   True
      AllowExpandableRows=   False
      AllowFocusRing  =   True
      AllowResizableColumns=   False
      AllowRowDragging=   False
      AllowRowReordering=   False
      Bold            =   False
      ColumnCount     =   1
      ColumnWidths    =   ""
      DefaultRowHeight=   -1
      DropIndicatorVisible=   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      GridLineStyle   =   0
      HasBorder       =   True
      HasHeader       =   True
      HasHorizontalScrollbar=   False
      HasVerticalScrollbar=   True
      HeadingIndex    =   -1
      Height          =   614
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   True
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      RequiresSelection=   False
      RowSelectionType=   0
      Scope           =   2
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   14
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   203
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
   End
   Begin DesktopButton btnNew
      AllowAutoDeactivate=   True
      Bold            =   True
      Cancel          =   False
      Caption         =   "+"
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   True
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      MacButtonStyle  =   0
      Scope           =   2
      TabIndex        =   12
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   640
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   20
   End
   Begin DesktopButton btnDelete
      AllowAutoDeactivate=   True
      Bold            =   True
      Cancel          =   False
      Caption         =   "-"
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   39
      LockBottom      =   True
      LockedInPosition=   True
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      MacButtonStyle  =   0
      Scope           =   2
      TabIndex        =   13
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   640
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   20
   End
   Begin TextFieldChanger fldVersion
      AllowAutoDeactivate=   True
      AllowFocusRing  =   True
      AllowSpellChecking=   False
      AllowTabs       =   False
      BackgroundColor =   &cFFFFFF
      Bold            =   False
      Enabled         =   False
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Format          =   ""
      HasBorder       =   True
      Height          =   22
      Hint            =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   380
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MaximumCharactersAllowed=   0
      Password        =   False
      ReadOnly        =   False
      Scope           =   2
      TabIndex        =   5
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextAlignment   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   52
      Transparent     =   False
      Underline       =   False
      ValidationMask  =   ""
      Visible         =   True
      Width           =   80
   End
   Begin DesktopLabel Label1
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   0
      InitialParent   =   ""
      Italic          =   False
      Left            =   251
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   2
      Selectable      =   False
      TabIndex        =   4
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Version:"
      TextAlignment   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   53
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   117
   End
   Begin DesktopTabPanel TabPanel1
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   509
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   251
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Panels          =   ""
      Scope           =   2
      SmallTabs       =   False
      TabDefinition   =   "Release Notes\rBinaries\r64-bit Binaries\rImage"
      TabIndex        =   10
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   119
      Transparent     =   False
      Underline       =   False
      Value           =   3
      Visible         =   True
      Width           =   653
      Begin TextAreaChanger fldReleaseNotes
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   True
         AllowStyledText =   True
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF
         Bold            =   False
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         HasHorizontalScrollbar=   False
         HasVerticalScrollbar=   True
         Height          =   177
         HideSelection   =   True
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   271
         LineHeight      =   0.0
         LineSpacing     =   1.0
         LockBottom      =   True
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Multiline       =   True
         ReadOnly        =   False
         Scope           =   2
         TabIndex        =   9
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   ""
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   194
         Transparent     =   False
         Underline       =   False
         UnicodeMode     =   0
         ValidationMask  =   ""
         Visible         =   True
         Width           =   613
      End
      Begin DesktopHTMLViewer hvReleaseNotesPreview
         AutoDeactivate  =   True
         Enabled         =   True
         Height          =   210
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Left            =   271
         LockBottom      =   True
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   False
         Renderer        =   1
         Scope           =   2
         TabIndex        =   12
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   408
         Visible         =   True
         Width           =   613
      End
      Begin DesktopLabel Label1
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   4
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   271
         LockBottom      =   True
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   False
         Multiline       =   False
         Scope           =   2
         Selectable      =   False
         TabIndex        =   10
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "Preview:"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   383
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   100
      End
      Begin DesktopLabel Label1
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   7
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   271
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   2
         Selectable      =   False
         TabIndex        =   0
         TabPanelIndex   =   1
         TabStop         =   True
         Text            =   "Release Notes (HTML):"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   171
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   142
      End
      Begin TextFieldChanger fldImageURL
         AllowAutoDeactivate=   True
         AllowFocusRing  =   True
         AllowSpellChecking=   False
         AllowTabs       =   False
         BackgroundColor =   &cFFFFFF
         Bold            =   False
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Format          =   ""
         HasBorder       =   True
         Height          =   22
         Hint            =   ""
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   350
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         MaximumCharactersAllowed=   0
         Password        =   False
         ReadOnly        =   False
         Scope           =   2
         TabIndex        =   4
         TabPanelIndex   =   4
         TabStop         =   True
         Text            =   ""
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   195
         Transparent     =   False
         Underline       =   False
         ValidationMask  =   ""
         Visible         =   True
         Width           =   534
      End
      Begin DesktopLabel Label1
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   8
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   271
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   2
         Selectable      =   False
         TabIndex        =   3
         TabPanelIndex   =   4
         TabStop         =   True
         Text            =   "URL:"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   195
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   67
      End
      Begin CheckBoxChanger cbImageUseTransparency
         AllowAutoDeactivate=   True
         Bold            =   False
         Caption         =   "Use Transparency"
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   350
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Scope           =   2
         TabIndex        =   0
         TabPanelIndex   =   4
         TabStop         =   True
         Tooltip         =   ""
         Top             =   162
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         VisualState     =   0
         Width           =   180
      End
      Begin DesktopHTMLViewer hvImagePreview
         AutoDeactivate  =   True
         Enabled         =   True
         Height          =   351
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Left            =   271
         LockBottom      =   True
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         Renderer        =   1
         Scope           =   2
         TabIndex        =   5
         TabPanelIndex   =   4
         TabStop         =   True
         Tooltip         =   ""
         Top             =   257
         Visible         =   True
         Width           =   613
      End
      Begin DesktopBevelButton btnStyle
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         BackgroundColor =   
         BevelStyle      =   0
         Bold            =   True
         ButtonStyle     =   0
         Caption         =   "B"
         CaptionAlignment=   0
         CaptionDelta    =   0
         CaptionPosition =   0
         Enabled         =   True
         FontName        =   ""
         FontSize        =   0.0
         FontUnit        =   0
         HasBackgroundColor=   False
         Height          =   22
         Icon            =   0
         IconAlignment   =   0
         IconDeltaX      =   0
         IconDeltaY      =   0
         Index           =   0
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   802
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   True
         MenuStyle       =   0
         PanelIndex      =   0
         Scope           =   2
         TabIndex        =   6
         TabPanelIndex   =   1
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   169
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         Width           =   26
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopBevelButton btnStyle
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         BackgroundColor =   
         BevelStyle      =   0
         Bold            =   False
         ButtonStyle     =   0
         Caption         =   "I"
         CaptionAlignment=   0
         CaptionDelta    =   0
         CaptionPosition =   0
         Enabled         =   True
         FontName        =   ""
         FontSize        =   0.0
         FontUnit        =   0
         HasBackgroundColor=   False
         Height          =   22
         Icon            =   0
         IconAlignment   =   0
         IconDeltaX      =   0
         IconDeltaY      =   0
         Index           =   1
         InitialParent   =   "TabPanel1"
         Italic          =   True
         Left            =   830
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   True
         MenuStyle       =   0
         PanelIndex      =   0
         Scope           =   2
         TabIndex        =   7
         TabPanelIndex   =   1
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   169
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         Width           =   26
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopBevelButton btnStyle
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         BackgroundColor =   
         BevelStyle      =   0
         Bold            =   False
         ButtonStyle     =   0
         Caption         =   "U"
         CaptionAlignment=   0
         CaptionDelta    =   0
         CaptionPosition =   0
         Enabled         =   True
         FontName        =   ""
         FontSize        =   0.0
         FontUnit        =   0
         HasBackgroundColor=   False
         Height          =   22
         Icon            =   0
         IconAlignment   =   0
         IconDeltaX      =   0
         IconDeltaY      =   0
         Index           =   2
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   858
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   True
         MenuStyle       =   0
         PanelIndex      =   0
         Scope           =   2
         TabIndex        =   8
         TabPanelIndex   =   1
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   169
         Transparent     =   False
         Underline       =   True
         Value           =   False
         Visible         =   True
         Width           =   26
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopBevelButton btnBreak
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         BackgroundColor =   
         BevelStyle      =   0
         Bold            =   False
         ButtonStyle     =   0
         Caption         =   "<br />"
         CaptionAlignment=   0
         CaptionDelta    =   0
         CaptionPosition =   0
         Enabled         =   True
         FontName        =   ""
         FontSize        =   0.0
         FontUnit        =   0
         HasBackgroundColor=   False
         Height          =   22
         Icon            =   0
         IconAlignment   =   0
         IconDeltaX      =   0
         IconDeltaY      =   0
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   601
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   True
         MenuStyle       =   0
         PanelIndex      =   0
         Scope           =   2
         TabIndex        =   1
         TabPanelIndex   =   1
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   169
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         Width           =   50
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopBevelButton btnStyle
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         BackgroundColor =   
         BevelStyle      =   0
         Bold            =   False
         ButtonStyle     =   0
         Caption         =   "p"
         CaptionAlignment=   0
         CaptionDelta    =   0
         CaptionPosition =   0
         Enabled         =   True
         FontName        =   ""
         FontSize        =   0.0
         FontUnit        =   0
         HasBackgroundColor=   False
         Height          =   22
         Icon            =   0
         IconAlignment   =   0
         IconDeltaX      =   0
         IconDeltaY      =   0
         Index           =   3
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   653
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   True
         MenuStyle       =   0
         PanelIndex      =   0
         Scope           =   2
         TabIndex        =   2
         TabPanelIndex   =   1
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   169
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         Width           =   26
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopBevelButton btnList
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         BackgroundColor =   
         BevelStyle      =   0
         Bold            =   False
         ButtonStyle     =   0
         Caption         =   "ul"
         CaptionAlignment=   0
         CaptionDelta    =   0
         CaptionPosition =   0
         Enabled         =   True
         FontName        =   ""
         FontSize        =   0.0
         FontUnit        =   0
         HasBackgroundColor=   False
         Height          =   22
         Icon            =   0
         IconAlignment   =   0
         IconDeltaX      =   0
         IconDeltaY      =   0
         Index           =   0
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   699
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   True
         MenuStyle       =   0
         PanelIndex      =   0
         Scope           =   2
         TabIndex        =   3
         TabPanelIndex   =   1
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   169
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         Width           =   26
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopBevelButton btnList
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         BackgroundColor =   
         BevelStyle      =   0
         Bold            =   False
         ButtonStyle     =   0
         Caption         =   "ol"
         CaptionAlignment=   0
         CaptionDelta    =   0
         CaptionPosition =   0
         Enabled         =   True
         FontName        =   ""
         FontSize        =   0.0
         FontUnit        =   0
         HasBackgroundColor=   False
         Height          =   22
         Icon            =   0
         IconAlignment   =   0
         IconDeltaX      =   0
         IconDeltaY      =   0
         Index           =   1
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   727
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   True
         MenuStyle       =   0
         PanelIndex      =   0
         Scope           =   2
         TabIndex        =   4
         TabPanelIndex   =   1
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   169
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         Width           =   26
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopBevelButton btnStyle
         Active          =   False
         AllowAutoDeactivate=   True
         AllowFocus      =   False
         AllowTabStop    =   True
         BackgroundColor =   
         BevelStyle      =   0
         Bold            =   False
         ButtonStyle     =   0
         Caption         =   "li"
         CaptionAlignment=   0
         CaptionDelta    =   0
         CaptionPosition =   0
         Enabled         =   True
         FontName        =   ""
         FontSize        =   0.0
         FontUnit        =   0
         HasBackgroundColor=   False
         Height          =   22
         Icon            =   0
         IconAlignment   =   0
         IconDeltaX      =   0
         IconDeltaY      =   0
         Index           =   6
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   755
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   True
         MenuStyle       =   0
         PanelIndex      =   0
         Scope           =   2
         TabIndex        =   5
         TabPanelIndex   =   1
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   169
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         Width           =   26
         _mIndex         =   0
         _mInitialParent =   ""
         _mName          =   ""
         _mPanelIndex    =   0
      End
      Begin DesktopLabel Label2
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   271
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         Multiline       =   False
         Scope           =   2
         Selectable      =   False
         TabIndex        =   0
         TabPanelIndex   =   2
         TabStop         =   True
         Text            =   "Use ‘$VERSION$’ in the URL to insert the version on export"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   589
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   613
      End
      Begin DesktopLabel Label3
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   271
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   True
         LockTop         =   True
         Multiline       =   False
         Scope           =   2
         Selectable      =   False
         TabIndex        =   0
         TabPanelIndex   =   3
         TabStop         =   True
         Text            =   "Use ‘$VERSION$’ in the URL to insert the version on export"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   589
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   613
      End
      Begin DesktopCheckBox cbPre2Preview
         AllowAutoDeactivate=   True
         Bold            =   False
         Caption         =   "Show Pre-2.0 Preview"
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   727
         LockBottom      =   True
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   False
         Scope           =   2
         TabIndex        =   11
         TabPanelIndex   =   1
         TabStop         =   True
         Tooltip         =   ""
         Top             =   382
         Transparent     =   False
         Underline       =   False
         Value           =   False
         Visible         =   True
         VisualState     =   0
         Width           =   157
      End
      Begin PopupMenuChanger pumImageScale
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   False
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   -2147483648
         InitialParent   =   "TabPanel1"
         InitialValue    =   "1\n2\n3"
         Italic          =   False
         Left            =   804
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   False
         LockRight       =   True
         LockTop         =   True
         Scope           =   2
         SelectedRowIndex=   0
         TabIndex        =   2
         TabPanelIndex   =   4
         TabStop         =   True
         Tooltip         =   ""
         Top             =   161
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   80
      End
      Begin DesktopLabel Label1
         AllowAutoDeactivate=   True
         Bold            =   False
         Enabled         =   True
         FontName        =   "System"
         FontSize        =   0.0
         FontUnit        =   0
         Height          =   20
         Index           =   10
         InitialParent   =   "TabPanel1"
         Italic          =   False
         Left            =   725
         LockBottom      =   False
         LockedInPosition=   False
         LockLeft        =   True
         LockRight       =   False
         LockTop         =   True
         Multiline       =   False
         Scope           =   2
         Selectable      =   False
         TabIndex        =   1
         TabPanelIndex   =   4
         TabStop         =   True
         Text            =   "Scale:"
         TextAlignment   =   0
         TextColor       =   &c00000000
         Tooltip         =   ""
         Top             =   162
         Transparent     =   False
         Underline       =   False
         Visible         =   True
         Width           =   67
      End
   End
   Begin TextFieldChanger fldAppName
      AllowAutoDeactivate=   True
      AllowFocusRing  =   True
      AllowSpellChecking=   False
      AllowTabs       =   False
      BackgroundColor =   &cFFFFFF
      Bold            =   False
      Enabled         =   False
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Format          =   ""
      HasBorder       =   True
      Height          =   22
      Hint            =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   380
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MaximumCharactersAllowed=   0
      Password        =   False
      ReadOnly        =   False
      Scope           =   2
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextAlignment   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   18
      Transparent     =   False
      Underline       =   False
      ValidationMask  =   ""
      Visible         =   True
      Width           =   229
   End
   Begin DesktopLabel Label1
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   3
      InitialParent   =   ""
      Italic          =   False
      Left            =   251
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   2
      Selectable      =   False
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "App Name:"
      TextAlignment   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   19
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   117
   End
   Begin CheckBoxChanger cbRequiresPayment
      AllowAutoDeactivate=   True
      Bold            =   False
      Caption         =   "Requires Payment"
      Enabled         =   False
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   693
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Scope           =   2
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   19
      Transparent     =   False
      Underline       =   False
      Value           =   False
      Visible         =   True
      VisualState     =   0
      Width           =   180
   End
   Begin DesktopHTMLViewer hvNewWindow
      AutoDeactivate  =   True
      Enabled         =   False
      Height          =   200
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   -371
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Renderer        =   0
      Scope           =   2
      TabIndex        =   11
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   156
      Visible         =   False
      Width           =   300
   End
   Begin Timer tmrUpdateReleaseNotesPreview
      Enabled         =   True
      Index           =   -2147483648
      InitialParent   =   ""
      LockedInPosition=   False
      Mode            =   0
      Period          =   750
      Scope           =   2
      TabPanelIndex   =   0
   End
   Begin DesktopButton btnCopyPublicKey
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "Copy RSA Public Key"
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   251
      LockBottom      =   True
      LockedInPosition=   True
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      MacButtonStyle  =   0
      Scope           =   2
      TabIndex        =   15
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   640
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   186
   End
   Begin Timer tmrUpdateImagePreview
      Enabled         =   True
      Index           =   -2147483648
      InitialParent   =   ""
      LockedInPosition=   False
      Mode            =   0
      Period          =   750
      Scope           =   2
      TabPanelIndex   =   0
   End
   Begin DesktopLabel lblPlatform
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   58
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   577
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   True
      Scope           =   2
      Selectable      =   False
      TabIndex        =   7
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "None"
      TextAlignment   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   55
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   327
   End
   Begin DesktopLabel Label1
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   9
      InitialParent   =   ""
      Italic          =   False
      Left            =   492
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   2
      Selectable      =   False
      TabIndex        =   6
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Platforms:"
      TextAlignment   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   53
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin DesktopButton btnExport
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "Export..."
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   819
      LockBottom      =   True
      LockedInPosition=   True
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      MacButtonStyle  =   0
      Scope           =   2
      TabIndex        =   17
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   639
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   85
   End
   Begin DesktopButton btnDuplicate
      AllowAutoDeactivate=   True
      Bold            =   False
      Cancel          =   False
      Caption         =   "Dupe"
      Default         =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   158
      LockBottom      =   True
      LockedInPosition=   True
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      MacButtonStyle  =   0
      Scope           =   2
      TabIndex        =   14
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   641
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   65
   End
   Begin DesktopPopupMenu btnPreview
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   22
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   ""
      Italic          =   False
      Left            =   684
      LockBottom      =   True
      LockedInPosition=   True
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Scope           =   2
      SelectedRowIndex=   -1
      TabIndex        =   16
      TabPanelIndex   =   0
      TabStop         =   True
      Tooltip         =   ""
      Top             =   639
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   113
   End
   Begin TextFieldChanger fldMinRequiredVersion
      AllowAutoDeactivate=   True
      AllowFocusRing  =   True
      AllowSpellChecking=   False
      AllowTabs       =   False
      BackgroundColor =   &cFFFFFF
      Bold            =   False
      Enabled         =   False
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Format          =   ""
      HasBorder       =   True
      Height          =   22
      Hint            =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   380
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      MaximumCharactersAllowed=   0
      Password        =   False
      ReadOnly        =   False
      Scope           =   2
      TabIndex        =   9
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextAlignment   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   85
      Transparent     =   False
      Underline       =   False
      ValidationMask  =   ""
      Visible         =   True
      Width           =   80
   End
   Begin DesktopLabel Label1
      AllowAutoDeactivate=   True
      Bold            =   False
      Enabled         =   True
      FontName        =   "System"
      FontSize        =   0.0
      FontUnit        =   0
      Height          =   20
      Index           =   2
      InitialParent   =   ""
      Italic          =   False
      Left            =   251
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   2
      Selectable      =   False
      TabIndex        =   8
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Min. Req. Version:"
      TextAlignment   =   0
      TextColor       =   &c00000000
      Tooltip         =   ""
      Top             =   86
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   117
   End
   Begin Kaju.UpdateInformation objReleaseNotesProcessor
      AppName         =   ""
      DisplayReleaseNotes=   ""
      Image           =   0
      ImageScale      =   1
      ImageURL        =   ""
      Index           =   -2147483648
      IsValid         =   False
      LockedInPosition=   False
      MinimumRequiredVersion=   ""
      ReleaseNotes    =   ""
      ReleaseNotesFromURL=   ""
      RequiresPayment =   False
      Scope           =   2
      StageCode       =   0
      TabPanelIndex   =   0
      UseTransparency =   True
      Version         =   ""
      VersionAsDouble =   0.0
   End
   Begin Timer tmrAdjustControls
      Enabled         =   True
      Index           =   -2147483648
      LockedInPosition=   False
      Mode            =   0
      Period          =   20
      Scope           =   2
      TabPanelIndex   =   0
   End
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Event
		Sub Activated()
		  UpdateWindowTitle
		End Sub
	#tag EndEvent

	#tag Event
		Function CancelClosing(appQuitting As Boolean) As Boolean
		  dim r as Boolean = false
		  
		  if self.Changed then
		    dim dlg as new MessageDialog
		    dlg.Message = "This document has been modified. Save before closing?"
		    dlg.ActionButton.Caption = "&Save"
		    dlg.ActionButton.Visible = true
		    dlg.CancelButton.Visible = true
		    dlg.AlternateActionButton.Caption = "&Don't Save"
		    dlg.AlternateActionButton.Visible = true
		    
		    dim btn as MessageDialogButton = dlg.ShowModal( self )
		    if btn is dlg.CancelButton then
		      r = true
		    elseif btn = dlg.AlternateActionButton then
		      r = false
		    else
		      r = not( DoSave() )
		    end if
		  end if
		  
		  return r
		  
		  #pragma unused appQuitting
		End Function
	#tag EndEvent

	#tag Event
		Sub Closing()
		  //
		  // Tear down the BinaryInformationControls
		  //
		  
		  for each bc as BinaryInformationControl in BinaryControls
		    RemoveHandler bc.CheckedChanged, AddressOf HandleBinaryInformationControlCheckedChanged
		  next
		  
		  redim BinaryControls( -1 )
		  
		End Sub
	#tag EndEvent

	#tag Event
		Sub Opening()
		  //
		  // Create binary controls
		  //
		  
		  dim binaryNames() as string = Kaju.UpdateInformation.BinaryNames
		  dim gap as integer = 10
		  dim useTop as integer = 50
		  dim useTop64 as integer = useTop
		  dim useLeft as integer = 15
		  
		  dim tpIndex as integer = 1
		  dim tpIndex64 as integer = 2
		  
		  for i as integer = 0 to binaryNames.Ubound
		    dim binaryName as string = binaryNames( i )
		    
		    dim cc as new BinaryInformationControl
		    cc.BinaryName = binaryName
		    cc.LockLeft = true
		    cc.LockRight = true
		    cc.LockTop = true
		    cc.LockBottom = false
		    'cc.Transparent = false
		    
		    if binaryName.InStr( "64" ) <> 0 then
		      cc.EmbedWithinPanel TabPanel1, tpIndex64, useLeft, useTop64
		      useTop64 = useTop64 + cc.Height + gap
		    else
		      cc.EmbedWithinPanel TabPanel1, tpIndex, useLeft, useTop
		      useTop = useTop + cc.Height + gap
		    end if
		    
		    AddHandler cc.CheckedChanged, AddressOf HandleBinaryInformationControlCheckedChanged
		    
		    BinaryControls.Append cc
		  next
		  
		  AdjustControls()
		End Sub
	#tag EndEvent


	#tag MenuHandler
		Function FileClose() As Boolean Handles FileClose.Action
		  self.Close
		  Return True
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function FileSave() As Boolean Handles FileSave.Action
		  call DoSave()
		  Return True
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function FileSaveAs() As Boolean Handles FileSaveAs.Action
		  call DoSaveAs()
		  Return True
		  
		End Function
	#tag EndMenuHandler


	#tag Method, Flags = &h21
		Private Sub AdjustControls(immediate As Boolean = False)
		  if not immediate or self.Loading then
		    tmrAdjustControls.Mode = Timer.ModeSingle
		    return
		  end if
		  
		  dim trueValue as boolean = lbVersions.SelectedRowIndex <> -1
		  
		  dim lastIndex as integer = ControlCount - 1
		  for i as integer = 0 to lastIndex
		    
		    Dim o As Object = Self.ControlAt( i )
		    
		    If Not ( o IsA DesktopControl ) Then
		      Continue
		    End If
		    
		    Var c As DesktopControl = DesktopControl( o )
		    
		    dim doIt as boolean
		    if ControlDataField( c ) <> "" then
		      doIt = true
		    elseif c IsA btnStyle then
		      doIt = true
		    elseif c IsA btnList then
		      doIt = true
		    end if
		    
		    if doIt then
		      DesktopUIControl( c ).Enabled = trueValue
		    end if
		  next i
		  
		  //
		  // Adjust the outliers that aren't directly set
		  //
		  
		  redim Platforms( -1 )
		  for each bc as BinaryInformationControl in BinaryControls
		    bc.Enabled = trueValue
		    bc.AdjustControls
		    
		    if bc.Enabled and bc.IsChecked then
		      Platforms.Append bc.BinaryName
		    end if
		  next
		  
		  btnDelete.Enabled = trueValue
		  btnDuplicate.Enabled = trueValue
		  
		  btnBreak.Enabled = trueValue
		  
		  //
		  // Set the platform summary
		  //
		  
		  if not trueValue then
		    lblPlatform.Text = ""
		  end if
		  
		  if Platforms.Ubound = -1 then
		    lblPlatform.Italic = true
		    lblPlatform.Text = "No platforms selected"
		  else
		    lblPlatform.Italic = false
		    lblPlatform.Text = join( Platforms, ", " )
		  end if
		  
		  UpdateWindowTitle
		  
		  //
		  // Do stuff for the first run here
		  //
		  if IsFirstRun then
		    lbVersions.SetFocus
		    
		    IsFirstRun = false
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ApplyStyle(tag As String, insertEOL As Boolean = False)
		  dim eol as string = EndOfLine
		  eol = if( insertEOL, eol, "" )
		  
		  tag = tag.Lowercase
		  
		  dim openTag as string = "<" + tag + ">" + eol
		  dim closeTag as string = eol + "</" + tag + ">"
		  
		  dim s as string = fldReleaseNotes.SelectedText
		  dim selStart as integer = fldReleaseNotes.SelectionStart
		  
		  fldReleaseNotes.SelectedText = openTag + s + closeTag
		  if s = "" then // No selection so position the cursor
		    fldReleaseNotes.SelectionStart = selStart + openTag.Len
		  end if
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ClearFields()
		  dim savedDirty as boolean = self.Changed
		  
		  self.Loading = true
		  
		  dim lastIndex as integer = ControlCount - 1
		  For i As Integer = 0 To lastIndex
		    
		    Dim o As Object = Self.ControlAt( i )
		    
		    If Not ( o IsA DesktopControl ) Then
		      Continue
		    End If
		    
		    Var c As DesktopControl = DesktopControl( o )
		    
		    select case c
		    Case IsA DesktopTextControl
		      
		    Case IsA DesktopCheckBox
		      
		    Case IsA DesktopPopupMenu
		      
		    else
		      continue for i
		    end
		    
		    ControlValue( c ) = nil
		  next i
		  
		  for each bc as BinaryInformationControl in BinaryControls
		    bc.Clear
		  next
		  
		  hvReleaseNotesPreview.LoadPage( kNoDataHTML, nil )
		  hvImagePreview.LoadPage( kNoDataHTML, nil )
		  
		  self.Loading = false
		  
		  AdjustControls()
		  self.Changed = savedDirty
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ControlDataField(c As DesktopControl) As String
		  // Returns the matching Kaju field name for each editable control.
		  
		  Select Case c
		  Case fldVersion
		    Return KajuFile.kVersionName
		  Case fldReleaseNotes
		    Return "ReleaseNotes"
		  Case fldImageURL
		    Return "ImageURL"
		  Case cbImageUseTransparency
		    Return "UseTransparency"
		  Case pumImageScale
		    Return "ImageScale"
		  Case fldAppName
		    Return "AppName"
		  Case cbRequiresPayment
		    Return "RequiresPayment"
		  Case fldMinRequiredVersion
		    Return "MinimumRequiredVersion"
		  Else
		    Return ""
		  End Select
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ControlValue(c As DesktopControl) As Variant
		  dim r as Variant
		  
		  select case c
		  case IsA DesktopTextControl
		    dim fld as DesktopTextControl = DesktopTextControl( c )
		    r = fld.Text
		    
		  case IsA DesktopCheckBox
		    dim cb as DesktopCheckBox = DesktopCheckBox( c )
		    r = cb.Value
		    
		  case IsA DesktopPopupMenu
		    dim pum as DesktopPopupMenu = DesktopPopupMenu( c )
		    if pum.SelectedRowIndex <> -1 then
		      r = pum.RowTextAt( pum.SelectedRowIndex )
		    end if
		    
		  end select
		  
		  return r
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ControlValue(c As DesktopControl, Assigns value As Variant)
		  
		  select case c
		    
		  case IsA DesktopTextControl
		    dim fld as DesktopTextControl = DesktopTextControl( c )
		    dim sValue as string = value.StringValue.DefineEncoding( Encodings.UTF8 )
		    fld.Text = sValue
		    
		  case IsA DesktopCheckBox
		    dim cb as DesktopCheckBox = DesktopCheckBox( c )
		    cb.Value = value.BooleanValue
		    
		  case IsA DesktopPopupMenu
		    Dim pum As DesktopPopupMenu = DesktopPopupMenu( c )
		    
		    if value.IsNull then
		      pum.SelectedRowIndex = 0
		      
		    else
		      for i as integer = pum.RowCount - 1 downto 0
		        if pum.RowTextAt( i ) = value.StringValue then
		          pum.SelectedRowIndex = i
		          exit
		        end if
		      next
		      
		    end if
		    
		  else
		    raise new KeyNotFoundException
		    
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DeleteVersion()
		  dim curIndex as integer = lbVersions.SelectedRowIndex
		  if curIndex = -1 then
		    return
		  end if
		  
		  StoreFieldsToVersionRow
		  
		  dim version as Kaju.UpdateInformation = lbVersions.RowTagAt( curIndex )
		  
		  dim dlg as new MessageDialog
		  dlg.Message = "Really delete version " + version.Version + "?"
		  dlg.Explanation = "This cannot be undone."
		  dlg.ActionButton.Caption = "&Delete"
		  dlg.CancelButton.Visible = true
		  dim btn as MessageDialogButton = dlg.ShowModal( self )
		  if btn is dlg.CancelButton then
		    return
		  end if
		  
		  lbVersions.SelectedRowIndex = -1
		  lbVersions.RemoveRowAt( curIndex )
		  
		  MyKajuFile.KajuData.Remove MyKajuFile.KajuData.IndexOf( version )
		  
		  if curIndex > 0 then
		    curIndex = curIndex - 1
		  end if
		  
		  if curIndex < lbVersions.RowCount then
		    lbVersions.SelectedRowIndex = curIndex
		  end if
		  
		  Changed = true
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DoSave(f As FolderItem = Nil) As Boolean
		  StoreFieldsToVersionRow()
		  
		  if f is nil then
		    f = Document
		  end if
		  
		  if f is nil then
		    return DoSaveAs()
		  end if
		  
		  dim savedContentsChanged as boolean = self.Changed
		  
		  dim r as boolean
		  MyKajuFile.SaveTo( f )
		  r = true
		  Document = f
		  Changed = false
		  
		  Exception err as KajuException
		    self.Changed = savedContentsChanged
		    ShowValidationError "Save failed!", err
		    
		  Exception err as IOException
		    self.Changed = savedContentsChanged
		    MsgBox "Save failed!"
		    
		  Finally
		    UpdateWindowTitle
		    return r
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function DoSaveAs() As Boolean
		  dim ext as string = FileTypes1.KajuDocument.Extensions
		  
		  dim dlg as new SaveAsDialog
		  dlg.PromptText = "Save the Kaju document as:"
		  dlg.Filter = FileTypes1.KajuDocument
		  dlg.SuggestedFileName = "New Version Line" + ext
		  
		  dim f as FolderItem = dlg.ShowModal( self )
		  
		  dim r as boolean
		  if f is nil then
		    r = false
		    
		  elseif f.Directory then
		    MsgBox "You chose a folder."
		    r = false
		    
		  else
		    //
		    // Adjust the name if needed
		    //
		    if not f.Exists then
		      if f.Name.Right( ext.Len ) <> ext then
		        f.Name = f.Name + ext
		      end if
		    end if
		    
		    r = DoSave( f )
		  end if
		  
		  UpdateWindowTitle
		  return r
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DuplicateVersion()
		  dim curIndex as integer = lbVersions.SelectedRowIndex
		  if curIndex = -1 then
		    return
		  end if
		  
		  StoreFieldsToVersionRow()
		  
		  dim tag as Variant = lbVersions.RowTagAt( curIndex )
		  dim infoCopy as new Kaju.UpdateInformation( Kaju.UpdateInformation( tag ) )
		  MyKajuFile.KajuData.Append infoCopy
		  
		  dim listing as string = lbVersions.CellTextAt( curIndex, 0 )
		  dim newIndex as integer = curIndex + 1
		  lbVersions.AddRowAt( newIndex, listing )
		  lbVersions.RowTagAt( newIndex ) = infoCopy
		  lbVersions.SelectedRowIndex = newIndex
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub HandleBinaryInformationControlCheckedChanged(sender As BinaryInformationControl)
		  #pragma unused sender
		  
		  AdjustControls()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function IsDataValid() As Boolean
		  StoreFieldsToVersionRow()
		  
		  dim r as boolean = true // Assume it's fine
		  
		  if lbVersions.RowCount = 0 then
		    return r
		  end if
		  
		  dim msg as string
		  dim lastRow as integer = lbVersions.RowCount - 1
		  for row as integer = 0 to lastRow
		    dim u as Kaju.UpdateInformation = lbVersions.RowTagAt( row )
		    if not u.IsValid then
		      r = false
		      msg = u.InvalidReason
		      lbVersions.SelectedRowIndex = row
		      exit for row
		    end if
		  next
		  
		  if not r then
		    MsgBox msg
		  end if
		  
		  return r
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MyKajuFile() As KajuFile
		  if mMyKajuFile is nil then
		    mMyKajuFile = new KajuFile
		  end if
		  
		  return mMyKajuFile
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub NewVersion()
		  StoreFieldsToVersionRow()
		  
		  lbVersions.AddRow "1.0.0d1"
		  dim version as new Kaju.UpdateInformation
		  version.Version = lbVersions.CellTextAt( lbVersions.LastRowIndex, 0 )
		  
		  dim prevIndex as integer = LastVersionRow
		  if prevIndex <> -1 and prevIndex < lbVersions.RowCount then
		    dim prevItem as Kaju.UpdateInformation = lbVersions.RowTagAt( prevIndex )
		    if prevItem <> nil then
		      version.AppName = prevItem.AppName
		    end if
		  end if
		  
		  MyKajuFile.KajuData.Append version
		  
		  lbVersions.RowTagAt( lbVersions.LastRowIndex ) = version
		  lbVersions.SelectedRowIndex = lbVersions.LastRowIndex
		  self.Changed = true
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub OpenDocument(f As FolderItem)
		  if f is nil or not f.Exists or f.Directory then
		    raise new NilObjectException
		  end if
		  MyKajuFile.Load f
		  
		  dim versions() as Kaju.UpdateInformation = MyKajuFile.KajuData
		  for i as integer = 0 to versions.Ubound
		    dim version as Kaju.UpdateInformation = versions( i )
		    lbVersions.AddRow version.Version
		    lbVersions.RowTagAt( lbVersions.LastRowIndex ) = version
		  next i
		  
		  if lbVersions.RowCount <> 0 then
		    lbVersions.SelectedRowIndex = 0
		  end if
		  
		  Document = f
		  self.Changed = false
		  
		  AdjustControls
		  
		  Exception err as RuntimeException
		    select case err
		    case IsA EndException, IsA ThreadEndException
		      //
		      // Pass it on
		      //
		      raise err
		      
		    case else
		      MsgBox "Could not open document."
		      
		      self.Close
		      return
		    end select
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub ShowValidationError(msg As String, err As RuntimeException)
		  dim mdlg as new MessageDialog
		  mdlg.Message = msg
		  mdlg.Explanation = err.Message
		  mdlg.CancelButton.Visible = false
		  
		  call mdlg.ShowModal self
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub StoreFieldsToVersionRow()
		  if LastVersionRow = -1 or LastVersionRow >= lbVersions.RowCount then
		    return
		  end if
		  
		  dim version as Kaju.UpdateInformation = lbVersions.RowTagAt( LastVersionRow )
		  
		  //
		  // Gather the textfield data first
		  //
		  
		  Dim lastIndex As Integer = ControlCount - 1
		  
		  For i As Integer = 0 To lastIndex
		    
		    Dim o As Object = Self.ControlAt( i )
		    
		    If Not ( o IsA DesktopControl ) Then
		      Continue
		    End If
		    
		    Var c As DesktopControl = DesktopControl( o )
		    
		    Dim fieldName As String = ControlDataField( c )
		    
		    if fieldName <> "" then
		      dim value as Variant = ControlValue( c )
		      version.SetByName( fieldName ) = value
		    end if
		    
		  next
		  
		  //
		  // Binaries
		  //
		  
		  for each bc as BinaryInformationControl in BinaryControls
		    
		    dim thisKey as string = bc.BinaryName
		    
		    if bc.IsChecked then
		      dim binary as Kaju.BinaryInformation = bc.GetBinaryInformation
		      version.Binaries.Value( thisKey ) = binary
		    else
		      version.RemoveBinary thisKey
		    end if
		    
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub UpdateWindowTitle()
		  dim f as FolderItem = Document
		  if f is nil then
		    self.Title = "Untitled"
		  else
		    dim t as string = f.Name
		    dim ext as string = FileTypes1.KajuDocument.Extensions
		    
		    if t.Right( ext.Len ) = ext then
		      t = t.Left( t.Len - ext.Len )
		    end if
		    
		    self.Title = t
		  end if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub VersionToFields(version As Kaju.UpdateInformation)
		  //
		  // Handle the named controls first
		  //
		  
		  dim savedDirty as boolean = self.Changed
		  
		  ClearFields()
		  
		  self.Loading = true
		  
		  dim lastIndex as integer = ControlCount - 1
		  for i as integer = 0 to lastIndex
		    
		    Dim o As Object = Self.ControlAt( i )
		    
		    If Not ( o IsA DesktopControl ) Then
		      Continue
		    End If
		    
		    Var c As DesktopControl = DesktopControl( o )
		    
		    dim fieldName as string = ControlDataField( c )
		    if fieldName <> "" then
		      try
		        ControlValue( c ) = version.GetByName( fieldName )
		      catch err as KeyNotFoundException
		        ControlValue( c ) = nil
		      end try
		    end if
		    
		  next
		  
		  //
		  // Binaries
		  //
		  
		  for each bc as BinaryInformationControl in BinaryControls
		    dim thisKey as string = bc.BinaryName
		    dim binary as Kaju.BinaryInformation = version.Binaries.Lookup( thisKey, nil )
		    bc.Load binary, version.Version
		  next
		  
		  self.Loading = false
		  
		  AdjustControls()
		  self.Changed = savedDirty
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private BinaryControls() As BinaryInformationControl
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  if mDocument is nil then
			    return nil
			  else
			    return mDocument.Resolve
			  end if
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mDocument = value
			End Set
		#tag EndSetter
		Document As FolderItem
	#tag EndComputedProperty

	#tag Property, Flags = &h21
		Private IsFirstRun As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h21
		Private LastVersionRow As Integer = -1
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mLoading > 0
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  if value then
			    mLoading = mLoading + 1
			  else
			    mLoading = mLoading - 1
			    if mLoading < 0 then
			      mLoading = 0
			    end if
			  end if
			  
			End Set
		#tag EndSetter
		Loading As Boolean
	#tag EndComputedProperty

	#tag Property, Flags = &h21
		Private mDocument As FolderItemAlias
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mLoading As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mMyKajuFile As KajuFile
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Platforms() As String
	#tag EndProperty


	#tag Constant, Name = kNoDataHTML, Type = String, Dynamic = False, Default = \"<!-- No data --><BR />", Scope = Private
	#tag EndConstant


#tag EndWindowCode

#tag Events lbVersions
	#tag Event
		Sub SelectionChanged()
		  StoreFieldsToVersionRow
		  
		  Dim row As Integer = Me.SelectedRowIndex
		  
		  if row <> -1 and me.RowTagAt( row ) <> nil then
		    VersionToFields( me.RowTagAt( row ) )
		  else
		    ClearFields
		  end if
		  
		  LastVersionRow = row
		  
		  AdjustControls()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnNew
	#tag Event
		Sub Pressed()
		  NewVersion
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnDelete
	#tag Event
		Sub Pressed()
		  DeleteVersion
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events fldVersion
	#tag Event
		Sub TextChange()
		  If lbVersions.SelectedRowIndex = -1 Then
		    Return
		  End If
		  
		  Dim v As String = Me.Text.Trim
		  If v = "" Then
		    v = "(no version)"
		  End If
		  
		  lbVersions.CellTextAt( lbVersions.SelectedRowIndex, 0 ) = v
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events fldReleaseNotes
	#tag Event
		Sub TextChange()
		  objReleaseNotesProcessor.ReleaseNotes = Me.Text
		  tmrUpdateReleaseNotesPreview.Mode = Timer.ModeSingle
		  tmrUpdateReleaseNotesPreview.Reset
		  
		  #If TargetWindows Then
		    hvReleaseNotesPreview.LoadPage( kNoDataHTML, Nil )
		    
		    Dim releaseNotes As String = ControlValue( fldReleaseNotes ).StringValue
		    hvReleaseNotesPreview.LoadPage releaseNotes, Nil
		    Me.SetFocus
		  #EndIf
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events hvReleaseNotesPreview
	#tag Event
		Function NewWindow(url as String) As DesktopHTMLViewer
		  #Pragma Unused url 
		  
		  return hvNewWindow
		End Function
	#tag EndEvent
	#tag Event
		Function CancelLoad(URL as String) As Boolean
		  return ( not self.Loading )
		  
		  #pragma unused URL
		End Function
	#tag EndEvent
	#tag Event
		Sub DocumentComplete(url as String)
		  #pragma unused URL
		  
		  self.Loading = false
		End Sub
	#tag EndEvent
	#tag Event
		Sub Error(error As RuntimeException)
		  #pragma unused Error
		  
		  break
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events fldImageURL
	#tag Event
		Sub TextChange()
		  tmrUpdateImagePreview.Mode = Timer.ModeSingle
		  tmrUpdateImagePreview.Reset
		  
		  #If TargetWindows Then
		    hvImagePreview.LoadPage kNoDataHTML, Nil
		  #endif
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events hvImagePreview
	#tag Event
		Function NewWindow(url as String) As DesktopHTMLViewer
		  #Pragma Unused url
		  
		  return hvNewWindow
		End Function
	#tag EndEvent
	#tag Event
		Function CancelLoad(URL as String) As Boolean
		  return ( not self.Loading )
		  
		  #pragma unused URL
		End Function
	#tag EndEvent
	#tag Event
		Sub DocumentComplete(url as String)
		  #pragma unused URL
		  
		  self.Loading = false
		End Sub
	#tag EndEvent
	#tag Event
		Sub Error(error As RuntimeException)
		  #Pragma unused Error
		  
		  break
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnStyle
	#tag Event
		Sub Pressed(index as Integer)
		  dim tag as string = me.Caption
		  ApplyStyle( tag )
		  
		  #pragma unused index
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnBreak
	#tag Event
		Sub Pressed()
		  fldReleaseNotes.SelectedText = "<br />" + EndOfLine
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnList
	#tag Event
		Sub Pressed(index as Integer)
		  dim tag as string = me.Caption
		  ApplyStyle( tag, true )
		  
		  #pragma unused index
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events cbPre2Preview
	#tag Event
		Sub ValueChanged()
		  tmrUpdateReleaseNotesPreview.Mode = Timer.ModeSingle
		  tmrUpdateReleaseNotesPreview.Reset
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events hvNewWindow
	#tag Event
		Function CancelLoad(URL as String) As Boolean
		  ShowURL( URL )
		  return true
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events tmrUpdateReleaseNotesPreview
	#tag Event
		Sub Action()
		  self.Loading = true
		  
		  dim releaseNotes as string
		  if cbPre2Preview.Value then
		    releaseNotes = ControlValue( fldReleaseNotes ).StringValue
		  else
		    releaseNotes = objReleaseNotesProcessor.DisplayReleaseNotes
		  end if
		  
		  dim saveControl as DesktopUIControl = self.Focus
		  hvReleaseNotesPreview.LoadPage( releaseNotes, nil )
		  self.Focus = saveControl
		  //
		  // The htmlViewer will set Loading to false
		  //
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnCopyPublicKey
	#tag Event
		Sub Pressed()
		  
		  dim c as new Clipboard
		  c.Text = MyKajuFile.PublicKey
		  c.Close
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events tmrUpdateImagePreview
	#tag Event
		Sub Action()
		  self.Loading = true
		  
		  hvImagePreview.LoadURL( fldImageURL.Text )
		  
		  //
		  // The htmlViewer will set Loading to false
		  //
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnExport
	#tag Event
		Sub Pressed()
		  if not IsDataValid() then
		    return
		  end if
		  
		  dim dlg as new SaveFileDialog
		  dlg.PromptText = "Export the file that will be served to your app through your web site:"
		  dlg.ActionButtonCaption = "Export"
		  dlg.SuggestedFileName = MyKajuFile.ExportFilename
		  
		  dim f as FolderItem = dlg.ShowModal( self )
		  if f is nil then
		    return
		  end if
		  
		  MyKajuFile.ExportTo( f )
		  
		  Exception err as KajuException
		    ShowValidationError "Could not export data.", err
		    
		  Exception err As RuntimeException
		    if err isa EndException or err isa ThreadEndException then
		      raise err
		    end if
		    
		    MsgBox "Could not export data."
		    
		    
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnDuplicate
	#tag Event
		Sub Pressed()
		  DuplicateVersion
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnPreview
	#tag Event
		Sub Opening()
		  me.AddRow "Development"
		  me.AddRow "Alpha"
		  me.AddRow "Beta"
		  me.AddRow "Final"
		  
		  me.SelectedRowIndex = -1
		End Sub
	#tag EndEvent
	#tag Event
		Sub SelectionChanged(item As DesktopMenuItem)
		  #Pragma Unused item
		  
		  If Me.SelectedRowIndex = -1 Then
		    Return
		  End If
		  
		  If Not IsDataValid Then
		    Return
		  End If
		  
		  Dim uc As New Kaju.UpdateChecker( App.PrefFolder )
		  uc.AllowedStage = Me.SelectedRowIndex
		  Dim s As String = MyKajuFile.DataToJSON.ToString
		  uc.TestUpdate( s )
		  
		  me.SelectedRowIndex = -1
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events objReleaseNotesProcessor
	#tag Event
		Sub ReleaseNotesReceived()
		  tmrUpdateReleaseNotesPreview.Mode = Timer.ModeSingle
		  tmrUpdateReleaseNotesPreview.Reset
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events tmrAdjustControls
	#tag Event
		Sub Action()
		  AdjustControls true
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="HasTitleBar"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumWidth"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumHeight"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumWidth"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumHeight"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Type"
		Visible=true
		Group="Frame"
		InitialValue="0"
		Type="Types"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasCloseButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMaximizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMinimizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasFullScreenButton"
		Visible=true
		Group="Frame"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="DefaultLocation"
		Visible=true
		Group="Behavior"
		InitialValue="0"
		Type="Locations"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="&hFFFFFF"
		Type="ColorGroup"
		EditorType="ColorGroup"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Appearance"
		InitialValue=""
		Type="Picture"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Visible=true
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Visible=false
		Group="Appearance"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Position"
		InitialValue="400"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Loading"
		Visible=false
		Group="Behavior"
		InitialValue=""
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Visible=true
		Group="Appearance"
		InitialValue="0"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Appearance"
		InitialValue=""
		Type="DesktopMenuBar"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=false
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Appearance"
		InitialValue="Untitled"
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Appearance"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Position"
		InitialValue="600"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
#tag EndViewBehavior
