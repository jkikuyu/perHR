Attribute VB_Name = "HR"
    Global NewClndEntry As Boolean
    Global dHoliday As Boolean
    Global dNoWork As Boolean
    
    Global uPass As String
    Global uUser As String
    Global uCo As String
    
    Global Operation_Year As Long
    Global PMember As Boolean
    Global PSocial As Boolean
    Global PEHist As Boolean
    Global PDHist As Boolean
    Global PDevt As Boolean
    Global PTrans As Boolean
    Global PLoans As Boolean
    Global PEmp As Boolean
    Global PPayroll As Boolean
    Global PBanking As Boolean
    Global POT As Boolean
    Global PRpt As Boolean
    Global PDepn As Boolean
    Global PLeave As Boolean
    Global PLeave2 As Boolean
    Global PLeave3 As Boolean
    Global PLeaveEmp As String
    Global PMed As Boolean
    Global PSch As Boolean
    Global PTr As Boolean
    Global PAbs As Boolean
    Global pApp As Boolean
    '*****************************************************************************
    'The following codes are used in registering and editing employee job details
    Global Dept_Code
    Global Sect_code
    Global Design_code
    Global Contr_code
    Global Family_code
    Global SexMar As String
    '*****************************************************************************
    Global Tp As Long
    Global Trsb As Boolean
    Global Trs As Long
    Global NewRpt As Boolean
    Global SelRpt
    Global WageSelect As String
    Global EmpSelect As String
    Global SelF As String
    Global coBank As Boolean
    Global coBankid As String
    
    Global eBatch As Boolean
    Global Batch As Boolean
    Global v_Batch As Boolean
    Global valid_key As Boolean
    
    
    Global HistStaff_code
    Global HistLoan_Code
    Global wagetype As Boolean
    Global DataSelection As Integer
    Global EmpSelected As Boolean
    Global EmpStaffCode
    Global WageEdit As Boolean
    Global StopEditWage As Boolean
    
    Global datExtra As Recordset
    Global ThisMonth
    Global ThisYear
    Global pass As String
    Global FirstCo As Boolean
    Global SlipOrder As Integer
    Global Error_Found As Boolean
    Global loading As Boolean
    Global MinTax As Boolean
    Global paramcode
    Global HOSPValue As Currency
    Global OOIValue As Currency
    Global HousingValue As Currency
    Global NSSFValue As Currency
    Global DBFamt As Currency
    Global OOIamt As Currency
    Global HOSPamt As Currency
    Global relief As Currency
    Global Period_Month As String
    Global Period_Year As Integer
    Global Reprint As Boolean
    Global monthp
    Global MonthInt
    Global yearp
    Global TrackMonth
    Global TrackYear
    Global payrollmonth
    Global pcount
    Global gross As Currency
    Global totaldeductions As Currency
    Global benefitamt As Currency
    Global tb As Recordset
    Global datasearch
    Global mover As Boolean
    Global pounds As Currency
    Global holder
    Global Payslip As Integer
    Global othertable As Integer
    Global TaxableInc As Currency
    Global sql
    Global DataChanged As Integer
    Global Opt
    Global Report
    Global defn As String
    Global ID As String
    Global ID2
    Global BackupRequest
    Global tabl
    Global paymonth As Integer
    Global posn As Long
    Global login As Boolean
    Global loginname
    Global spacefound
    Global Action As Integer
    Global vbAdd As Boolean
    Global vbEdit As Boolean
    Global vbDelete As Boolean
    Global datExtra2 As Recordset
    
Global Const MB_OK = 0, MB_OKCANCEL = 1    ' Define buttons.
Global Const MB_YESNOCANCEL = 3, MB_YESNO = 4
Global Const MB_ICONSTOP = 16, MB_ICONQUESTION = 32    ' Define Icons.
Global Const MB_ICONEXCLAMATION = 48, MB_ICONtbRMATION = 64
Global Const MB_DEFBUTTON2 = 256, IDYES = 6, IDNO = 7  ' Define other.
Global Const MB_APPLMODAL = 0, MB_SYSTEMMODAL = 4096
Global Const MB_DEFBUTTON1 = 0, MB_DEFBUTTON3 = 512

'This variable is declared so that the same screen can be used for
'both service charge data and NHIF data.

Global Tracktable As Integer
''''''''''''''''''''''''''''
' Visual Basic global constant file. This file can be loaded
' into a code module.
'
' Some constants are commented out because they have
' duplicates (e.g., NONE appears several places).
'
' If you are updating a Visual Basic application written with
' an older version, you should replace your global constants
' with the constants in this file.
'
''''''''''''''''''''''''''''

' General

' Clipboard formats
Global Const CF_LINK = &HBF00
Global Const CF_TEXT = 1
Global Const CF_BITMAP = 2
Global Const CF_METAFILE = 3
Global Const CF_DIB = 8
Global Const CF_PALETTE = 9

' DragOver
Global Const ENTER = 0
Global Const LEAVE = 1
Global Const OVER = 2

' Drag (controls)
Global Const Cancel = 0
Global Const BEGIN_DRAG = 1
Global Const END_DRAG = 2

' Show parameters
Global Const MODAL = 1
Global Const MODELESS = 0

' Arrange Method
' for MDI Forms
Global Const CASCADE = 0
Global Const TILE_HORIZONTAL = 1
Global Const TILE_VERTICAL = 2
Global Const ARRANGE_ICONS = 3

'ZOrder Method
Global Const BRINGTOFRONT = 0
Global Const SENDTOBACK = 1

' Key Codes
Global Const KEY_LBUTTON = &H1
Global Const KEY_RBUTTON = &H2
Global Const KEY_CANCEL = &H3
Global Const KEY_MBUTTON = &H4    ' NOT contiguous with L & RBUTTON
Global Const KEY_BACK = &H8
Global Const KEY_TAB = &H9
Global Const KEY_CLEAR = &HC
Global Const KEY_RETURN = &HD
Global Const KEY_SHIFT = &H10
Global Const KEY_CONTROL = &H11
Global Const KEY_MENU = &H12
Global Const KEY_PAUSE = &H13
Global Const KEY_CAPITAL = &H14
Global Const KEY_ESCAPE = &H1B
Global Const KEY_SPACE = &H20
Global Const KEY_PRIOR = &H21
Global Const KEY_NEXT = &H22
Global Const KEY_END = &H23
Global Const KEY_HOME = &H24
Global Const KEY_LEFT = &H25
Global Const KEY_UP = &H26
Global Const KEY_RIGHT = &H27
Global Const KEY_DOWN = &H28
Global Const KEY_SELECT = &H29
Global Const KEY_PRINT = &H2A
Global Const KEY_EXECUTE = &H2B
Global Const KEY_SNAPSHOT = &H2C
Global Const KEY_INSERT = &H2D
Global Const KEY_DELETE = &H2E
Global Const KEY_HELP = &H2F

' KEY_A thru KEY_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' KEY_0 thru KEY_9 are the same as their ASCII equivalents: '0' thru '9'

Global Const KEY_NUMPAD0 = &H60
Global Const KEY_NUMPAD1 = &H61
Global Const KEY_NUMPAD2 = &H62
Global Const KEY_NUMPAD3 = &H63
Global Const KEY_NUMPAD4 = &H64
Global Const KEY_NUMPAD5 = &H65
Global Const KEY_NUMPAD6 = &H66
Global Const KEY_NUMPAD7 = &H67
Global Const KEY_NUMPAD8 = &H68
Global Const KEY_NUMPAD9 = &H69
Global Const KEY_MULTIPLY = &H6A
Global Const KEY_ADD = &H6B
Global Const KEY_SEPARATOR = &H6C
Global Const KEY_SUBTRACT = &H6D
Global Const KEY_DECIMAL = &H6E
Global Const KEY_DIVIDE = &H6F
Global Const KEY_F1 = &H70
Global Const KEY_F2 = &H71
Global Const KEY_F3 = &H72
Global Const KEY_F4 = &H73
Global Const KEY_F5 = &H74
Global Const KEY_F6 = &H75
Global Const KEY_F7 = &H76
Global Const KEY_F8 = &H77
Global Const KEY_F9 = &H78
Global Const KEY_F10 = &H79
Global Const KEY_F11 = &H7A
Global Const KEY_F12 = &H7B
Global Const KEY_F13 = &H7C
Global Const KEY_F14 = &H7D
Global Const KEY_F15 = &H7E
Global Const KEY_F16 = &H7F

Global Const KEY_NUMLOCK = &H90

' Variant VarType tags

Global Const V_EMPTY = 0
Global Const V_NULL = 1
Global Const V_INTEGER = 2
Global Const V_LONG = 3
Global Const V_SINGLE = 4
Global Const V_DOUBLE = 5
Global Const V_CURRENCY = 6
Global Const V_DATE = 7
Global Const V_STRING = 8


' Event Parameters

' ErrNum (LinkError)
Global Const WRONG_FORMAT = 1
Global Const DDE_SOURCE_CLOSED = 6
Global Const TOO_MANY_LINKS = 7
Global Const DATA_TRANSFER_FAILED = 8

' QueryUnload
Global Const FORM_CONTROLMENU = 0
Global Const FORM_CODE = 1
Global Const APP_WINDOWS = 2
Global Const APP_TASKMANAGER = 3
Global Const FORM_MDIFORM = 4

' Properties

' Colors
Global Const BLACK = &H0&
Global Const RED = &HFF&
Global Const GREEN = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF

' System Colors
Global Const SCROLL_BARS = &H80000000           ' Scroll-bars gray area.
Global Const DESKTOP = &H80000001               ' Desktop.
Global Const ACTIVE_TITLE_BAR = &H80000002      ' Active window caption.
Global Const INACTIVE_TITLE_BAR = &H80000003    ' Inactive window caption.
Global Const MENU_BAR = &H80000004              ' Menu background.
Global Const WINDOW_BACKGROUND = &H80000005     ' Window background.
Global Const WINDOW_FRAME = &H80000006          ' Window frame.
Global Const MENU_TEXT = &H80000007             ' Text in menus.
Global Const WINDOW_TEXT = &H80000008           ' Text in windows.
Global Const TITLE_BAR_TEXT = &H80000009        ' Text in caption, size box, scroll-bar arrow box..
Global Const ACTIVE_BORDER = &H8000000A         ' Active window border.
Global Const INACTIVE_BORDER = &H8000000B       ' Inactive window border.
Global Const APPLICATION_WORKSPACE = &H8000000C ' Background color of multiple document interface (MDI) applications.
Global Const HIGHLIGHT = &H8000000D             ' Items selected item in a control.
Global Const HIGHLIGHT_TEXT = &H8000000E        ' Text of item selected in a control.
Global Const BUTTON_FACE = &H8000000F           ' Face shading on command buttons.
Global Const BUTTON_SHADOW = &H80000010         ' Edge shading on command buttons.
Global Const GRAY_TEXT = &H80000011             ' Grayed (disabled) text.  This color is set to 0 if the current display driver does not support a solid gray color.
Global Const BUTTON_TEXT = &H80000012           ' Text on push buttons.

' Enumerated Types

' Align (picture box)
Global Const NONE = 0
Global Const ALIGN_TOP = 1
Global Const ALIGN_BOTTOM = 2

' Alignment
Global Const LEFT_JUSTIFY = 0  ' 0 - Left Justify
Global Const RIGHT_JUSTIFY = 1 ' 1 - Right Justify
Global Const CENTER = 2        ' 2 - Center

' BorderStyle (form)
'Global Const NONE = 0          ' 0 - None
Global Const FIXED_SINGLE = 1   ' 1 - Fixed Single
Global Const SIZABLE = 2        ' 2 - Sizable (Forms only)
Global Const FIXED_DOUBLE = 3   ' 3 - Fixed Double (Forms only)

' BorderStyle (Shape and Line)
'Global Const TRANSPARENT = 0    '0 - Transparent
'Global Const SOLID = 1          '1 - Solid
'Global Const DASH = 2         ' 2 - Dash
'Global Const DOT = 3          ' 3 - Dot
'Global Const DASH_DOT = 4     ' 4 - Dash-Dot
'Global Const DASH_DOT_DOT = 5 ' 5 - Dash-Dot-Dot
'Global Const INSIDE_SOLID = 6 ' 6 - Inside Solid

' MousePointer
Global Const Default = 0        ' 0 - Default
Global Const ARROW = 1          ' 1 - Arrow
Global Const CROSSHAIR = 2      ' 2 - Cross
Global Const IBEAM = 3          ' 3 - I-Beam
Global Const ICON_POINTER = 4   ' 4 - Icon
Global Const SIZE_POINTER = 5   ' 5 - Size
Global Const SIZE_NE_SW = 6     ' 6 - Size NE SW
Global Const SIZE_N_S = 7       ' 7 - Size N S
Global Const SIZE_NW_SE = 8     ' 8 - Size NW SE
Global Const SIZE_W_E = 9       ' 9 - Size W E
Global Const UP_ARROW = 10      ' 10 - Up Arrow
Global Const HOURGLASS = 11     ' 11 - Hourglass
Global Const NO_DROP = 12       ' 12 - No drop

' DragMode
Global Const MANUAL = 0    ' 0 - Manual
Global Const AUTOMATIC = 1 ' 1 - Automatic

' DrawMode
Global Const BLACKNESS = 1      ' 1 - Blackness
Global Const NOT_MERGE_PEN = 2  ' 2 - Not Merge Pen
Global Const MASK_NOT_PEN = 3   ' 3 - Mask Not Pen
Global Const NOT_COPY_PEN = 4   ' 4 - Not Copy Pen
Global Const MASK_PEN_NOT = 5   ' 5 - Mask Pen Not
Global Const INVERT = 6         ' 6 - Invert
Global Const XOR_PEN = 7        ' 7 - Xor Pen
Global Const NOT_MASK_PEN = 8   ' 8 - Not Mask Pen
Global Const MASK_PEN = 9       ' 9 - Mask Pen
Global Const NOT_XOR_PEN = 10   ' 10 - Not Xor Pen
Global Const NOP = 11           ' 11 - Nop
Global Const MERGE_NOT_PEN = 12 ' 12 - Merge Not Pen
Global Const COPY_PEN = 13      ' 13 - Copy Pen
Global Const MERGE_PEN_NOT = 14 ' 14 - Merge Pen Not
Global Const MERGE_PEN = 15     ' 15 - Merge Pen
Global Const WHITENESS = 16     ' 16 - Whiteness

' DrawStyle
Global Const SOLID = 0        ' 0 - Solid
Global Const DASH = 1         ' 1 - Dash
Global Const DOT = 2          ' 2 - Dot
Global Const DASH_DOT = 3     ' 3 - Dash-Dot
Global Const DASH_DOT_DOT = 4 ' 4 - Dash-Dot-Dot
Global Const INVISIBLE = 5    ' 5 - Invisible
Global Const INSIDE_SOLID = 6 ' 6 - Inside Solid

' FillStyle
' Global Const SOLID = 0           ' 0 - Solid
Global Const Transparent = 1       ' 1 - Transparent
Global Const HORIZONTAL_LINE = 2   ' 2 - Horizontal Line
Global Const VERTICAL_LINE = 3     ' 3 - Vertical Line
Global Const UPWARD_DIAGONAL = 4   ' 4 - Upward Diagonal
Global Const DOWNWARD_DIAGONAL = 5 ' 5 - Downward Diagonal
Global Const CROSS = 6             ' 6 - Cross
Global Const DIAGONAL_CROSS = 7    ' 7 - Diagonal Cross

' LinkMode (forms and controls)
' Global Const NONE = 0         ' 0 - None
Global Const LINK_SOURCE = 1    ' 1 - Source (forms only)
Global Const LINK_AUTOMATIC = 1 ' 1 - Automatic (controls only)
Global Const LINK_MANUAL = 2    ' 2 - Manual (controls only)
Global Const LINK_NOTIFY = 3    ' 3 - Notify (controls only)

' LinkMode (kept for VB1.0 compatibility, use new constants instead)
Global Const HOT = 1    ' 1 - Hot (controls only)
Global Const SERVER = 1 ' 1 - Server (forms only)
Global Const COLD = 2   ' 2 - Cold (controls only)


' ScaleMode
Global Const User = 0        ' 0 - User
Global Const TWIPS = 1       ' 1 - Twip
Global Const POINTS = 2      ' 2 - Point
Global Const PIXELS = 3      ' 3 - Pixel
Global Const CHARACTERS = 4  ' 4 - Character
Global Const INCHES = 5      ' 5 - Inch
Global Const MILLIMETERS = 6 ' 6 - Millimeter
Global Const CENTIMETERS = 7 ' 7 - Centimeter

' ScrollBar
' Global Const NONE     = 0 ' 0 - None
Global Const HORIZONTAL = 1 ' 1 - Horizontal
Global Const VERTICAL = 2   ' 2 - Vertical
Global Const BOTH = 3       ' 3 - Both

' Shape
Global Const SHAPE_RECTANGLE = 0
Global Const SHAPE_SQUARE = 1
Global Const SHAPE_OVAL = 2
Global Const SHAPE_CIRCLE = 3
Global Const SHAPE_roundit_RECTANGLE = 4
Global Const SHAPE_roundit_SQUARE = 5

' WindowState
Global Const NORMAL = 0    ' 0 - Normal
Global Const MINIMIZED = 1 ' 1 - Minimized
Global Const MAXIMIZED = 2 ' 2 - Maximized

' Check Value
Global Const UNCHECKED = 0 ' 0 - Unchecked
Global Const CHECKED = 1   ' 1 - Checked
Global Const GRAYED = 2    ' 2 - Grayed

' Shift parameter masks
Global Const SHIFT_MASK = 1
Global Const CTRL_MASK = 2
Global Const ALT_MASK = 4

' Button parameter masks
Global Const LEFT_BUTTON = 1
Global Const RIGHT_BUTTON = 2
Global Const MIDDLE_BUTTON = 4

' Function Parameters
' MsgBox parameters
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons



' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed

' SetAttr, Dir, GetAttr functions
Global Const ATTR_NORMAL = 0
Global Const ATTR_READONLY = 1
Global Const ATTR_HIDDEN = 2
Global Const ATTR_SYSTEM = 4
Global Const ATTR_VOLUME = 8
Global Const ATTR_DIRECTORY = 16
Global Const ATTR_ARCHIVE = 32

'Grid
'ColAlignment,FixedAlignment Properties
Global Const GRID_ALIGNLEFT = 0
Global Const GRID_ALIGNRIGHT = 1
Global Const GRID_ALIGNCENTER = 2

'Fillstyle Property
Global Const GRID_SINGLE = 0
Global Const GRID_REPEAT = 1


'Data control
'Error event Response arguments
Global Const DATA_ERRCONTINUE = 0
Global Const DATA_ERRDISPLAY = 1

'Editmode property values
Global Const DATA_EDITNONE = 0
Global Const DATA_EDITMODE = 1
Global Const DATA_EDITADD = 2

' Options property values
Global Const DATA_DENYWRITE = &H1
Global Const DATA_DENYREAD = &H2
Global Const DATA_READONLY = &H4
Global Const DATA_APPENDONLY = &H8
Global Const DATA_INCONSISTENT = &H10
Global Const DATA_CONSISTENT = &H20
Global Const DATA_SQLPASSTHROUGH = &H40

'Validate event Action arguments
Global Const DATA_ACTIONCANCEL = 0
Global Const DATA_ACTIONMOVEFIRST = 1
Global Const DATA_ACTIONMOVEPREVIOUS = 2
Global Const DATA_ACTIONMOVENEXT = 3
Global Const DATA_ACTIONMOVELAST = 4
Global Const DATA_ACTIONADDNEW = 5
Global Const DATA_ACTIONUPDATE = 6
Global Const DATA_ACTIONDELETE = 7
Global Const DATA_ACTIONFIND = 8
Global Const DATA_ACTIONBOOKMARK = 9
Global Const DATA_ACTIONCLOSE = 10
Global Const DATA_ACTIONUNLOAD = 11


'OLE Client Control
'Actions
Global Const OLE_CREATE_EMBED = 0
Global Const OLE_CREATE_NEW = 0           'from ole1 control
Global Const OLE_CREATE_LINK = 1
Global Const OLE_CREATE_FROM_FILE = 1     'from ole1 control
Global Const OLE_COPY = 4
Global Const OLE_PASTE = 5
Global Const OLE_UPDATE = 6
Global Const OLE_ACTIVATE = 7
Global Const OLE_CLOSE = 9
Global Const OLE_DELETE = 10
Global Const OLE_SAVE_TO_FILE = 11
Global Const OLE_READ_FROM_FILE = 12
Global Const OLE_INSERT_OBJ_DLG = 14
Global Const OLE_PASTE_SPECIAL_DLG = 15
Global Const OLE_FETCH_VERBS = 17
Global Const OLE_SAVE_TO_OLE1FILE = 18

'OLEType
Global Const OLE_LINKED = 0
Global Const OLE_EMBEDDED = 1
Global Const OLE_NONE = 3

'OLETypeAllowed
Global Const OLE_EITHER = 2

'UpdateOptions
Global Const OLE_AUTOMATIC = 0
Global Const OLE_FROZEN = 1
Global Const OLE_MANUAL = 2

'AutoActivate modes
'Note that OLE_ACTIVATE_GETFOCUS only applies to objects that
'support "inside-out" activation.  See related Verb notes below.
Global Const OLE_ACTIVATE_MANUAL = 0
Global Const OLE_ACTIVATE_GETFOCUS = 1
Global Const OLE_ACTIVATE_DOUBLECLICK = 2

'SizeModes
Global Const OLE_SIZE_CLIP = 0
Global Const OLE_SIZE_STRETCH = 1
Global Const OLE_SIZE_AUTOSIZE = 2

'DisplayTypes
Global Const OLE_DISPLAY_CONTENT = 0
Global Const OLE_DISPLAY_ICON = 1

'Update Event Constants
Global Const OLE_CHANGED = 0
Global Const OLE_SAVED = 1
Global Const OLE_CLOSED = 2
Global Const OLE_RENAMED = 3

'Special Verb Values
Global Const VERB_PRIMARY = 0
Global Const VERB_SHOW = -1
Global Const VERB_OPEN = -2
Global Const VERB_HIDE = -3
Global Const VERB_INPLACEUIACTIVATE = -4
Global Const VERB_INPLACEACTIVATE = -5
'The last two verbs are for objects that support "inside-out" activation,
'meaning they can be edited in-place, and that they support being left
'in-place-active even when the input focus moves to another control or form.
'These objects actually have 2 levels of being active.  "InPlace Active"
'means that the object is ready for the user to click inside it and start
'working with it.  "In-Place UI-Active" means that, in addition, if the object
'has any other UI associated with it, such as floating palette windows,
'that those windows are visible and ready for use.  Any number of objects
'can be "In-Place Active" at a time, although only one can be
'"InPlace UI-Active".

'You can cause an object to move to either one of states programmatically by
'setting the Verb property to the appropriate verb and setting
'Action=OLE_ACTIVATE.

'Also, if you set AutoActivate = OLE_ACTIVATE_GETFOCUS, the server will
'automatically be put into "InPlace UI-Active" state when the user clicks
'on or tabs into the control.

'VerbFlag Bit Masks
Global Const VERBFLAG_GRAYED = &H1
Global Const VERBFLAG_DISABLED = &H2
Global Const VERBFLAG_CHECKED = &H8
Global Const VERBFLAG_SEPARATOR = &H800

'MiscFlag Bits - Or these together as desired for special behaviors

'MEMSTORAGE causes the control to use memory to store the object while
'           it is loaded.  This is faster than the default (disk-tempfile),
'           but can consume a lot of memory for objects whose data takes
'           up a lot of space, such as the bitmap for a paint program.
Global Const OLE_MISCFLAG_MEMSTORAGE = &H1

'DISABLEINPLACE overrides the control's default behavior of allowing
'           in-place activation for objects that support it.  If you
'           are having problems activating an object inplace, you can
'           force it to always activate in a separate window by setting this
'           bit
Global Const OLE_MISCFLAG_DISABLEINPLACE = &H2

'Common Dialog Control
'Action Property
Global Const DLG_FILE_OPEN = 1
Global Const DLG_FILE_SAVE = 2
Global Const DLG_COLOR = 3
Global Const DLG_FONT = 4
Global Const DLG_PRINT = 5
Global Const DLG_HELP = 6

'File Open/Save Dialog Flags
Global Const OFN_READONLY = &H1&
Global Const OFN_OVERWRITEPROMPT = &H2&
Global Const OFN_HIDEREADONLY = &H4&
Global Const OFN_NOCHANGEDIR = &H8&
Global Const OFN_SHOWHELP = &H10&
Global Const OFN_NOVALIDATE = &H100&
Global Const OFN_ALLOWMULTISELECT = &H200&
Global Const OFN_EXTENSIONDIFFERENT = &H400&
Global Const OFN_PATHMUSTEXIST = &H800&
Global Const OFN_FILEMUSTEXIST = &H1000&
Global Const OFN_CREATEPROMPT = &H2000&
Global Const OFN_SHAREAWARE = &H4000&
Global Const OFN_NOREADONLYRETURN = &H8000&

'Color Dialog Flags
Global Const CC_RGBINIT = &H1&
Global Const CC_FULLOPEN = &H2&
Global Const CC_PREVENTFULLOPEN = &H4&
Global Const CC_SHOWHELP = &H8&

'Fonts Dialog Flags
Global Const CF_SCREENFONTS = &H1&
Global Const CF_PRINTERFONTS = &H2&
Global Const CF_BOTH = &H3&
Global Const CF_SHOWHELP = &H4&
Global Const CF_INITTOLOGFONTSTRUCT = &H40&
Global Const CF_USESTYLE = &H80&
Global Const CF_EFFECTS = &H100&
Global Const CF_APPLY = &H200&
Global Const CF_ANSIONLY = &H400&
Global Const CF_NOVECTORFONTS = &H800&
Global Const CF_NOSIMULATIONS = &H1000&
Global Const CF_LIMITSIZE = &H2000&
Global Const CF_FIXEDPITCHONLY = &H4000&
Global Const CF_WYSIWYG = &H8000&         'must also have CF_SCREENFONTS & CF_PRINTERFONTS
Global Const CF_FORCEFONTEXIST = &H10000
Global Const CF_SCALABLEONLY = &H20000
Global Const CF_TTONLY = &H40000
Global Const CF_NOFACESEL = &H80000
Global Const CF_NOSTYLESEL = &H100000
Global Const CF_NOSIZESEL = &H200000

'Printer Dialog Flags
Global Const PD_ALLPAGES = &H0&
Global Const PD_SELECTION = &H1&
Global Const PD_PAGENUMS = &H2&
Global Const PD_NOSELECTION = &H4&
Global Const PD_NOPAGENUMS = &H8&
Global Const PD_COLLATE = &H10&
Global Const PD_PRINTTOFILE = &H20&
Global Const PD_PRINTSETUP = &H40&
Global Const PD_NOWARNING = &H80&
Global Const PD_RETURNDC = &H100&
Global Const PD_RETURNIC = &H200&
Global Const PD_RETURNDEFAULT = &H400&
Global Const PD_SHOWHELP = &H800&
Global Const PD_USEDEVMODECOPIES = &H40000
Global Const PD_DISABLEPRINTTOFILE = &H80000
Global Const PD_HIDEPRINTTOFILE = &H100000

'Help Constants
Global Const HELP_CONTEXT = &H1           'Display topic in ulTopic
Global Const HELP_QUIT = &H2              'Terminate help
Global Const HELP_INDEX = &H3             'Display index
Global Const HELP_CONTENTS = &H3
Global Const HELP_HELPONHELP = &H4        'Display help on using help
Global Const HELP_SETINDEX = &H5          'Set the current Index for multi index help
Global Const HELP_SETCONTENTS = &H5
Global Const HELP_CONTEXTPOPUP = &H8
Global Const HELP_FORCEFILE = &H9
Global Const HELP_KEY = &H101             'Display topic for keyword in offabData
Global Const HELP_COMMAND = &H102
Global Const HELP_PARTIALKEY = &H105      'call the search engine in winhelp

'Error Constants
Global Const CDERR_DIALOGFAILURE = -32768

Global Const CDERR_GENERALCODES = &H7FFF
Global Const CDERR_STRUCTSIZE = &H7FFE
Global Const CDERR_INITIALIZATION = &H7FFD
Global Const CDERR_NOTEMPLATE = &H7FFC
Global Const CDERR_NOHINSTANCE = &H7FFB
Global Const CDERR_LOADSTRFAILURE = &H7FFA
Global Const CDERR_FINDRESFAILURE = &H7FF9
Global Const CDERR_LOADRESFAILURE = &H7FF8
Global Const CDERR_LOCKRESFAILURE = &H7FF7
Global Const CDERR_MEMALLOCFAILURE = &H7FF6
Global Const CDERR_MEMLOCKFAILURE = &H7FF5
Global Const CDERR_NOHOOK = &H7FF4

'Added for CMDIALOG.VBX
Global Const CDERR_CANCEL = &H7FF3
Global Const CDERR_NODLL = &H7FF2
Global Const CDERR_ERRPROC = &H7FF1
Global Const CDERR_ALLOC = &H7FF0
Global Const CDERR_HELP = &H7FEF

Global Const PDERR_PRINTERCODES = &H6FFF
Global Const PDERR_SETUPFAILURE = &H6FFE
Global Const PDERR_PARSEFAILURE = &H6FFD
Global Const PDERR_RETDEFFAILURE = &H6FFC
Global Const PDERR_LOADDRVFAILURE = &H6FFB
Global Const PDERR_GETDEVMODEFAIL = &H6FFA
Global Const PDERR_INITFAILURE = &H6FF9
Global Const PDERR_NODEVICES = &H6FF8
Global Const PDERR_NODEFAULTPRN = &H6FF7
Global Const PDERR_DNDMMISMATCH = &H6FF6
Global Const PDERR_CREATEICFAILURE = &H6FF5
Global Const PDERR_PRINTERNOTFOUND = &H6FF4

Global Const CFERR_CHOOSEFONTCODES = &H5FFF
Global Const CFERR_NOFONTS = &H5FFE

Global Const FNERR_FILENAMECODES = &H4FFF
Global Const FNERR_SUBCLASSFAILURE = &H4FFE
Global Const FNERR_INVALIDFILENAME = &H4FFD
Global Const FNERR_BUFFERTOOSMALL = &H4FFC

Global Const FRERR_FINDREPLACECODES = &H3FFF
Global Const CCERR_CHOOSECOLORCODES = &H2FFF


'---------------------------------------------------------
'      Table of Contents for Visual Basic Professional
'
'       1.  3-D Controls
'           (Frame/Panel/Option/Check/Command/Group Push)
'       2.  Animated Button
'       3.  Gauge Control
'       4.  Graph Control Section
'       5.  Key Status Control
'       6.  Spin Button
'       7.  MCI Control (Multimedia)
'       8.  Masked Edit Control
'       9.  Comm Control
'       10. Outline Control
'---------------------------------------------------------


'-------------------------------------------------------------------
'3D Controls
'-------------------------------------------------------------------
'Alignment (Check Box)
Global Const SSCB_TEXT_RIGHT = 0         '0 - Text to the right
Global Const SSCB_TEXT_LEFT = 1          '1 - Text to the left

'Alignment (Option Button)
Global Const SSOB_TEXT_RIGHT = 0         '0 - Text to the right
Global Const SSOB_TEXT_LEFT = 1          '1 - Text to the left

'Alignment (Frame)
Global Const SSFR_LEFT_JUSTIFY = 0       '0 - Left justify text
Global Const SSFR_RIGHT_JUSTIFY = 1      '1 - Right justify text
Global Const SSFR_CENTER = 2             '2 - Center text

'Alignment (Panel)
Global Const SSPN_LEFT_TOP = 0           '0 - Text to left and top
Global Const SSPN_LEFT_MIDDLE = 1        '1 - Text to left and middle
Global Const SSPN_LEFT_BOTTOM = 2        '2 - Text to left and bottom
Global Const SSPN_RIGHT_TOP = 3          '3 - Text to right and top
Global Const SSPN_RIGHT_MIDDLE = 4       '4 - Text to right and middle
Global Const SSPN_RIGHT_BOTTOM = 5       '5 - Text to right and bottom
Global Const SSPN_CENTER_TOP = 6         '6 - Text to center and top
Global Const SSPN_CENTER_MIDDLE = 7      '7 - Text to center and middle
Global Const SSPN_CENTER_BOTTOM = 8      '8 - Text to center and bottom

'Autosize (Command Button)
Global Const SS_AUTOSIZE_NONE = 0        '0 - No Autosizing
Global Const SSPB_AUTOSIZE_PICTOBUT = 1  '0 - Autosize Picture to Button
Global Const SSPB_AUTOSIZE_BUTTOPIC = 2  '0 - Autosize Button to Picture

'Autosize (Ribbon Button)
'Global Const SS_AUTOSIZE_NONE      = 0  '0 - No Autosizing
Global Const SSRI_AUTOSIZE_PICTOBUT = 1  '0 - Autosize Picture to Button
Global Const SSRI_AUTOSIZE_BUTTOPIC = 2  '0 - Autosize Button to Picture

'Autosize (Panel)
'Global Const SS_AUTOSIZE_NONE    = 0    '0 - No Autosizing
Global Const SSPN_AUTOSIZE_WIDTH = 1     '1 - Autosize Panel width to Caption
Global Const SSPN_AUTOSIZE_HEIGHT = 2    '2 - Autosize Panel height to Caption
Global Const SSPN_AUTOSIZE_CHILD = 3     '3 - Autosize Child to Panel

'BevelInner (Panel)
Global Const SS_BEVELINNER_NONE = 0      '0 - No Inner Bevel
Global Const SS_BEVELINNER_INSET = 1     '1 - Inset Inner Bevel
Global Const SS_BEVELINNER_RAISED = 2    '2 - Raised Inner Bevel

'BevelOuter (Panel)
Global Const SS_BEVELOUTER_NONE = 0      '0 - No Outer Bevel
Global Const SS_BEVELOUTER_INSET = 1     '1 - Inset Outer Bevel
Global Const SS_BEVELOUTER_RAISED = 2    '2 - Raised Outer Bevel

'FloodType (Panel)
Global Const SS_FLOODTYPE_NONE = 0       '0 - No flood
Global Const SS_FLOODTYPE_L_TO_R = 1     '1 - Left to light
Global Const SS_FLOODTYPE_R_TO_L = 2     '2 - Right to left
Global Const SS_FLOODTYPE_T_TO_B = 3     '3 - Top to bottom
Global Const SS_FLOODTYPE_B_TO_T = 4     '4 - Bottom to top
Global Const SS_FLOODTYPE_CIRCLE = 5     '5 - Widening circle

'Font3D (Panel, Command Button, Option Button, Check Box, Frame)
Global Const SS_FONT3D_NONE = 0          '0 - No 3-D text
Global Const SS_FONT3D_RAISED_LIGHT = 1  '1 - Raised with light shading
Global Const SS_FONT3D_RAISED_HEAVY = 2  '2 - Raised with heavy shading
Global Const SS_FONT3D_INSET_LIGHT = 3   '3 - Inset with light shading
Global Const SS_FONT3D_INSET_HEAVY = 4   '4 - Inset with heavy shading

'PictureDnChange (Ribbon Button)
Global Const SS_PICDN_NOCHANGE = 0       '0 - Use 'Up'bitmap with no change
Global Const SS_PICDN_DITHER = 1         '1 - Dither 'Up'bitmap
Global Const SS_PICDN_INVERT = 2         '2 - Invert 'Up'bitmap

'ShadowColor (Panel, Frame)
Global Const SS_SHADOW_DARKGREY = 0      '0 - Dark grey shadow
Global Const SS_SHADOW_BLACK = 1         '1 - Black shadow

'ShadowStyle (Frame)
Global Const SS_SHADOW_INSET = 0         '0 - Shadow inset
Global Const SS_SHADOW_RAISED = 1        '1 - Shadow raised


'---------------------------------------
'Animated Button
'---------------------------------------
'Cycle property
Global Const ANI_ANIMATED = 0
Global Const ANI_MULTISTATE = 1
Global Const ANI_TWO_STATE = 2

'Click Filter property
Global Const ANI_ANYWHERE = 0
Global Const ANI_IMAGE_AND_TEXT = 1
Global Const ANI_IMAGE = 2
Global Const ANI_TEXT = 3

'PicDrawMode Property
Global Const ANI_XPOS_YPOS = 0
Global Const ANI_AUTOSIZE = 1
Global Const ANI_STRETCH = 2

'SpecialOp Property
Global Const ANI_CLICK = 1

'TextPosition Property
Global Const ANI_CENTER = 0
Global Const ANI_LEFT = 1
Global Const ANI_RIGHT = 2
Global Const ANI_BOTTON = 3
Global Const ANI_TOP = 4


'---------------------------------------
'GAUGE
'---------------------------------------
'Style Property
Global Const GAUGE_HORIZ = 0
Global Const GAUGE_VERT = 1
Global Const GAUGE_SEMI = 2
Global Const GAUGE_FULL = 3


'----------------------------------------
'Graph Control
'----------------------------------------
'General
Global Const G_NONE = 0
Global Const G_DEFAULT = 0

Global Const G_OFF = 0
Global Const G_ON = 1

Global Const G_MONO = 0
Global Const G_COLOR = 1

'Graph Types
Global Const G_PIE2D = 1
Global Const G_PIE3D = 2
Global Const G_BAR2D = 3
Global Const G_BAR3D = 4
Global Const G_GANTT = 5
Global Const G_LINE = 6
Global Const G_LOGLIN = 7
Global Const G_AREA = 8
Global Const G_SCATTER = 9
Global Const G_POLAR = 10
Global Const G_HLC = 11

'Colors
Global Const G_BLACK = 0
Global Const G_BLUE = 1
Global Const G_GREEN = 2
Global Const G_CYAN = 3
Global Const G_RED = 4
Global Const G_MAGENTA = 5
Global Const G_BROWN = 6
Global Const G_LIGHT_GRAY = 7
Global Const G_DARK_GRAY = 8
Global Const G_LIGHT_BLUE = 9
Global Const G_LIGHT_GREEN = 10
Global Const G_LIGHT_CYAN = 11
Global Const G_LIGHT_RED = 12
Global Const G_LIGHT_MAGENTA = 13
Global Const G_YELLOW = 14
Global Const G_WHITE = 15
Global Const G_AUTOBW = 16

'Patterns
Global Const G_SOLID = 0
Global Const G_HOLLOW = 1
Global Const G_HATCH1 = 2
Global Const G_HATCH2 = 3
Global Const G_HATCH3 = 4
Global Const G_HATCH4 = 5
Global Const G_HATCH5 = 6
Global Const G_HATCH6 = 7
Global Const G_BITMAP1 = 16
Global Const G_BITMAP2 = 17
Global Const G_BITMAP3 = 18
Global Const G_BITMAP4 = 19
Global Const G_BITMAP5 = 20
Global Const G_BITMAP6 = 21
Global Const G_BITMAP7 = 22
Global Const G_BITMAP8 = 23
Global Const G_BITMAP9 = 24
Global Const G_BITMAP10 = 25
Global Const G_BITMAP11 = 26
Global Const G_BITMAP12 = 27
Global Const G_BITMAP13 = 28
Global Const G_BITMAP14 = 29
Global Const G_BITMAP15 = 30
Global Const G_BITMAP16 = 31

'Symbols
Global Const G_CROSS_PLUS = 0
Global Const G_CROSS_TIMES = 1
Global Const G_TRIANGLE_UP = 2
Global Const G_SOLID_TRIANGLE_UP = 3
Global Const G_TRIANGLE_DOWN = 4
Global Const G_SOLID_TRIANGLE_DOWN = 5
Global Const G_SQUARE = 6
Global Const G_SOLID_SQUARE = 7
Global Const G_DIAMOND = 8
Global Const G_SOLID_DIAMOND = 9

'Line Styles
'Global Const G_SOLID = 0
Global Const G_DASH = 1
Global Const G_DOT = 2
Global Const G_DASHDOT = 3
Global Const G_DASHDOTDOT = 4

'Grids
Global Const G_HORIZONTAL = 1
Global Const G_VERTICAL = 2

'Statistics
Global Const G_MEAN = 1
Global Const G_MIN_MAX = 2
Global Const G_STD_DEV = 4
Global Const G_BEST_FIT = 8

'Data Arrays
Global Const G_GRAPH_DATA = 1
Global Const G_COLOR_DATA = 2
Global Const G_EXTRA_DATA = 3
Global Const G_LABEL_TEXT = 4
Global Const G_LEGEND_TEXT = 5
Global Const G_PATTERN_DATA = 6
Global Const G_SYMBOL_DATA = 7
Global Const G_XPOS_DATA = 8
Global Const G_ALL_DATA = 9

'Draw Mode
Global Const G_NO_ACTION = 0
Global Const G_CLEAR = 1
Global Const G_DRAW = 2
Global Const G_BLIT = 3
Global Const G_COPY = 4
Global Const G_PRINT = 5
Global Const G_WRITE = 6

'Print Options
Global Const G_BORDER = 2

'Pie Chart Options             '
Global Const G_NO_LINES = 1
Global Const G_COLORED = 2
Global Const G_PERCENTS = 4

'Bar Chart Options             '
'Global Const G_HORIZONTAL = 1
Global Const G_STACKED = 2
Global Const G_PERCENTAGE = 4
Global Const G_Z_CLUSTERED = 6

'Gantt Chart Options           '
Global Const G_SPACED_BARS = 1

'Line/Polar Chart Options      '
Global Const G_SYMBOLS = 1
Global Const G_STICKS = 2
Global Const G_LINES = 4

'Area Chart Options            '
Global Const G_ABSOLUTE = 1
Global Const G_PERCENT = 2

'HLC Chart Options             '
Global Const G_NO_CLOSE = 1
Global Const G_NO_HIGH_LOW = 2


'---------------------------------------
'Key Status Control
'---------------------------------------
'Style
Global Const KEYSTAT_CAPSLOCK = 0
Global Const KEYSTAT_NUMLOCK = 1
Global Const KEYSTAT_INSERT = 2
Global Const KEYSTAT_SCROLLLOCK = 3


'---------------------------------------
'MCI Control (Multimedia)
'---------------------------------------
'NOTE:
'Please use the updated Multimedia constants
'in the WINMMSYS.TXT file from the \VB\WINAPI
'subdirectory.

'Mode Property
'Global Const MCI_MODE_NOT_OPEN = 11
'Global Const MCI_MODE_STOP = 12
'Global Const MCI_MODE_PLAY = 13
'Global Const MCI_MODE_RECORD = 14
'Global Const MCI_MODE_SEEK = 15
'Global Const MCI_MODE_PAUSE = 16
'Global Const MCI_MODE_READY = 17

'NotifyValue Property
'Global Const MCI_NOTIFY_SUCCESSFUL = 1
'Global Const MCI_NOTIFY_SUPERSEDED = 2
'Global Const MCI_ABORTED = 4
'Global Const MCI_FAILURE = 8

'Orientation Property
'Global Const MCI_ORIENT_HORZ = 0
'Global Const MCI_ORIENT_VERT = 1

'RecordMode Porperty
'Global Const MCI_RECORD_INSERT = 0
'Global Const MCI_RECORD_OVERWRITE = 1

'TimeFormat Property
'Global Const MCI_FORMAT_MILLISECONDS = 0
'Global Const MCI_FORMAT_HMS = 1
'Global Const MCI_FORMAT_MSF = 2
'Global Const MCI_FORMAT_FRAMES = 3
'Global Const MCI_FORMAT_SMPTE_24 = 4
'Global Const MCI_FORMAT_SMPTE_25 = 5
'Global Const MCI_FORMAT_SMPTE_30 = 6
'Global Const MCI_FORMAT_SMPTE_30DROP = 7
'Global Const MCI_FORMAT_BYTES = 8
'Global Const MCI_FORMAT_SAMPLES = 9
'Global Const MCI_FORMAT_TMSF = 10


'---------------------------------------
'Spin Button
'---------------------------------------
'SpinOrientation
Global Const SPIN_VERTICAL = 0
Global Const SPIN_HORIZONTAL = 1


'---------------------------------------
'Masked Edit Control
'---------------------------------------
'ClipMode
Global Const ME_INCLIT = 0
Global Const ME_EXCLIT = 1


'---------------------------------------
'Comm Control
'---------------------------------------
'Handshaking
Global Const MSCOMM_HANDSHAKE_NONE = 0
Global Const MSCOMM_HANDSHAKE_XONXOFF = 1
Global Const MSCOMM_HANDSHAKE_RTS = 2
Global Const MSCOMM_HANDSHAKE_RTSXONXOFF = 3

'Event constants
Global Const MSCOMM_EV_SEND = 1
Global Const MSCOMM_EV_RECEIVE = 2
Global Const MSCOMM_EV_CTS = 3
Global Const MSCOMM_EV_DSR = 4
Global Const MSCOMM_EV_CD = 5
Global Const MSCOMM_EV_RING = 6
Global Const MSCOMM_EV_EOF = 7

'Error code constants
Global Const MSCOMM_ER_BREAK = 1001
Global Const MSCOMM_ER_CTSTO = 1002
Global Const MSCOMM_ER_DSRTO = 1003
Global Const MSCOMM_ER_FRAME = 1004
Global Const MSCOMM_ER_OVERRUN = 1006
Global Const MSCOMM_ER_CDTO = 1007
Global Const MSCOMM_ER_RXOVER = 1008
Global Const MSCOMM_ER_RXPARITY = 1009
Global Const MSCOMM_ER_TXFULL = 1010


'---------------------------------------
' MAPI SESSION CONTROL CONSTANTS
'---------------------------------------
'Action
Global Const SESSION_SIGNON = 1
Global Const SESSION_SIGNOFF = 2


'---------------------------------------
' MAPI MESSAGE CONTROL CONSTANTS
'---------------------------------------
'Action
Global Const MESSAGE_FETCH = 1             ' Load all messages from message store
Global Const MESSAGE_SENDDLG = 2           ' Send mail bring up default mapi dialog
Global Const MESSAGE_SEND = 3              ' Send mail without default mapi dialog
Global Const MESSAGE_SAVEMSG = 4           ' Save message in the compose buffer
Global Const MESSAGE_COPY = 5              ' Copy current message to compose buffer
Global Const MESSAGE_COMPOSE = 6           ' Initialize compose buffer (previous
                                           ' data is lost
Global Const MESSAGE_REPLY = 7             ' Fill Compose buffer as REPLY
Global Const MESSAGE_REPLYALL = 8          ' Fill Compose buffer as REPLY ALL
Global Const MESSAGE_FORWARD = 9           ' Fill Compose buffer as FORWARD
Global Const MESSAGE_DELETE = 10           ' Delete current message
Global Const MESSAGE_SHOWADBOOK = 11       ' Show Address book
Global Const MESSAGE_SHOWDETAILS = 12      ' Show details of the current recipient
Global Const MESSAGE_RESOLVENAME = 13      ' Resolve the display name of the recipient
Global Const RECIPIENT_DELETE = 14            ' Fill Compose buffer as FORWARD
Global Const ATTACHMENT_DELETE = 15          ' Delete current message


'---------------------------------------
'  ERROR CONSTANT DECLARATIONS (MAPI CONTROLS)
'---------------------------------------
Global Const SUCCESS_SUCCESS = 32000
Global Const MAPI_USER_ABORT = 32001
Global Const MAPI_E_FAILURE = 32002
Global Const MAPI_E_LOGIN_FAILURE = 32003
Global Const MAPI_E_DISK_FULL = 32004
Global Const MAPI_E_INSUFFICIENT_MEMORY = 32005
Global Const MAPI_E_ACCESS_DENIED = 32006
Global Const MAPI_E_TOO_MANY_SESSIONS = 32008
Global Const MAPI_E_TOO_MANY_FILES = 32009
Global Const MAPI_E_TOO_MANY_RECIPIENTS = 32010
Global Const MAPI_E_ATTACHMENT_NOT_FOUND = 32011
Global Const MAPI_E_ATTACHMENT_OPEN_FAILURE = 32012
Global Const MAPI_E_ATTACHMENT_WRITE_FAILURE = 32013
Global Const MAPI_E_UNKNOWN_RECIPIENT = 32014
Global Const MAPI_E_BAD_RECIPTYPE = 32015
Global Const MAPI_E_NO_MESSAGES = 32016
Global Const MAPI_E_INVALID_MESSAGE = 32017
Global Const MAPI_E_TEXT_TOO_LARGE = 32018
Global Const MAPI_E_INVALID_SESSION = 32019
Global Const MAPI_E_TYPE_NOT_SUPPORTED = 32020
Global Const MAPI_E_AMBIGUOUS_RECIPIENT = 32021
Global Const MAPI_E_MESSAGE_IN_USE = 32022
Global Const MAPI_E_NETWORK_FAILURE = 32023
Global Const MAPI_E_INVALID_EDITFIELDS = 32024
Global Const MAPI_E_INVALID_RECIPS = 32025
Global Const MAPI_E_NOT_SUPPORTED = 32026

Global Const CONTROL_E_SESSION_EXISTS = 32050
Global Const CONTROL_E_INVALID_BUFFER = 32051
Global Const CONTROL_E_INVALID_READ_BUFFER_ACTION = 32052
Global Const CONTROL_E_NO_SESSION = 32053
Global Const CONTROL_E_INVALID_RECIPIENT = 32054
Global Const CONTROL_E_INVALID_COMPOSE_BUFFER_ACTION = 32055
Global Const CONTROL_E_FAILURE = 32056
Global Const CONTROL_E_NO_RECIPIENTS = 32057
Global Const CONTROL_E_NO_ATTACHMENTS = 32058


'---------------------------------------
'  MISCELLANEOUS GLOBAL CONSTANT DECLARATIONS (MAPI CONTROLS)
'---------------------------------------
Global Const RECIPTYPE_ORIG = 0
Global Const RECIPTYPE_TO = 1
Global Const RECIPTYPE_CC = 2
Global Const RECIPTYPE_BCC = 3

Global Const ATTACHTYPE_DATA = 0
Global Const ATTACHTYPE_EOLE = 1
Global Const ATTACHTYPE_SOLE = 2


'-------------------------------------------------
'  Outline
'-------------------------------------------------
' PictureType
Global Const MSOUTLINE_PICTURE_CLOSED = 0
Global Const MSOUTLINE_PICTURE_OPEN = 1
Global Const MSOUTLINE_PICTURE_LEAF = 2

'Outline Control Error Constants
Global Const MSOUTLINE_BADPICFORMAT = 32000
Global Const MSOUTLINE_BADINDENTATION = 32001
Global Const MSOUTLINE_MEM = 32002
Global Const MSOUTLINE_PARENTNOTEXPANDED = 32003
Global prn As Printer
Global cn As Connection
'Global datExtra2 As Recordset
Global datPeriod As Recordset
Global Users As Recordset
Global datEmp As Recordset
Dim paygen As New Recordset
Global datPayroll As Recordset




Function CopyData() As Integer
  Dim recRecordset1 As Recordset, recRecordset2 As Recordset
  Dim i As Integer
  Dim nRC As Integer
  Dim fld As Field

  'open both recordsets
  MsgBar "Transferring data to temporary file", False
  If Reprint = False Then
    Set recRecordset1 = newdatabase.OpenRecordset(payrollmonth)
  Else
    Set recRecordset1 = newdatabase.OpenRecordset("Reprint Table")
  End If
  Set recRecordset2 = newdatabase.OpenRecordset("temp")
  While recRecordset1.EOF = False
    recRecordset2.AddNew
    'this loop copies the data from each field to
    'the new table
'    For Each fld In recRecordset1.Fields
    recRecordset2.Fields.Refresh
    recRecordset1.Fields.Refresh
    For i = 0 To recRecordset1.Fields.count - 1
      Set fld = recRecordset1.Fields(i)
      recRecordset2(fld.Name).Value = fld.Value
    Next
    recRecordset2.Update
    recRecordset1.MoveNext
  Wend
End Function


Public Function GetMonthNum(MName As String)
    Dim Mnum As Integer
    Select Case MName
        Case "January"
            Mnum = 1
        Case "February"
            Mnum = 2
        Case "March"
            Mnum = 3
        Case "April"
            Mnum = 4
        Case "May"
            Mnum = 5
        Case "June"
            Mnum = 6
        Case "July"
            Mnum = 7
        Case "August"
            Mnum = 8
        Case "September"
            Mnum = 9
        Case "October"
            Mnum = 10
        Case "November"
            Mnum = 11
        Case "December"
            Mnum = 12
    End Select
    GetMonthNum = Mnum
End Function

Sub CollectPayrollData()
'On Error GoTo errhandle
    Dim Tax As Currency, NSSFamt As Currency
    Dim LSCamt As Currency, NHIFamt As Currency
    Dim Advanceamt As Currency
    Dim BasicSal As Currency
    'Dim gross
    Dim y As Boolean
    Dim Address 'As String
    Dim t
    Dim addr
    Dim countchr As Integer
    Dim first As Boolean
    Dim sql
    Dim incmsg
    
    Screen.MousePointer = vbHourglass
    Set datEmp = New Recordset
    datEmp.Open "SELECT * FROM EmpData WHERE [Company] =" & ID, cn, adOpenStatic, adLockOptimistic
    If datEmp.RecordCount = 0 Then
        MsgBox ("No employee data for this company." & Chr(13) & "Payroll generation aborted."), vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        If Not datEmp.EOF Then
           datEmp.MoveLast
           pcount = Fix(100 / (datEmp.RecordCount * 2))
        End If
    End If
    
    datEmp.MoveFirst
    
    'Check for basic salary for all the employees
    MsgBar ("Searching for basic salaries"), True
    Set paygen = New Recordset
    paygen.Open "SELECT * FROM PayGen WHERE Company =" & ID & " AND [Apply_Month] ='" & payrolls.Combo1 & "'" & " AND [Apply_Year] =" & CLng(payrolls.Combo2), cn, adOpenStatic, adLockOptimistic
    With paygen
        While Not datEmp.EOF
              
              .Filter = "PayrollNo='" & datEmp("PayrollNo") & "'"
              If .RecordCount <> 0 Then
                If !BasicSal = 0 Or !BasicSal = Null Then
                   incmsg = incmsg & Chr(13) & !PayrollNo
                End If
              End If
              datEmp.MoveNext
              payrolls.PBar.Value = payrolls.PBar.Value + pcount
        Wend
        If incmsg <> "" Then
           MsgBox ("Following employees do not have values for the basic salary:" & Chr(13) & incmsg & Chr(13) & "Cannot proceed with payroll generation"), vbCritical
           payrolls.PBar.Value = 0
           MsgBar "Payroll Generation Aborted!", False
           Exit Sub
        End If
        .Close
    End With
    ValidateTables
    If Error_Found = True Then
       Exit Sub
       Error_Found = False
    End If
        datEmp.MoveFirst
        Screen.MousePointer = vbHourglass
          Do While datEmp.EOF = False
                SlipOrder = 0
                LSCamt = 0
                NHIFamt = 0
                NSSFamt = 0
                DBFamt = 0
                OOIamt = 0
                HOSPamt = 0
                HOSPValue = 0
                NSSFValue = 0
                OOIValue = 0
                TaxableInc = 0
                MinTax = False
                
                Set paygen = New Recordset
                paygen.Open "select * from PayGen WHERE Company =" & ID & " AND [Apply_Month] ='" & payrolls.Combo1 & "'" & " AND [Apply_Year] =" & CLng(payrolls.Combo2) & " AND PayrollNo='" & datEmp!PayrollNo & "'", cn, adOpenStatic, adLockOptimistic
                With paygen
                     If .RecordCount = 0 Then
                        Set datExtra = New Recordset
                        datExtra.Open "SELECT * FROM Transactions", cn, adOpenStatic, adLockOptimistic
                        datExtra.AddNew
                        datExtra![Payroll_No] = datEmp!PayrollNo
                        datExtra![Apply_Month] = payrolls.Combo1
                        datExtra![Apply_Year] = CLng(payrolls.Combo2)
                        datExtra![paramcode] = "Z0001"
                        datExtra.Update
                    End If
                    .Close
                End With
                
                Set paygen = New Recordset
                paygen.Open "SELECT * FROM PayGen WHERE Company =" & ID & " AND [Apply_Month] ='" & payrolls.Combo1 & "'" & " AND [Apply_Year] =" & CLng(payrolls.Combo2) & " AND PayrollNo='" & datEmp!PayrollNo & "'", cn, adOpenStatic, adLockOptimistic
                With paygen
                If CLng(!Company) = CLng(ID) Then
                   'If !Delete <> True Then
                       
                       'MsgBar "Retrieving company details", True
                       
                       MsgBar ("Preparing payslip for ") & !Surname & Chr(32) & !Name, False
                       gross = 0
                       totaldeductions = 0
                       
                       Set datExtra = New Recordset
                       datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode = 'MinTax'", cn, adOpenStatic, adLockOptimistic
                       'If !BasicSal = 0 Then
                       '   MsgBox (("Basic Pay not found for ") & !PayrollNo & Chr(32) & !Surname & " " & !Name & Chr(13) & "Exiting Payroll Generation"), vbtbrmation
                       '   Exit Sub
                       'if !BasicSal <= datExtra!Amount Then
                       '   MinTax = True
                       'End If
                       BasicSal = !BasicSal
                       
                       Set datPayroll = New Recordset
                       datPayroll.Open "SELECT * FROM Payroll", cn, adOpenStatic, adLockOptimistic
                       
                       datPayroll.AddNew
                       datPayroll!PayrollNo = datEmp!PayrollNo
                       datPayroll![ParamName] = "BASIC PAY"
                       datPayroll![amount] = BasicSal
                       datPayroll![Month] = payrolls.Combo1
                       datPayroll![Year] = payrolls.Combo2
                       datPayroll![Category] = "Income"
                       datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                       datPayroll.Update
                       gross = gross + BasicSal
                       
                       'Look for any overtime pay
                       GetOvertime
                       
                       'Put house allowance
                       If !HouseAll <> 0 Or !HouseAll <> Null Then
                          gross = gross + !HouseAll
                          
                            datPayroll.AddNew
                            datPayroll!PayrollNo = datEmp!PayrollNo
                            datPayroll![ParamName] = "HOUSE ALLOWANCE"
                            datPayroll![amount] = !HouseAll
                            datPayroll![Month] = payrolls.Combo1
                            datPayroll![Year] = payrolls.Combo2
                            datPayroll![Category] = "Income"
                            datPayroll!SlipOrder = SlipOrder
                            SlipOrder = SlipOrder + 1
                            datPayroll.Update
                       End If
                       
                       
                       'Put other allowances
                       GetOtherAllowances
                       
                       'Put other Income
                       GetOtherIncome
                       
                       GetBenefits
                       
                       'Check for free housing
                       GetHousing
                       
                            
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "GROSS PAY:"
                        datPayroll![amount] = CCur(gross)
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Income"
                        datPayroll!SlipOrder = SlipOrder
                        SlipOrder = SlipOrder + 1
                        datPayroll.Update

                      If CCur(paygen!NonTaxable_Ref) <> 0 Then
                            datPayroll.AddNew
                            datPayroll!PayrollNo = datEmp!PayrollNo
                            datPayroll![ParamName] = "NON-TAXABLE REFUND"
                            datPayroll![amount] = CCur(paygen!NonTaxable_Ref)
                            datPayroll![Month] = payrolls.Combo1
                            datPayroll![Year] = payrolls.Combo2
                            datPayroll![Category] = "Income"
                            datPayroll!SlipOrder = SlipOrder
                            SlipOrder = SlipOrder + 1
                            datPayroll.Update
                      End If
                       
                        'Get Other Deductions
                       GetDeductions
                       
                       If paygen!Unionisable = True Then
                            Set datExtra = New Recordset
                            datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode ='UnionDues'", cn, adOpenStatic, adLockOptimistic
                            datPayroll.AddNew
                            datPayroll!PayrollNo = datEmp!PayrollNo
                            datPayroll![ParamName] = "UNION DUES"
                            datPayroll![amount] = CCur(datExtra("Amount"))
                            datPayroll![Month] = payrolls.Combo1
                            datPayroll![Year] = payrolls.Combo2
                            datPayroll![Category] = "Out"
                            datPayroll!SlipOrder = SlipOrder
                            SlipOrder = SlipOrder + 1
                            datPayroll.Update
                            
                            totaldeductions = totaldeductions + datExtra("Amount")
                       Else
                       End If
                       
                       'Get LSC
                       LSCamt = GetLSC(CCur(paygen!BasicSal))
                       
                       totaldeductions = totaldeductions + LSCamt
                       
                       'Get NHIF
                       NHIFamt = GetNHIF(CCur(paygen!BasicSal))
                       
                       totaldeductions = totaldeductions + NHIFamt
                       
                       'Get Pension data
                       NSSFamt = GetNSSF(CCur(paygen!BasicSal))
                       
                       
                       totaldeductions = totaldeductions + NSSFamt
                       
                       
                       'GetHOSP
                       HOSPamt = GetHOSP
                       totaldeductions = totaldeductions + HOSPamt
                       
                       OOIamt = GetOOI
                       totaldeductions = totaldeductions + OOIamt
                       
                       'CopyStruct ("Defined Contribution Fund")
                       
                       'GetDBFund
                       DBFamt = GetDBFund
                       totaldeductions = totaldeductions + DBFamt
                       
                       
                       '888888888888888888888888888888888888888888888888888888
                       'Get the rent
                       'If Paygen!Rent <> 0 Or Paygen!Rent <> Null Then
                       '   prn.CurrentX = 950
                       '   prn.CurrentY = posn
                       '   prn.Print "RENT"
                       '   prn.CurrentX = 3500 - Len(!Rent)
                       '   prn.CurrentY = posn
                       '   prn.Print roundit(CDbl(!Rent))
                       '   posn = posn + 200
                       '   totaldeductions = totaldeductions + !Rent
                       'End If
                       '888888888888888888888888888888888888888888888888888888
                       Set datExtra = New Recordset
                       datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode = 'MinTax'", cn, adOpenStatic, adLockOptimistic
                       TaxableInc = (gross - paygen!NonTaxable_Ref - NSSFamt - HOSPValue - OOIValue - CCur(paygen!Defined_Contr) - paygen!rent)
                       If TaxableInc < datExtra!amount Then
                          MinTax = True
                       End If

                       'Get personal relief
                       relief = 0
                       If MinTax = False Then
                            If datEmp("MarStat") = "Single" Then
                                Set datExtra = New Recordset
                               datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode ='SingleRelief'", cn, adOpenStatic, adLockOptimistic
                               relief = datExtra!amount
                            ElseIf datEmp("MarStat") = "Married" Then
                                Set datExtra = New Recordset
                               datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode ='MarriedRelief'", cn, adOpenStatic, adLockOptimistic
                               relief = datExtra!amount
                            End If
                       End If
                       
                       If MinTax = False Then
                            Tax = Computetax(CCur(TaxableInc), "1001")
                            datPayroll.AddNew
                            datPayroll!PayrollNo = datEmp!PayrollNo
                            datPayroll![ParamName] = "P.A.Y.E. Amount"
                            datPayroll![amount] = CCur(Tax - relief)
                            datPayroll![Month] = payrolls.Combo1
                            datPayroll![Year] = payrolls.Combo2
                            datPayroll![Category] = "Out"
                            datPayroll!SlipOrder = SlipOrder
                            SlipOrder = SlipOrder + 1
                            datPayroll.Update
                       Else
                            Tax = 0
                            datPayroll.AddNew
                            datPayroll!PayrollNo = datEmp!PayrollNo
                            datPayroll![ParamName] = "P.A.Y.E. Amount"
                            datPayroll![amount] = CCur(Tax)
                            datPayroll![Month] = payrolls.Combo1
                            datPayroll![Year] = payrolls.Combo2
                            datPayroll![Category] = "Out"
                            datPayroll!SlipOrder = SlipOrder
                            SlipOrder = SlipOrder + 1
                            datPayroll.Update
                       End If
                            
                       totaldeductions = totaldeductions + ((Tax) - relief)    'tb("Personal Relief"))
                       
                       'Get loans
                        GetLoans
                       
                       If paygen!Unionisable = True Then
                            Set datExtra = New Recordset
                            datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode ='UnionDues'", cn, adOpenStatic, adLockOptimistic
                            datPayroll.AddNew
                            datPayroll!PayrollNo = datEmp!PayrollNo
                            datPayroll![ParamName] = "UNION DUES"
                            datPayroll![amount] = CCur(datExtra("Amount"))
                            datPayroll![Month] = payrolls.Combo1
                            datPayroll![Year] = payrolls.Combo2
                            datPayroll![Category] = "Out"
                            datPayroll!SlipOrder = SlipOrder
                            SlipOrder = SlipOrder + 1
                            datPayroll.Update
                            
                            totaldeductions = totaldeductions + datExtra("Amount")
                       Else
                       End If
                       
                        Total = totaldeductions
                       
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "TOTAL DEDUCTIONS"
                        datPayroll![amount] = CCur(Total)
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Out"
                        datPayroll!SlipOrder = SlipOrder
                        SlipOrder = SlipOrder + 1
                        datPayroll.Update
                       
                       
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "NET PAY"
                        datPayroll![amount] = CCur((gross - Total - (benefitamt)))
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Out"
                        datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                        datPayroll.Update
                        
                        
                       Dim Pmode
                       If paygen!PayMode = "Cash" Then
                           Pmode = 1
                       ElseIf paygen!PayMode = "Cheque" Then
                           Pmode = 2
                       ElseIf paygen!PayMode = "Bank Transfer" Then
                           Pmode = 3
                       End If
                        
                        
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "PAYMENT BY"
                        datPayroll![amount] = CCur(Pmode)
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Out"
                        datPayroll!SlipOrder = SlipOrder
                        SlipOrder = SlipOrder + 1
                        datPayroll.Update
                        
                          
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "TAXABLE PAY"
                        datPayroll![amount] = CCur((gross) - paygen!NonTaxable_Ref)
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Tax"
                        datPayroll!SlipOrder = SlipOrder
                        SlipOrder = SlipOrder + 1
                        datPayroll.Update
                          
            If NSSFamt <> 0 Then
                datPayroll.AddNew
                datPayroll!PayrollNo = datEmp!PayrollNo
                datPayroll![ParamName] = "N.S.S.F."
                datPayroll![amount] = CCur(NSSFamt)
                datPayroll![Month] = payrolls.Combo1
                datPayroll![Year] = payrolls.Combo2
                datPayroll![Category] = "Tax"
                datPayroll![System] = True
                datPayroll!SlipOrder = SlipOrder
                SlipOrder = SlipOrder + 1
                datPayroll.Update
            End If
            If HOSPamt <> 0 Then
                datPayroll.AddNew
                datPayroll!PayrollNo = datEmp!PayrollNo
                datPayroll![ParamName] = "HOS Plan"
                datPayroll![amount] = CCur(HOSPamt)
                datPayroll![Month] = payrolls.Combo1
                datPayroll![Year] = payrolls.Combo2
                datPayroll![Category] = "Tax"
                datPayroll![System] = True
                datPayroll!SlipOrder = SlipOrder
                SlipOrder = SlipOrder + 1
                datPayroll.Update
            End If
            If OOIamt <> 0 Then
                datPayroll.AddNew
                datPayroll!PayrollNo = datEmp!PayrollNo
                datPayroll![ParamName] = "OWNER OCC. INTEREST"
                datPayroll![amount] = CCur(OOIamt)
                datPayroll![Month] = payrolls.Combo1
                datPayroll![Year] = payrolls.Combo2
                datPayroll![Category] = "Tax"
                datPayroll![System] = True
                datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                datPayroll.Update
            End If
            If DBFamt <> 0 Then
                datPayroll.AddNew
                datPayroll!PayrollNo = datEmp!PayrollNo
                datPayroll![ParamName] = "DEFINED BENEFIT FUND"
                datPayroll![amount] = CCur(DBFamt)
                datPayroll![Month] = payrolls.Combo1
                datPayroll![Year] = payrolls.Combo2
                datPayroll![Category] = "Tax"
                datPayroll![System] = True
                datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                datPayroll.Update
            End If
 
                        
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "NET TAXABLE"
                        datPayroll![amount] = CCur(TaxableInc)
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Tax"
                        datPayroll!SlipOrder = SlipOrder
                        SlipOrder = SlipOrder + 1
                        datPayroll.Update
                        
                   If MinTax <> True Then
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "TAXABLE-POUNDS"
                        datPayroll![amount] = CCur(Fix(pounds))
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Tax"
                        datPayroll!SlipOrder = SlipOrder
                        SlipOrder = SlipOrder + 1
                        datPayroll.Update
                   End If
                        
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "TAX CHARGED"
                        datPayroll![amount] = CCur(Tax)
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Tax"
                        datPayroll!SlipOrder = SlipOrder
                        SlipOrder = SlipOrder + 1
                        datPayroll.Update
                        
                   If MinTax <> True Then
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "PERSONAL RELIEF"
                        datPayroll![amount] = CCur(relief)
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Tax"
                        datPayroll!SlipOrder = SlipOrder
                        SlipOrder = SlipOrder + 1
                        datPayroll.Update
                   End If

                   If MinTax <> True Then
                        Dim ded
                        ded = (Tax - relief)
                        
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "TAX DEDUCTED"
                        datPayroll![amount] = CCur((Tax - relief))
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Tax"
                        datPayroll!SlipOrder = SlipOrder
                        SlipOrder = SlipOrder + 1
                        datPayroll.Update
                   Else
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = "TAX DEDUCTED"
                        datPayroll![amount] = CCur(0)
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Tax"
                        datPayroll!SlipOrder = SlipOrder
                        SlipOrder = SlipOrder + 1
                        datPayroll.Update
                   End If
                        
               Else
               End If
               datEmp.MoveNext
               payrolls.PBar.Value = payrolls.PBar.Value + pcount
            End With
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM Transactions WHERE ParamCode = 'Z0001'", cn, adOpenStatic, adLockOptimistic
            While Not datExtra.EOF
                  datExtra.Delete
                  datExtra.MoveNext
            Wend
            MinTax = False
          Loop
        
            payrolls.PBar.Value = 100
            MsgBar "Payroll generation complete", False
            Screen.MousePointer = vbDefault
            'MsgBox ("Payroll Generation successfully completed."), vbInformation
            'payrolls.percentage.Caption = payrolls.PBar.Value & "%"
            payrolls.PBar.Value = 0
            MsgBar "", False


End Sub

Function Computetax(Income As Currency, staff_code As String)

'If payrolls.Data1.Recordset("tax") = True Then
    'This procedure computes the tax an employee is suppose to pay
    'on his income
    'Dim
    Dim Dedinc As Long
    Dim Tax, TaxAmt
    Dim PreviousUpper
    TaxAmt = 0
    PreviousUpper = 0
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM General_SysInfo WHERE ParamCode ='POUNDVALUE'", cn, adOpenStatic, adLockOptimistic
    
    pounds = (Income / datExtra("Amount"))
    pounds = Fix(pounds)
    
     'Store the taxable in pounds
     Set datExtra = New Recordset
     datExtra.Open "SELECT TOP 1 * FROM Payroll_results WHERE wage_code ='TPNDS' " & _
     " AND Staff_Code ='" & staff_code & "' ORDER BY staff_code", cn, adOpenStatic, adLockOptimistic
     datExtra.AddNew
     datExtra!wage_code = "TPNDS"
     datExtra!staff_code = staff_code
     datExtra!amount = pounds
     datExtra!co_code = ID
     datExtra!Month = Period_Month
     datExtra!Year = Period_Year
     datExtra.Update
    
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM TaxTable", cn, adOpenStatic, adLockOptimistic
    With datExtra
        .MoveFirst
        
        'Test code for the computation of tax lower than
        'the first upper limit.
        Dedinc = pounds
        'End of test code.
        
        Do While Dedinc >= 0
           If pounds > !Upper Then
              Tax = (!Upper - PreviousUpper) * !amount
              PreviousUpper = !Upper
              Dedinc = pounds - !Upper
              .MoveNext
           Else
              Tax = Dedinc * !amount
              Dedinc = -1
           End If
           TaxAmt = TaxAmt + Tax
        Loop
    End With
    Computetax = TaxAmt
End Function

Sub GetAdvances()
    Dim Counter%
    Dim Cumulative As Long
    Cumulative = 0
    
        If Advances.RecordCount <> 0 Then
           criteria = "SELECT * FROM Advances WHERE [Apply_Month] = " & paymonth & " AND [Apply_year] = " & payrolls.Combo2
           Set Dn = perioddb.CreateDynaset(criteria)
           'criteria = "[Paynum] =" & "'" & Empdata("PayrollNo") & "'"
           Dn.Filter = criteria
           Set Dn = Dn.CreateDynaset()
            
           If Not Dn.RecordCount = 0 Then
              Dn.MoveLast
              Dn.MoveFirst
              For Counter = Counter To (Dn.RecordCount - 1)
                  Cumulative = Cumulative + Dn("Amount")
                  Dn.MoveNext
              Next Counter
              prn.CurrentX = 950
              prn.CurrentY = posn
              prn.Print "ADVANCES"
              prn.CurrentX = 3500 - Len(Cumulative)
              prn.CurrentY = posn
              prn.Print Roundit(CDbl(Cumulative))
              posn = posn + 200
              'CopyStruct ("Advances")
              totaldeductions = totaldeductions + Cumulative
           End If
        Else
           
        End If
    'tb("Advances") = Cumulative

End Sub

Sub GetBenefits()
    Dim Counter%
    Dim Cumulative As Long
    Cumulative = 0
    benefitamt = 0
    
        'Set datExtra = New Recordset
        'datExtra.Open "SELECT * FROM PayGen WHERE Company =" & ID & " AND [Apply_Month] ='" & payrolls.Combo1 & "'" & " AND [Apply_Year] =" & CLng(payrolls.Combo2) & " AND PayrollNo='" & datEmp!PayrollNo & "'", cn, adopenstatic, adLockOptimistic
        With paygen
            .Filter = "ParamCode LIKE 'B*'"
           If Not .RecordCount = 0 Then
              .MoveLast
              .MoveFirst
              For Counter = Counter To (.RecordCount - 1)
                    datPayroll.AddNew
                    datPayroll!PayrollNo = datEmp!PayrollNo
                    datPayroll![ParamName] = UCase(!Description)
                    datPayroll![amount] = CCur(!amount)
                    datPayroll![Month] = payrolls.Combo1
                    datPayroll![Year] = payrolls.Combo2
                    datPayroll![Category] = "Income"
                    datPayroll!SlipOrder = SlipOrder
                    SlipOrder = SlipOrder + 1
                    datPayroll.Update
                  Cumulative = Cumulative + !amount
                  gross = gross + !amount
                  benefitamt = benefitamt + !amount
                  .MoveNext
              Next Counter
           End If
           .Filter = ""
        End With

End Sub

Sub GetDeductions()
    Dim Counter%
    Dim Cumulative As Long
    Cumulative = 0
    
        'Set datExtra = New Recordset
        'datExtra.Open "SELECT * FROM PayGen WHERE Company =" & ID & " AND [Apply_Month] ='" & payrolls.Combo1 & "'" & " AND [Apply_Year] =" & CLng(payrolls.Combo2) & " AND PayrollNo='" & datEmp!PayrollNo & "'", cn, adopenstatic, adLockOptimistic
        With paygen
            .Filter = "ParamCode LIKE 'D*'"
           If Not .RecordCount = 0 Then
              .MoveLast
              .MoveFirst
              For Counter = Counter To (.RecordCount - 1)
                datPayroll.AddNew
                datPayroll!PayrollNo = datEmp!PayrollNo
                datPayroll![ParamName] = UCase(!Description)
                datPayroll![amount] = CCur(!amount)
                datPayroll![Month] = payrolls.Combo1
                datPayroll![Year] = payrolls.Combo2
                datPayroll![Category] = "Out"
                datPayroll!SlipOrder = SlipOrder
                SlipOrder = SlipOrder + 1
                datPayroll.Update
                  
                  Cumulative = Cumulative + !amount
                  totaldeductions = totaldeductions + !amount
                  .MoveNext
              Next Counter
           End If
           .Filter = ""
        End With
End Sub

Sub GetLoans()
    'Dim Dn As Dynaset
    'Dim Ds As Dynaset
    Dim Counter%
    Dim Cumulative As Long
    Cumulative = 0
    Counter = 1
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM Loans WHERE InstallNo<>InstallPaid AND PayrollNo ='" & datEmp!PayrollNo & "'", cn, adOpenStatic, adLockOptimistic
    With datExtra
        If .RecordCount <> 0 Then

           If !InstallPaid <> !Time * 12 Then
              'Dn.MoveLast
              'Dn.MoveFirst
              For Counter = Counter To (.RecordCount)
                  'If Dn("InstallPaid") <> Dn("InstallNo") Then
                     If !InstallPaid = 0 Then
                        If !Apply_Month = payrolls.Combo1 And !Apply_Year = CInt(payrolls.Combo2) Then
                            Cumulative = Cumulative + !InstallAmt
                            'Dn.Edit
                            !InstallPaid = !InstallPaid + 1
                            
                            datPayroll.AddNew
                            datPayroll!PayrollNo = datEmp!PayrollNo
                            datPayroll![ParamName] = UCase(!Description)
                            datPayroll![amount] = CCur(!InstallAmt)
                            datPayroll![Month] = payrolls.Combo1
                            datPayroll![Year] = payrolls.Combo2
                            datPayroll![Category] = "Out"
                            datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                            datPayroll.Update
                            
                            
                            totaldeductions = totaldeductions + !InstallAmt
                        Else
                        End If
                     Else
                        Cumulative = Cumulative + !InstallAmt
                        'Dn.Edit
                        !InstallPaid = !InstallPaid + 1
                        
                        datPayroll.AddNew
                        datPayroll!PayrollNo = datEmp!PayrollNo
                        datPayroll![ParamName] = UCase(!Description)
                        datPayroll![amount] = CCur(!InstallAmt)
                        datPayroll![Month] = payrolls.Combo1
                        datPayroll![Year] = payrolls.Combo2
                        datPayroll![Category] = "Out"
                        datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                        datPayroll.Update
                        
                        totaldeductions = totaldeductions + !InstallAmt
                    End If
                  'Else
                  'End If
                  .MoveNext
              Next Counter
           Else
           End If
           
        Else
        End If
    End With
End Sub

Function GetLSC(Income As Currency)
    'If payrolls.Data1.Recordset("LASC") = True Then
    'With paygen
        'If !LSC = True Then
             Set datExtra = New Recordset
             datExtra.Open "SELECT * FROM Servchrg", cn, adOpenStatic, adLockOptimistic
             With datExtra
                If .RecordCount = 0 Then
                   MsgBox ("LSC data missing, cannot proceed"), 16
                   Exit Function
                Else
                   Do While Not .EOF
                       If Income < !Upper Then
                          GetLSC = !Charge
                          Exit Function
                       Else
                          .MoveNext
                       End If
                   Loop
                End If
            End With
        'Else
        'End If
    'End With
End Function


Function GetNHIF(Income As Currency)
    'If payrolls.Data1.Recordset("NHIF") = True Then
    'Set datExtra = New Recordset
    'datExtra.Open "SELECT * FROM NHIFFile", cn, adopenstatic, adLockOptimistic
    '     If !NHIF = True Then
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM NHIFfile", cn, adOpenStatic, adLockOptimistic
            With datExtra
                 If .RecordCount = 0 Then
                    MsgBox ("NHIF data missing, cannot proceed"), 16
                    Exit Function
                 Else
                    Do While Not .EOF
                        If Income < !Upper Then
                           GetNHIF = !Charge
                           Exit Function
                        Else
                           .MoveNext
                        End If
                    Loop
                 End If
            End With
        'Else
        'End If
    'End With
End Function

Function GetNSSF(Income As Currency)
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM General_SysInfo WHERE ParamCode = 'NSSF'", cn, adOpenStatic, adLockOptimistic
    With paygen
        If !NSSF = True Then
            GetNSSF = datExtra!amount
            NSSFValue = GetNSSF
        Else
        End If
    End With
End Function

Sub GetOtherAllowances()
    Dim Counter%
    Dim Cumulative As Long
    Cumulative = 0
        Set prn = Printer
        
        'Set datExtra = New Recordset
        'datExtra.Open "SELECT * FROM PayGen WHERE Company =" & ID & " AND [Apply_Month] ='" & payrolls.Combo1 & "'" & " AND [Apply_Year] =" & CLng(payrolls.Combo2) & " AND PayrollNo='" & datEmp!PayrollNo & "'", cn, adopenstatic, adLockOptimistic
           With paygen
               .Filter = "ParamCode LIKE 'A*'"
                If Not .RecordCount = 0 Then
                   .MoveLast
                   .MoveFirst
                   For Counter = Counter To (.RecordCount - 1)
                       datPayroll.AddNew
                       datPayroll!PayrollNo = datEmp!PayrollNo
                       datPayroll![ParamName] = UCase(!Description)
                       datPayroll![amount] = CCur(!amount)
                       datPayroll![Month] = payrolls.Combo1
                       datPayroll![Year] = payrolls.Combo2
                       datPayroll![Category] = "Income"
                       datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                       datPayroll.Update
                       
                       Cumulative = Cumulative + !amount
                       gross = gross + !amount
                       .MoveNext
                   Next Counter
                End If
                .Filter = ""
           End With
End Sub

Sub GetOtherIncome()
    Dim Counter%
    Dim Cumulative As Long
    Cumulative = 0
    
        'Set datExtra = New Recordset
        'datExtra.Open "SELECT * FROM PayGen WHERE Company =" & ID & " AND [Apply_Month] ='" & payrolls.Combo1 & "'" & " AND [Apply_Year] =" & CLng(payrolls.Combo2) & " AND PayrollNo='" & datEmp!PayrollNo & "'", cn, adopenstatic, adLockOptimistic
        With paygen
            .Filter = "ParamCode LIKE 'I*'"
           If Not .RecordCount = 0 Then
              .MoveLast
              .MoveFirst
              For Counter = Counter To (.RecordCount - 1)
                    datPayroll.AddNew
                    datPayroll!PayrollNo = datEmp!PayrollNo
                    datPayroll![ParamName] = UCase(!Description)
                    datPayroll![amount] = CCur(!amount)
                    datPayroll![Month] = payrolls.Combo1
                    datPayroll![Year] = payrolls.Combo2
                    datPayroll![Category] = "Income"
                    datPayroll!SlipOrder = SlipOrder
                    SlipOrder = SlipOrder + 1
                    datPayroll.Update
                  
                  Cumulative = Cumulative + !amount
                  gross = gross + !amount
                  .MoveNext
              Next Counter
           End If
           .Filter = ""
        End With

End Sub

Sub GetOvertime()
    Dim Counter%
    Dim Cumulative As Long
    Dim Hrs As Long
    Cumulative = 0
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM OverTime WHERE [Apply_Month] = '" & payrolls.Combo1 & "'" & " AND [Apply_year] = " & payrolls.Combo2 & " AND [PayrollNo] ='" & datEmp!PayrollNo & "'", cn, adOpenStatic, adLockOptimistic
    With datExtra
        If .RecordCount <> 0 Then
            For Counter = Counter To (.RecordCount - 1)
               If Not !Normalhrs = "" Then
                  Cumulative = Cumulative + (!Normalhrs * !HourRate)
                  Hrs = Hrs + !Normalhrs
               End If
               If Not !Halfhrs = "" Then
                  Cumulative = Cumulative + ((!Halfhrs * 1.5) * !HourRate)
                  Hrs = Hrs + !Halfhrs
               End If
               If Not !twicehrs = "" Then
                  Cumulative = Cumulative + ((!twicehrs * 2) * !HourRate)
                  Hrs = Hrs + !twicehrs
               End If
               .MoveNext
            Next Counter
            datPayroll.AddNew
            datPayroll!PayrollNo = datEmp!PayrollNo
            datPayroll![ParamName] = UCase("OVERTIME PAY")
            datPayroll![amount] = CCur(Cumulative)
            datPayroll![Month] = payrolls.Combo1
            datPayroll![Year] = payrolls.Combo2
            datPayroll![Category] = "Income"
            datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
            datPayroll.Update

        End If
    End With
        
        gross = gross + Cumulative
End Sub

Sub MsgBar(Msg As String, pw As Integer)
  If Msg = "" Then
    frmMain.StatusBar1.Panels(2) = "Ready..."
  Else
    If pw = True Then
      frmMain.StatusBar1.Panels(2) = Msg + ", please wait..."
    Else
      frmMain.StatusBar1.Panels(2) = Msg
    End If
  End If
  'frmmain.StatusBar1.Panels(2).Refresh
End Sub

Sub OpenFiles()

End Sub

Sub Outlines(formname As Form)
    Dim drkgray As Long, fullwhite As Long
    Dim i As Integer
    Dim ctop As Integer, cleft As Integer, cright As Integer, cbottom As Integer

    ' Outline a form's controls for 3D look unless control's TAG
    ' property is set to "skip".

    Dim cname As Control
    drkgray = RGB(128, 128, 128)
    fullwhite = RGB(255, 255, 255)

    For i = 0 To (formname.Controls.count - 1)
        Set cname = formname.Controls(i)
        If TypeOf cname Is Menu Then
            'Debug.Print "menu item"
        ElseIf (UCase(cname.Tag) = "OL") Then
                ctop = cname.Top - Screen.TwipsPerPixelY
                cleft = cname.Left - Screen.TwipsPerPixelX
                cright = cname.Left + cname.Width
                cbottom = cname.Top + cname.Height
                formname.Line (cleft, ctop)-(cright, ctop), drkgray
                formname.Line (cleft, ctop)-(cleft, cbottom), drkgray
                formname.Line (cleft, cbottom)-(cright, cbottom), fullwhite
                formname.Line (cright, ctop)-(cright, cbottom), fullwhite
        End If
    Next i
End Sub

Sub PicOutlines(pic As Control, ctl As Control)
    Dim drkgray As Long, fullwhite As Long
    Dim ctop As Integer, cleft As Integer, cright As Integer, cbottom As Integer

    ' Outline a form's controls for 3D look unless control's TAG
    ' property is set to "skip".

    Dim cname As Control
    drkgray = RGB(128, 128, 128)
    fullwhite = RGB(255, 255, 255)

    ctop = ctl.Top - Screen.TwipsPerPixelY
    cleft = ctl.Left - Screen.TwipsPerPixelX
    cright = ctl.Left + ctl.Width
    cbottom = ctl.Top + ctl.Height
    pic.Line (cleft, ctop)-(cright, ctop), drkgray
    pic.Line (cleft, ctop)-(cleft, cbottom), drkgray
    pic.Line (cleft, cbottom)-(cright, cbottom), fullwhite
    pic.Line (cright, ctop)-(cright, cbottom), fullwhite

End Sub

Sub UpdateInctables()
    'This procedure updates other tables of income when an
    'employee is registered.

     'Update the income data
     'IncData("PayrollNo") = Trim(frmEmp.txtPayroll)
     'IncData("Basic Sal") = CCur(0)
     'IncData("House All") = CCur(0)
     'IncData("Leave All") = CCur(0)
     'IncData("Non-taxable Ref") = CCur(0)
     'IncData("CCode") = ID
     'If frmEmp.cash.Value = True Then
     '   IncData("Paymode") = "Cash"
     'ElseIf frmEmp.cheque.Value = True Then
     '   IncData("Paymode") = "Cheque"
     'Else
     '   IncData("Paymode") = "Bank Transfer"
     'End If
     'IncData.Update
     'frmEmp.Data3.Recordset.AddNew
     'frmEmp.Data3.Recordset("PayrollNo") = Trim(frmEmp.txtPayroll)
     'frmEmp.Data3.Recordset("Tax") = True
     'frmEmp.Data3.Recordset.Update
End Sub


Public Function GetHOSP()
    Dim x
    With paygen
        If Trim(!HOSP) <> "0" Then
           If !HOSP <> "(Maximum)" Then
               GetHOSP = !HOSP
               HOSPValue = !HOSP
                
                datPayroll.AddNew
                datPayroll!PayrollNo = datEmp!PayrollNo
                datPayroll![ParamName] = "HOS Plan"
                datPayroll![amount] = CCur(HOSPValue)
                datPayroll![Month] = payrolls.Combo1
                datPayroll![Year] = payrolls.Combo2
                datPayroll![Category] = "Out"
                datPayroll!SlipOrder = SlipOrder
                SlipOrder = SlipOrder + 1
                datPayroll.Update
                
           Else  'HOSP is maximum
                Set datExtra = New Recordset
                datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode ='MaxHOSP'", cn, adOpenStatic, adLockOptimistic
                x = (datExtra("Amount") / 12)
                x = Fix(x)
                
                If payrolls.Combo1 <> "December" Then 'InStr(1, ((!No_of_HOSP + 1) / 12), ".") <> 0 Then
                   'MsgBox ("Value is a whole number")
                   GetHOSP = x
                   HOSPValue = x
                Else
                    'MsgBox ("Value is a fraction.")
                    x = Fix(datExtra("Amount") - (x * 11))
                   GetHOSP = x
                   HOSPValue = x
                End If
                
                datPayroll.AddNew
                datPayroll!PayrollNo = datEmp!PayrollNo
                datPayroll![ParamName] = "HOS Plan"
                datPayroll![amount] = CCur(HOSPValue)
                datPayroll![Month] = payrolls.Combo1
                datPayroll![Year] = payrolls.Combo2
                datPayroll![Category] = "Out"
                datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                datPayroll.Update

           End If
           '!No_of_HOSP = !No_of_HOSP + 1
           
           ''**********Removed because of above modification. Consider Physical month of December
           'frmmain.datExtra.RecordSource = "SELECT * FROM IncData WHERE PayrollNo='" & !PayrollNo & "'"
           'frmmain.datExtra.Refresh
           'datExtra!No_of_HOSP = !No_of_HOSP + 1
           'datExtra.Update
           
        Else
            GetHOSP = CCur(0)
        End If
    End With
End Function

Public Function GetOOI()
    Dim x
    With paygen
        If Trim(!OOInt) <> "0" Then
           If !OOInt <> "(Maximum)" Then
                GetOOI = !OOInt
                OOIValue = !OOInt
                datPayroll.AddNew
                datPayroll!PayrollNo = datEmp!PayrollNo
                datPayroll![ParamName] = "OWNER OCC. INTEREST"
                datPayroll![amount] = CCur(OOIValue)
                datPayroll![Month] = payrolls.Combo1
                datPayroll![Year] = payrolls.Combo2
                datPayroll![Category] = "Out"
                datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                datPayroll.Update

           Else  'HOSP=0
                Set datExtra = New Recordset
                datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode ='MaxOOI'", cn, adOpenStatic, adLockOptimistic
                x = (datExtra("Amount") / 12)
                x = Fix(x)
                
                If payrolls.Combo1 <> "December" Then 'InStr(1, ((!No_of_OOI + 1) / 12), ".") <> 0 Then
                   'MsgBox ("Value is a whole number")
                   GetOOI = x
                   OOIValue = x
                Else
                    'MsgBox ("Value is a fraction.")
                    x = Fix(datExtra("Amount") - (x * 11))
                   GetOOI = x
                   OOIValue = x
                End If
                
                datPayroll.AddNew
                datPayroll!PayrollNo = datEmp!PayrollNo
                datPayroll![ParamName] = "OWNER OCC. INTEREST"
                datPayroll![amount] = CCur(OOIValue)
                datPayroll![Month] = payrolls.Combo1
                datPayroll![Year] = payrolls.Combo2
                datPayroll![Category] = "Out"
                datPayroll!SlipOrder = SlipOrder
                SlipOrder = SlipOrder + 1
                datPayroll.Update
                
            End If
           ' !No_of_OOI = !No_of_OOI + 1
           
           ''********Removed due to the above modification. Consider Physical month of December
           'frmmain.datExtra.RecordSource = "SELECT * FROM IncData WHERE PayrollNo='" & !PayrollNo & "'"
           'frmmain.datExtra.Refresh
           'datExtra!No_of_OOI = !No_of_OOI + 1
           'datExtra.Update
        Else
            GetOOI = CCur(0)
        End If
    End With
End Function

Public Function GetDBFund()
    With paygen
        If CCur(!Defined_Contr) <> 0 Then
            Dim A, B, C
            
            'A = Fix((30 / 100) * ((!BasicSal + tb("Overtime Pay") + ![House Allowance] + tb("Other Allowances") + tb("Other Income") + tb("Medical Refund") + tb("Benefits")) - !NonTaxable Ref))
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode ='PENSPERCENT'", cn, adOpenStatic, adLockOptimistic
            A = Fix((datExtra!amount / 100) * ((gross - HousingValue) - !NonTaxable_Ref))
            B = CCur(!Defined_Contr)
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode ='MaxPenc'", cn, adOpenStatic, adLockOptimistic
            C = datExtra!amount
            
            If A < B And A < C Then
                GetDBFund = A
            ElseIf B < A And B < C Then
                GetDBFund = B
            ElseIf C < A And C < B Then
                GetDBFund = C
            End If
        Else
            GetDBFund = CCur(0)
            Exit Function
        End If
        If C <> Empty Then
            If !NSSF <> False Then
                If (CCur(NSSFValue + (CCur(GetDBFund)))) > C Then
                End If
            End If
        End If
    End With
    
        datPayroll.AddNew
        datPayroll!PayrollNo = datEmp!PayrollNo
        datPayroll![ParamName] = "DEFINED BENEFIT FUND"
        datPayroll![amount] = CCur(GetDBFund)
        datPayroll![Month] = payrolls.Combo1
        datPayroll![Year] = payrolls.Combo2
        datPayroll![Category] = "Out"
        datPayroll!SlipOrder = SlipOrder
        SlipOrder = SlipOrder + 1
        datPayroll.Update
End Function

Public Sub GetHousing()
        Dim y
        Dim rent
        HousingValue = 0
        With paygen
            Set datExtra = New Recordset
            datExtra.Open "SELECT * FROM General_SysInfo WHERE ParamCode ='MaxHousingSal'", cn, adOpenStatic, adLockOptimistic
            If !Housing = True Then
                'Check if the gross of salary is more than the maximum salary
                'considered in housing benefit
                If !BasicSal >= CCur(datExtra!amount) Then
                   'If the gross of salary is more, then we only consider the rent
                   'charged by the employer and the market rental value, whichever is
                   'the higher of the two
                Else
                   'Now we look at the recommended percentages
                   Set datExtra = New Recordset
                   datExtra.Open "SELECT * FROM General_SysInfo WHERE ParamCode ='HousingPercent'", cn, adOpenStatic, adLockOptimistic
                   y = ((datExtra!amount) / 100) * (!BasicSal)
                   If y > !Quarters Then  'Percentage is more than the market
                                          'rental value so we take y
                        
                       datPayroll.AddNew
                       datPayroll!PayrollNo = datEmp!PayrollNo
                       datPayroll![ParamName] = "HOUSING (@" & datExtra!amount & "%)"
                       datPayroll![amount] = CCur(y)
                       datPayroll![Month] = payrolls.Combo1
                       datPayroll![Year] = payrolls.Combo2
                       datPayroll![Category] = "Income"
                       datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                       datPayroll.Update
                       
                       datPayroll.AddNew
                       datPayroll!PayrollNo = datEmp!PayrollNo
                       datPayroll![ParamName] = "RENT"
                       datPayroll![amount] = !rent
                       datPayroll![Month] = payrolls.Combo1
                       datPayroll![Year] = payrolls.Combo2
                       datPayroll![Category] = "Out"
                       datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                       datPayroll.Update
                       rent = !rent

                      
                   Else
                       datPayroll.AddNew
                       datPayroll!PayrollNo = datEmp!PayrollNo
                       datPayroll![ParamName] = "HOUSING (Value of Quarters)"
                       datPayroll![amount] = !Quarters
                       datPayroll![Month] = payrolls.Combo1
                       datPayroll![Year] = payrolls.Combo2
                       datPayroll![Category] = "Income"
                       datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                       datPayroll.Update
                       y = !Quarters
                       
                       datPayroll.AddNew
                       datPayroll!PayrollNo = datEmp!PayrollNo
                       datPayroll![ParamName] = "RENT"
                       datPayroll![amount] = !rent
                       datPayroll![Month] = payrolls.Combo1
                       datPayroll![Year] = payrolls.Combo2
                       datPayroll![Category] = "Out"
                       datPayroll!SlipOrder = SlipOrder
                       SlipOrder = SlipOrder + 1
                       datPayroll.Update
                       rent = !rent
                     
                      
                      
                   End If
                'Else
                     Set datExtra = New Recordset
                     datExtra.Open "SELECT * FROM [General SysInfo] WHERE ParamCode ='HousingPercent'", cn, adOpenStatic, adLockOptimistic
                     datPayroll.AddNew
                     datPayroll!PayrollNo = datEmp!PayrollNo
                     datPayroll![ParamName] = "HOUSING (@" & datExtra!amount & "%)"
                     datPayroll![amount] = (((datExtra!amount) / 100) * gross)
                     datPayroll![Month] = payrolls.Combo1
                     datPayroll![Year] = payrolls.Combo2
                     datPayroll![Category] = "Income"
                     datPayroll!SlipOrder = SlipOrder
                     SlipOrder = SlipOrder + 1
                     datPayroll.Update
                     y = (((datExtra!amount) / 100) * gross)
                     
                     datPayroll.AddNew
                     datPayroll!PayrollNo = datEmp!PayrollNo
                     datPayroll![ParamName] = "RENT"
                     datPayroll![amount] = !rent
                     datPayroll![Month] = payrolls.Combo1
                     datPayroll![Year] = payrolls.Combo2
                     datPayroll![Category] = "Out"
                     datPayroll!SlipOrder = SlipOrder
                     SlipOrder = SlipOrder + 1
                     datPayroll.Update
                     rent = !rent
                    
                    
                End If
                HousingValue = y
                gross = gross + y
                totaldeductions = totaldeductions + rent
            Else
            End If
        End With
End Sub

Public Function Roundit(Fig As Double)
    If Fig <> 0 Then
        Fig = Round(Fig, 2)
        lenvar = Fig
        z = InStr(1, Fig, ".")
        If z <> 0 Then
           y = Right(Fig, (Len(lenvar)) - (z))
           If Len(y) = 2 Then
              y = Right(y, 1)
                If y <> 0 And y <> 5 Then
                   If y < 5 Then
                      If y < 2.5 Then
                         Fig = Left(Fig, Len(lenvar) - 1)
                      Else
                         Fig = Left(Fig, Len(lenvar) - 1) & CCur(5)
                      End If
                  Else
                      If y < 7.5 Then
                         Fig = Left(Fig, Len(lenvar) - 1) & CCur(5)
                      Else
                         Fig = Left(Fig, Len(lenvar) - 1) + 0.1
                      End If
                  End If
                End If
           End If
        End If
    Else
    End If
    Roundit = Fig
End Function

Public Sub CapturePic()
' Set CancelError is True
    frmEmp.CommonDialog1.CancelError = True
    On Error GoTo errhandler
    ' Set flags
    'CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    frmEmp.CommonDialog1.Filter = "All Picture Files|*.bmp;*.jpg;*.dib;*.gif;*.wmf;*.emf;*.ico;*.cur|Bitmap Files" & _
    "(*.bmp;*.dib)|*.bmp;*.dib|JPEG Images (*.jpg)|*.jpg|GIF Images (*.gif)|*.gif|Metafiles (*.wmf;*.emf)|*.wmf;*.emf|Icons (*.ico;*.cur)|*.ico;*.cur|All Files(*.*)|*.*"
    ' Specify default filter
    frmEmp.CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    frmEmp.CommonDialog1.ShowOpen
    ' Display name of selected file

    frmEmp.txtField(21) = frmEmp.CommonDialog1.FileName
    frmEmp.Picture1.Picture = LoadPicture(frmEmp.CommonDialog1.FileName) '("C:\program files\devstudio\vb\graphics\metafile\business\calendar.wmf")
    Exit Sub
    
errhandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Public Sub ConvertMonth()
      If monthp = "January" Then
         MonthInt = 1
      ElseIf monthp = "February" Then
         MonthInt = 2
      ElseIf monthp = "March" Then
         MonthInt = 3
      ElseIf monthp = "April" Then
         MonthInt = 4
      ElseIf monthp = "May" Then
         MonthInt = 5
      ElseIf monthp = "June" Then
         MonthInt = 6
      ElseIf monthp = "July" Then
         MonthInt = 7
      ElseIf monthp = "August" Then
         MonthInt = 8
      ElseIf monthp = "September" Then
         MonthInt = 9
      ElseIf monthp = "October" Then
         MonthInt = 10
      ElseIf monthp = "November" Then
         MonthInt = 11
      ElseIf monthp = "December" Then
         MonthInt = 12
      End If
      'ConvertMonth(monthp) = monthint
End Sub

Public Sub ValidateTables()
    MsgBar "Validating System Tables", True
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM TaxTable", cn, adOpenStatic, adLockOptimistic
    With datExtra
         If .RecordCount <> 0 Then
            If !lower <> CCur(1) Then
               MsgBox ("An error has been detected in the tax table." & Chr(13) & "Starting value must be 1" & Chr(13) & "Process Halted!"), vbCritical, "Validation Error"
               Screen.MousePointer = vbDefault
               Error_Found = True
               Exit Sub
            Else
                prev = !Upper
                .MoveNext
                While Not .EOF
                      If !lower <> (prev + 1) Then
                         MsgBox ("An error has been detected in the tax table." & Chr(13) & "Descrepancy detected in the transition between rows  " & .AbsolutePosition - 1 & " and " & .AbsolutePosition & Chr(13) & "Process Halted!"), vbCritical
                         Screen.MousePointer = vbDefault
                         Error_Found = True
                         Exit Sub
                      End If
                      prev = !Upper
                      .MoveNext
                Wend
            End If
        End If
        .Close
    End With
    
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM ServChrg", cn, adOpenStatic, adLockOptimistic
    With datExtra
         If .RecordCount <> 0 Then
            If !lower <> CCur(1) Then
               MsgBox ("An error has been detected in the Service Charge table." & Chr(13) & "Starting value must be 1" & Chr(13) & "Process Halted!"), vbCritical, "Validation Error"
               Screen.MousePointer = vbDefault
               Error_Found = True
               Exit Sub
            Else
                prev = !Upper
                .MoveNext
                While Not .EOF
                      If !lower <> (prev + 1) Then
                         MsgBox ("An error has been detected in the Service Charge table." & Chr(13) & "Descrepancy detected in the transition between rows  " & .AbsolutePosition - 1 & " and " & .AbsolutePosition & Chr(13) & "Process Halted!"), vbCritical
                         Screen.MousePointer = vbDefault
                         Error_Found = True
                         Exit Sub
                      End If
                      prev = !Upper
                      .MoveNext
                Wend
            End If
        End If
        .Close
    End With
    
    Set datExtra = New Recordset
    datExtra.Open "SELECT * FROM NHIFFile", cn, adOpenStatic, adLockOptimistic
    With datExtra
         If .RecordCount <> 0 Then
            If !lower <> CCur(1) Then
               MsgBox ("An error has been detected in the NHIF table." & Chr(13) & "Starting value must be 1" & Chr(13) & "Process Halted!"), vbCritical, "Validation Error"
               Screen.MousePointer = vbDefault
               Error_Found = True
               Exit Sub
            Else
                prev = !Upper
                .MoveNext
                While Not .EOF
                      If !lower <> (prev + 1) Then
                         MsgBox ("An error has been detected in the NHIF table." & Chr(13) & "Descrepancy detected in the transition between rows  " & .AbsolutePosition - 1 & " and " & .AbsolutePosition & Chr(13) & "Process Halted!"), vbCritical
                         Screen.MousePointer = vbDefault
                         Error_Found = True
                         Exit Sub
                      End If
                      prev = !Upper
                      .MoveNext
                Wend
            End If
        End If
        .Close
    End With
End Sub

Public Sub Main()
    'Procedure executed first to open the database and initialize it
    'Set cn = New Connection
    'cn.CursorLocation = adUseClient
    'cn.Open "PROVIDER=MSDASQL;dsn=PMngr;uid=Admin;pwd=soni;"
    
    'Call PackTables
    'frmEntry.Show vbModal
    frmSplash.Show
End Sub

Public Sub PackTables()
    'This procedure only applies to FoxPro Table and is designed to clear the tables
    'of any deleted data
    'An error capture is included incase the table is not opened exlusively so as to
    'ensure continued program execution
    
On Error Resume Next
    'Clear the tables of any deleted data
    cn.Execute "PACK Personal_data"
    cn.Execute "PACK Employment_data"
    cn.Execute "PACK bank_data"
    cn.Execute "PACK Basic_Pay_data"
    cn.Execute "PACK Family_data"
    cn.Execute "PACK Loans_data"
    cn.Execute "PACK Loan_Schedule"
    cn.Execute "PACK Current_Periods"
    cn.Execute "PACK Periods_History"
    cn.Execute "PACK Payroll_Results"
    cn.Execute "PACK Transactions"
    cn.Execute "PACK Company_data"
    cn.Execute "PACK Overtime"
    cn.Execute "PACK Pay_Summary"
    cn.Execute "PACK Reports"
    cn.Execute "PACK ULogins"
    cn.Execute "PACK elec_ht"
    cn.Execute "PACK elec_data"
    cn.Execute "PACK banks_information"
    cn.Execute "PACK bank_branches"
    'cn.Execute "PACK elec_data"
End Sub


Public Function GetMonths(StartMonth As Integer, EndMonth As Integer, FieldName As String)
    'Called by report viewer to filter the reports for particular months
    'This function gets the months required, converts them to physical months
    'and passes the required months to the report engine
    'Author: Patrick Odawo
    'Date:   29th/Nov/2000
    
    Dim xCount As Integer
    Dim MStr As String
    For xCount = StartMonth To EndMonth
        If MStr = "" Then
           MStr = FieldName & "='" & MonthName(xCount) & "'"
        Else
           MStr = MStr & " AND " & FieldName & "='" & MonthName(xCount) & "'"
        End If
    Next
    GetMonths = MStr
    
End Function

Public Sub GetSummary(vStaff_Code As String, vMonth As String, vYear As Integer)
    'This procedure has been written to retrieve all payroll results,
    'summarize them and send the data to a summary report
    Dim vTotal As Currency
    Dim datSum As Recordset
    
    Set datExtra = New Recordset
    datExtra.Open "SELECT Payroll_Results.*,Wage_Types.* FROM Payroll_Results INNER JOIN " & _
    " Wage_Types ON Payroll_Results.Wage_code = Wage_Types.Wage_Code WHERE Staff_code ='" & _
    vStaff_Code & "' AND co_code ='" & ID & "' AND Month ='" & _
    vMonth & "' AND Year =" & vYear, cn, adOpenStatic, adLockOptimistic
    
    If frmPSummary.prg.Value + prgv < 100 Then
       frmPSummary.prg.Value = frmPSummary.prg.Value + prgv
    Else
        frmPSummary.prg.Value = 100
    End If
    
    
    If datExtra.RecordCount <> 0 Then  'Employee's pay was generated during the month
        Set datExtra2 = New Recordset
        datExtra2.Open "SELECT TOP 1 * FROM Pay_Summary ORDER BY staff_code", cn, adOpenStatic, adLockOptimistic
        cn.BeginTrans
            With datExtra2
                .AddNew
                !staff_code = vStaff_Code
                !co_code = ID
                
                'Now get the basic pay
                datExtra.Find "Wage_code ='BSAL'", 0, adSearchForward, 1
                !Basic_Pay = CCur(datExtra!amount)
                
                If frmPSummary.prg.Value + prgv < 100 Then
                   frmPSummary.prg.Value = frmPSummary.prg.Value + prgv
                Else
                    frmPSummary.prg.Value = 100
                End If
                
                
                'Now get the summation of the other payments
                vTotal = 0
                datExtra.Requery
                While Not datExtra.EOF
                    If Trim(datExtra!wage_code) <> "HOUSE" And Trim(datExtra!wage_code) <> "BSAL" And Trim(datExtra!wage_code) <> "PMODE" And Trim(datExtra!wage_code) <> "GROSS" And Trim(datExtra!wage_code) <> "TAXPA" And Trim(datExtra!wage_code) <> "TAXPA" And Trim(datExtra!wage_code) <> "NET" And Trim(datExtra!wage_code) <> "NETAX" Then
                       If datExtra!Wage_operation = 1 Then
                            vTotal = vTotal + datExtra!amount
                       End If
                    End If
                    datExtra.MoveNext
                Wend
                !other_pmts = CCur(vTotal)
                
                If frmPSummary.prg.Value + prgv < 100 Then
                   frmPSummary.prg.Value = frmPSummary.prg.Value + prgv
                Else
                    frmPSummary.prg.Value = 100
                End If
                
                'Now get the summation of the benefits
                vTotal = 0
                datExtra.Requery
                While Not datExtra.EOF
                    If datExtra!Wage_operation = 3 Then
                         vTotal = vTotal + datExtra!amount
                    End If
                    datExtra.MoveNext
                Wend
                !benefits = CCur(vTotal)
                
                
                If frmPSummary.prg.Value + prgv < 100 Then
                   frmPSummary.prg.Value = frmPSummary.prg.Value + prgv
                Else
                    frmPSummary.prg.Value = 100
                End If
                
                'Get the housing if any
                datExtra.Filter = "Wage_Code ='HOUSE'"
                If datExtra.RecordCount <> 0 Then
                   !Housing = CCur(datExtra!amount)
                Else
                    !Housing = CCur(0)
                End If
                datExtra.Filter = ""
                
                
                If frmPSummary.prg.Value + prgv < 100 Then
                   frmPSummary.prg.Value = frmPSummary.prg.Value + prgv
                Else
                    frmPSummary.prg.Value = 100
                End If
                
                'Get the PAYE Charged
                datExtra.Find "Wage_code ='PAYE'", 0, adSearchForward, 1
                !PAYE = CCur(datExtra!amount)
                
                'Get the LSC
                datExtra.Filter = "Wage_Code ='LSC'"
                If datExtra.RecordCount <> 0 Then
                   !s_charge = CCur(datExtra!amount)
                Else
                    !s_charge = CCur(0)
                End If
                datExtra.Filter = ""
                
                If frmPSummary.prg.Value + prgv < 100 Then
                   frmPSummary.prg.Value = frmPSummary.prg.Value + prgv
                Else
                    frmPSummary.prg.Value = 100
                End If
                
                
                'Get the NSSF
                datExtra.Filter = "Wage_Code ='NSSF'"
                If datExtra.RecordCount <> 0 Then
                   !NSSF = CCur(datExtra!amount)
                Else
                    !NSSF = CCur(0)
                End If
                datExtra.Filter = ""
                
                
                If frmPSummary.prg.Value + prgv < 100 Then
                   frmPSummary.prg.Value = frmPSummary.prg.Value + prgv
                Else
                    frmPSummary.prg.Value = 100
                End If
                
                'Get the NHIF
                datExtra.Filter = "Wage_Code ='NHIF'"
                If datExtra.RecordCount <> 0 Then
                   !NHIF = CCur(datExtra!amount)
                Else
                    !NHIF = CCur(0)
                End If
                datExtra.Filter = ""
                
                
                If frmPSummary.prg.Value + prgv < 100 Then
                   frmPSummary.prg.Value = frmPSummary.prg.Value + prgv
                Else
                    frmPSummary.prg.Value = 100
                End If
                
                'Get the NET
                datExtra.Find "Wage_Code ='NET'", 0, adSearchForward, 1
                'If datExtra.RecordCount <> 0 Then
                   !NET = CCur(datExtra!amount)
                'Else
                '    !NET = CCur(0)
                'End If
                
                If frmPSummary.prg.Value + prgv < 100 Then
                   frmPSummary.prg.Value = frmPSummary.prg.Value + prgv
                Else
                    frmPSummary.prg.Value = 100
                End If
                
                
                'Now get the summation of the other deductions
                vTotal = 0
                datExtra.Requery
                While Not datExtra.EOF
                    If Trim(datExtra!wage_code) <> "LSC" And Trim(datExtra!wage_code) <> "NHIF" And Trim(datExtra!wage_code) <> "NSSF" And Trim(datExtra!wage_code) <> "PAYE" And Trim(datExtra!wage_code) <> "PRELI" And Trim(datExtra!wage_code) <> "PTAX" And Trim(datExtra!wage_code) <> "TPNDS" And Trim(datExtra!wage_code) <> "TAX" Then
                       If datExtra!Wage_operation = 2 Or datExtra!Wage_operation = 4 And datExtra!slip_display <> True Then
                            vTotal = vTotal + datExtra!amount
                       End If
                    End If
                    datExtra.MoveNext
                Wend
                
                
                If frmPSummary.prg.Value + prgv < 100 Then
                   frmPSummary.prg.Value = frmPSummary.prg.Value + prgv
                Else
                    frmPSummary.prg.Value = 100
                End If
                
                !other_ded = CCur(vTotal)
                !Month = vMonth
                !Year = vYear
                .Update
            End With
        cn.CommitTrans
    End If
End Sub

Public Function FillData(vdata As String, vFill As String, vPosn As Integer, vLen As Integer)
    'This function is written to make the filling either to the right or left of some data
    'and return a filled in string
    'Author: Patrick Odawo
    '03/Dec/2000
    
    'Data Received:  The data to be filled
    '                The fill position
    '                The filling character
    '                The required length of the return value
    
    Dim exLen As Integer
    Dim exFill As String
    Dim pLen
    'exFill = ""
    If vdata <> "" Then
       'Calculate the required fill length
       exLen = vLen - Len(vdata)
       For pLen = 1 To exLen
           exFill = exFill & vFill
       Next
       'MsgBox (Len(exFill))
       'Determine the fill position
       If vPosn = 1 Then   'This is a left fill
          vdata = exFill & vdata
       ElseIf vPosn = 2 Then 'This is a rigth fill
           vdata = vdata & exFill
       End If
    End If
    FillData = vdata
    'MsgBox (Len(vdata))
End Function

Public Sub ContinueProgram()
    CompID.Show vbModal
End Sub

Public Function CalculateDays(Start_Date As Date, End_Date As Date)
           'Get the number of non-working days between the two dates
           'Loop through each date there checking whether it is a saturday
           'or a sunday or set as a non working date
           Screen.MousePointer = vbHourglass
           Dim sDate As Date
           Dim eDate As Date
           Dim dtCount As Date
           Dim nwCount As Long
           Dim Saturday As Boolean
           Dim Sunday As Boolean
           sDate = CVDate(Start_Date)
           eDate = CVDate(End_Date)
           
           'Check if saturdays and sundays are non working days
           Set datExtra = New Recordset
           datExtra.Open "SELECT * FROM clnd_options WHERE co_code ='" & ID & "'", cn, adOpenStatic, adLockOptimistic
           If datExtra.RecordCount <> 0 Then
              If datExtra!saturdays = True Then
                 Saturday = True
              Else
                 Saturday = False
              End If
              If datExtra!sundays = True Then
                 Sunday = True
              Else
                 Sunday = False
              End If
           Else
              Sunday = False
              Saturday = False
           End If
           
           For dtCount = sDate To eDate Step 1
               'Get the non working days and check if this date is among them
               Set datExtra = New Recordset
               datExtra.Open "SELECT * FROM non_working WHERE co_code ='" & ID & _
               "' AND year =" & Year(dtCount) & " AND month =" & Month(dtCount) & _
               " AND date =" & Day(dtCount), cn, adOpenStatic, adLockOptimistic
               If datExtra.RecordCount <> 0 Then
                  'This day has been identified as a non working day
                  nwCount = nwCount + 1
               Else
                  'This day is not a recorded non working day, so check whether it is
                  'a saturday or a sunday
                    If Weekday(dtCount, vbSunday) = 1 Then
                       If Sunday = True Then
                          nwCount = nwCount + 1
                       End If
                    ElseIf Weekday(dtCount, vbSunday) = 7 Then
                       If Saturday = True Then
                          nwCount = nwCount + 1
                       End If
                    End If
               End If
           Next
           CalculateDays = DateDiff("d", CVDate(Start_Date) - 1, CVDate(End_Date)) - nwCount
           'txtleave(4) = DateDiff("d", CVDate(txtleave(2)) - 1, CVDate(txtleave(3))) - nwCount
           'MsgBox (nwCount)
           'txtleave(7) = txtleave(4)
           Screen.MousePointer = vbDefault
End Function
