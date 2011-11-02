Attribute VB_Name = "modWin32Consts"
Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001

Public Const FILE_BEGIN = 0
Public Const FILE_END = 2

Public Const KEY_READ = &H20019

Public Const CF_SCREENFONTS As Long = &H1&
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_INITTOLOGFONTSTRUCT As Long = &H40&
Public Const CF_LIMITSIZE As Long = &H2000&

Public Const SB_VERT = 1

Public Const S_FALSE = 1
Public Const S_OK = 0

Public Const INADDR_NONE = &HFFFFFFFF

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_SHARE_DELETE = &H4
Public Const OPEN_ALWAYS = 4
Public Const FILE_ATTRIBUTE_NORMAL = &H80

Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100 'shorts
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_3DHILIGHT = 20
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_3DSHADOW = 16
Public Const COLOR_3DDKSHADOW = 21

Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Public Const GWL_WNDPROC = (-4)
Public Const GWL_STYLE = (-16)
Public Const WH_CALLWNDPROC = 4
Public Const WM_CREATE = &H1

Public Const WM_SETFOCUS = &H7

Public Const WM_TIMER = &H113

Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_SETFONT = &H30

Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORSCROLLBAR = &H137

Public Const LBS_OWNERDRAWFIXED = &H10&
Public Const LBS_HASSTRINGS = &H40&
Public Const LBS_NOINTEGRALHEIGHT = &H100&
Public Const LBS_EXTENDEDSEL = &H800&

'Listbox messages
Public Const LB_ERR = (-1)
Public Const LB_ADDSTRING = &H180
Public Const LB_INSERTSTRING = &H181
Public Const LB_DELETESTRING = &H182
Public Const LB_RESETCONTENT = &H184
Public Const LB_GETTEXT = &H189
Public Const LB_INITSTORAGE = &H1A8
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTOPINDEX = &H18E
Public Const LB_SETTOPINDEX = &H197
Public Const LB_GETSEL = &H187
Public Const LB_SETSEL = &H185
Public Const LB_GETITEMHEIGHT = &H1A1
Public Const LB_SETITEMHEIGHT = &H1A0
Public Const LB_GETSELCOUNT = &H190
Public Const LB_GETSELITEMS = &H191
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETCURSEL = &H186

Public Const ODS_SELECTED = &H1
Public Const ODS_FOCUS = &H10

Public Const ODA_FOCUS = &H4

Public Const RDW_INVALIDATE = &H1
Public Const RDW_ERASE = &H4
Public Const RDW_UPDATENOW = &H100
Public Const RDW_ERASENOW = &H200

Public Const LF_FACESIZE = 32

Public Const ESB_DISABLE_BOTH = &H3
Public Const ESB_ENABLE_BOTH = &H0

Public Const SPI_GETNONCLIENTMETRICS = 41

Public Const DRAFT_QUALITY = 1
Public Const NONANTIALIASED_QUALITY = 3

Public Const SYSTEM_FONT = 13

Public Const HS_DIAGCROSS = 5

Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

Public Const OBJ_PEN = 1
Public Const OBJ_BRUSH = 2
Public Const OBJ_FONT = 6

Public Const DLGC_WANTARROWS = &H1

'Windows messages
Public Const WM_SIZE = &H5
Public Const WM_WINDOWPOSCHANGED = &H47

Public Const WM_KILLFOCUS = &H8

Public Const WM_GETDLGCODE = &H87

Public Const WM_VSCROLL = &H115
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_MOUSELEAVE = &H2A3
Public Const WM_PASTE = &H302
Public Const WM_COMPAREITEM = &H39

Public Const WM_NCPAINT = &H85
Public Const WM_ERASEBKGND = &H14
Public Const WM_PAINT = &HF

Public Const WM_SYSKEYDOWN = &H104
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204

Public Const VK_CONTROL = &H11

Public Const TME_HOVER = &H1
Public Const TME_LEAVE = &H2
Public Const TME_CANCEL = &H80000000

Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_NOPREFIX = &H800
Public Const DT_END_ELLIPSIS = &H8000&

'ExtTextOut flags
Public Const ETO_CLIPPED = 4
Public Const ETO_OPAQUE = 2

'Window styles
Public Const WS_GROUP = &H20000
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_BORDER = &H800000
Public Const WS_CHILD = &H40000000

Public Const SBS_VERT = &H1&
Public Const SB_CTL = 2

'Scrollbar op codes
Public Const SB_BOTTOM = 7
Public Const SB_LINEDOWN = 1
Public Const SB_LINEUP = 0
Public Const SB_PAGEDOWN = 3
Public Const SB_PAGEUP = 2
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5
Public Const SB_TOP = 6
Public Const SB_ENDSCROLL = 8

Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10

Public Const SB_WIDTH = 16

Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5
Public Const SW_HIDE = 0

Public Const SM_CXBORDER = 5

Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_LINEINDEX = &HBB
Public Const WM_SETREDRAW = &HB
Public Const WM_USER = &H400

Public Const EM_GETSCROLLPOS = (WM_USER + 221)
Public Const EM_SETSCROLLPOS = (WM_USER + 222)

Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_CHARFROMPOS = &HD7
Public Const EM_GETLINECOUNT = &HBA

Public Const SCF_SELECTION As Long = &H1&
Public Const SCF_WORD As Long = &H2&
Public Const SCF_ALL As Long = &H4&

Public Const CFM_BACKCOLOR = &H4000000

Public Const EM_SETCHARFORMAT = (WM_USER + 68)

Public Const LR_LOADFROMFILE = 16
Public Const LR_SHARED = &H8000
Public Const LR_DEFAULTSIZE = &H40
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2

Public Const IDI_HAND = 32513&
Public Const IDC_HAND = 32649&

Public Const OCR_NORMAL = 32512
Public Const OCR_SIZEWE = 32644

Public Const IDC_SIZEWE = 32644
Public Const IDC_ARROW = 32512

Public Const PS_SOLID = 0
Public Const PS_DOT = 2
Public Const PS_NULL = 5

Public Const AC_SRC_OVER = &H0
