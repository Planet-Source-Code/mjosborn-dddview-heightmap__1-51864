Attribute VB_Name = "modAPIUsed"
'The GetPixel function retrieves the red, green, blue (RGB) color value of the pixel at the specified coordinates.
'Params: hdc, Identifies the device context
'nXPos, Specifies the logical x-coordinate of the pixel to be examined.
'nYPos, Specifies the logical y-coordinate of the pixel to be examined.
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'The SetPixelV function sets the pixel at the specified coordinates to the closest approximation of the specified color.
'The point must be in the clipping region and the visible part of the device surface.
'Params: hdc Identifies the device context
'nXPos, Specifies the logical x-coordinate of the pixel to be examined.
'nYPos, Specifies the logical y-coordinate of the pixel to be examined.
'crColor, Specifies the color to be used to paint the point.
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'The SetRect function sets the coordinates of the specified rectangle. This is equivalent to assigning
'the left, top, right, and bottom arguments to the appropriate members of the RECT structure.
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'The PtInRect function determines whether the specified point lies within the specified rectangle.
'A point is within a rectangle if it lies on the left or top side or is within all four sides.
'A point on the right or bottom side is considered outside the rectangle.
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

'The ShellExecute function opens or prints a specified file. The file can be an executable file or a document file.
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long

'Types
'RECT structure that contains the specified rectangle.
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type COORD
    x As Long
    y As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type
'Other API's. I did'nt want to move them becouse I wanted to leave them where
'the original author placed them.

'The GetWindowsDirectory function retrieves the path of the Windows directory.
'The Windows directory contains such files as Windows-based applications, initialization files, and Help files.
'Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'The GetPrivateProfileString function retrieves a string from the specified section in an initialization file.
'This function is provided for compatibility with 16-bit Windows-based applications.
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'The WritePrivateProfileString function copies a string into the specified section of the specified initialization file.
'This function is provided for compatibility with 16-bit Windows-based applications.
'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'The QueryPerformanceCounter function retrieves the current value of the high-resolution performance counter, if one exists.
'Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
'The QueryPerformanceFrequency function retrieves the frequency of the high-resolution performance counter, if one exists.
'Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
