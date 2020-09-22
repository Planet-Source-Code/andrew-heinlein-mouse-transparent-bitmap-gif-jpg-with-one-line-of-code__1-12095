Attribute VB_Name = "TransparentBltMod"
'This was put together by Andrew Heinlein (Mouse)
'mouse@theblackhand.net
'I Found the below function in C++ while surfing on the web
'and changed it over to Visual Basic API cause i thought
'it would be nice to have around... Transparent Bitmaps... with simple API
'who woulda thought....

'This will return TRUE if it succeeds. If you are using Win2000/NT you
'can even use GetLastError to give a better error than just FALSE.
'99.9% of the time, if it returns false, its because a parameter is wrong

'BOOL TransparentBlt(
'  HDC hdcDest,        // handle to destination DC
'  int nXOriginDest,   // x-coord of destination upper-left corner
'  int nYOriginDest,   // y-coord of destination upper-left corner
'  int nWidthDest,     // width of destination rectangle
'  int hHeightDest,    // height of destination rectangle
'  HDC hdcSrc,         // handle to source DC
'  int nXOriginSrc,    // x-coord of source upper-left corner
'  int nYOriginSrc,    // y-coord of source upper-left corner
'  int nWidthSrc,      // width of source rectangle
'  int nHeightSrc,     // height of source rectangle
'  UINT crTransparent  // color to make transparent
');

Public Declare Function TransparentBlt Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Integer, ByVal nYOriginDest As Integer, ByVal nWidthDest As Integer, ByVal nHeightDest As Integer, ByVal hdcSrc As Long, ByVal nXOriginSrc As Integer, ByVal nYOriginSrc As Integer, ByVal nWidthSrc As Integer, ByVal nHeightSrc As Integer, ByVal crTransparent As Long) As Boolean
