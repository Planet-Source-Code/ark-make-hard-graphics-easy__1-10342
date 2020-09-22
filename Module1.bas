Attribute VB_Name = "Module1"
Private Declare Function AlphaBlending Lib "msimg32.dll" Alias "AlphaBlend" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BF As Long) As Long
Private Declare Function DrawTransparent Lib "msimg32.dll" Alias "TransparentBlt" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

Public Function AlphaBlend(ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long
  Dim lngBlend As Long
  lngBlend = Val("&h" & Hex(AlphaSource) & "00" & "00")
  AlphaBlending destHDC, XDest, YDest, destWidth, destHeight, srcHDC, XSrc, ycrc, srcWidth, srcHeight, lngBlend
End Function

Public Function TransparentBlt(ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal TransparentColor As Long) As Long
  DrawTransparent destHDC, XDest, YDest, destWidth, destHeight, srcHDC, XSrc, ycrc, srcWidth, srcHeight, TransparentColor
End Function

