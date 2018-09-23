Attribute VB_Name = "mdRoundForm"
'声明API函数
Public Declare Function SetWindowRgn Lib "USER32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'函数CreateRoundRectRgn用于创建一个圆角矩形，该矩形由X1，Y1-X2，Y2确定，
'并由X3，Y3确定的椭圆描述圆角弧度
'CreateRoundRectRgn参数 类型及说明
'X1,Y1 Long，矩形左上角的X，Y坐标
'X2,Y2 Long，矩形右下角的X，Y坐标
'X3 Long，圆角椭圆的宽。其范围从0（没有圆角）到矩形宽（全圆）
'Y3 Long，圆角椭圆的高。其范围从0（没有圆角）到矩形高（全圆）
'SetWindowRgn用于将CreateRoundRectRgn创建的圆角区域赋给窗体
'DeleteObject用于将CreateRoundRectRgn创建的区域删除，这是必要的，否则不必要的占用电脑内存
'接下来声明一个全局变量,用来获得区域句柄，如下：
Dim outrgn As Long
'然后分别在窗体Activate()事件和Unload事件中输入以下代码
'Private Sub Form_Activate()
'Call rgnform(Me, 20, 20) '调用子过程
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'DeleteObject outrgn '将圆角区域使用的所有系统资源释放
'End Sub
'接下来我们开始编写子过程
Public Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long)
Dim w As Long, h As Long
w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels)
h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels)
outrgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
Call SetWindowRgn(frmbox.hWnd, outrgn, True)
End Sub

