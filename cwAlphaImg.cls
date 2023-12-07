VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwAlphaImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cwAlphaImg
' Author    : Olaf Schmidt
' Date      : 05/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit 'a very simple Image-Widget, which expects an ImageKey - and ensures ClickThrough-behaviour
 
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
      W.BackColor = -1  'we don't use any BackColor here
      W.HoverColor = vbRed 'show no Hover-Effect by default
      W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
      W.ImplementsWheelMessages = True
      W.Tag = 0.01 'we use this common cWidgetBase-Property here, to allow influencing the HoverColor-Alpha
End Sub

Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets: Set Widgets = W.Widgets: End Property
 
Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
  HitResultHit = False
  Dim Srf As cCairoSurface, Pxl() As Long
  If Cairo.ImageList.Exists(W.ImageKey) And W.HoverColor <> -1 Then Set Srf = Cairo.ImageList(W.ImageKey) Else Exit Sub
  If Not Srf.BindToArrayLong(Pxl) Or W.Width = 0 Or W.Height = 0 Then Exit Sub
     HitResultHit = Pxl(x * Srf.Width / W.Width, y * Srf.Height / W.Height) 'only when the Pixel==0==FullAlpha, will HitResultHit be returned False
  Srf.ReleaseArrayLong Pxl
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As RC6.cWidgetBase)
  W.Parent.Refresh 'if we want to support widget-refreshs "on-hover", we have to trigger a Re-Paint
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As RC6.cWidgetBase)
  W.Parent.Refresh 'same here (trigger dynamic re-rendering, when the hover-state changes)
End Sub

Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  Dim Srf As cCairoSurface ', x As Single, y As Single
  If Cairo.ImageList.Exists(W.ImageKey) Then Set Srf = Cairo.ImageList(W.ImageKey) Else Exit Sub

  CC.RenderSurfaceContent Srf, 0, 0, W.Width, W.Height, , W.Alpha 'render the current W.ImageKey-Surface (as loaded priorily into the ImageList)

  If W.MouseOver And W.HoverColor <> -1 Then 'render a colored, slightly blurred copy of the Srf with 25% Alpha (in case of being hovered)
     CC.RenderSurfaceContent Srf.GaussianBlur(0.1, , True, W.HoverColor), 0, 0, W.Width, W.Height, , W.Tag
  End If
End Sub

