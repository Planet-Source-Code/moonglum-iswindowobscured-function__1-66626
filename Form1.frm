VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   600
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum eObscureType
    eNone
    ePartial
    eTotal
End Enum

Private Const GW_HWNDPREV = 3


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Function IsWindowObscured(ByVal hWnd As Long) As eObscureType
    Dim hPrevWnd As Long                                            'Handle to previous window in z-order
    Dim hNextWnd As Long                                            'Handle to next window in z-order
    Dim tOther As RECT                                              'Rectangle of other window in z-order
    Dim tDest As RECT                                               'Rectangle of the intersection
    Dim tHwnd As RECT                                               'Rectangle of the passed window
    
    Call GetWindowRect(hWnd, tHwnd)                                 'Get the rectangle of the passed window
    hPrevWnd = hWnd                                                 'Hand off the passed window handle just to be consistentsy
    hNextWnd = GetWindow(hPrevWnd, GW_HWNDPREV)                     'Get next window in z-order
    Do                                                              'Loop until we have gone through all windows above target window
        Call GetWindowRect(hNextWnd, tOther)
        
        If IsWindowVisible(hNextWnd) Then                           'Is the window in question visible ?
            If IntersectRect(tDest, tHwnd, tOther) Then             'Does it intersect the target window at all ?
                If tHwnd.Top = tDest.Top And _
                   tHwnd.Bottom = tDest.Bottom And _
                   tHwnd.Right = tDest.Right And _
                   tHwnd.Left = tDest.Left Then                     'Totally obscured
                    
                    IsWindowObscured = eTotal
                    Exit Do '>>>>>                                  'Exit out, we are done
                Else    'If tHwnd.Top = tDest.Top And _...
                    IsWindowObscured = ePartial                     'Partially obscured, keep going up the z-order
                End If  'If tHwnd.Top = tDest.Top And _...
            End If  'If IntersectRect(tDest, tHwnd, tOther) Then
        End If  'If IsWindowVisible(hNextWnd) Then
        
        hPrevWnd = hNextWnd
        hNextWnd = GetWindow(hPrevWnd, GW_HWNDPREV)                 'Get next window in z-order
    Loop Until (hNextWnd = 0)                                       '0 = desktop & done
End Function
Private Sub Timer1_Timer()
    Select Case IsWindowObscured(Me.hWnd)
        Case eObscureType.eNone
            Debug.Print Now, "Not obscured"
            
        Case eObscureType.ePartial
            Debug.Print Now, "Partially obscured"
            
        Case eObscureType.eTotal
            Debug.Print Now, "Totally obscured"
    End Select
End Sub
