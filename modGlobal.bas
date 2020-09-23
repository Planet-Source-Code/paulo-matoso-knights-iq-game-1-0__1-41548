Attribute VB_Name = "modGlobal"
'**********************************************************************************
'**Knights IQ Game 1.0
'**Copyright by Paulo Matoso
'**E-Mail: paulomt1@clix.pt
'**
'**                                 Global definitions
'**Last Modification ---> 11/12/2002
'**********************************************************************************
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Global Const MF_BYPOSITION = &H400&

Global Const Copyright = "Paulo Matoso"
Global Const EMail = "paulomt1@clix.pt"
Global Const Version = "1.0"

Global SquareOcupied(63) As Integer
Global ArraySolution(63) As Integer
Global GlobalCol(1 To 8) As Integer
Global GlobalRow(1 To 8) As Integer
Global KnightPlaces(1 To 8) As Integer
Global nKnightsInSquares As Integer 'Total of knights in squares
Global ShowHints As Integer 'Show or not Hints
Global LastSquare  As Integer ' Previous square, used for memove RedDot

'This disable the button X(Close button in main form)
Public Sub DisableFormButton(tForm As Form)
Dim lMenu As Long

    lMenu = GetSystemMenu(tForm.hWnd, False)
    DeleteMenu lMenu, 6, MF_BYPOSITION
End Sub

