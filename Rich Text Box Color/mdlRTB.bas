Attribute VB_Name = "mdlRTB"
'This was done for your advantage, there are several code's
'changed especially fit for the color codes.
Global Const vbBrown = &H66&
Global Const vbDBlue = &H660000
Global Const vbDGreen = &H9900&
Global Const vbPurple = &H990099
Global Const vbOrange = &H66FF&
Global Const vbBGreen = &H999900
Global Const vbAqua = &HFFFF00
Global Const vbLightBlue = &HFF0000
Global Const vbPink = &HFF00FF
Global Const vbDGrey = &H666666
Global Const vbLGrey = &HCCCCCC
'Binding all the colors as the mIRC colour table
Public Const Col0 = vbBlack
Public Const Col1 = vbWhite
Public Const Col2 = vbDBlue
Public Const Col3 = vbDGreen
Public Const Col4 = vbRed
Public Const Col5 = vbBrown
Public Const Col6 = vbPurple
Public Const Col7 = vbOrange
Public Const Col8 = vbYellow
Public Const Col9 = vbGreen
Public Const Col10 = vbBGreen
Public Const Col11 = vbAqua
Public Const Col12 = vbLightBlue
Public Const Col13 = vbPink
Public Const Col14 = vbDGrey
Public Const Col15 = vbLGrey
'Parser
'The syntaxt is: Call Parse(Form Name, RichTextBox Name, String as an actual String)
'This may not be 100% perfect but it's the closest I can get it!
Function Parse(txtFrm As Form, txtView As RichTextBox, txtEcho As String)
    Dim Color As Boolean
    Color = False
    txtFrm.txtView.SelColor = vbBlack
    txtFrm.txtView.SelUnderline = False
    txtFrm.txtView.SelBold = False
    For i = 1 To Len(txtEcho)
        If Mid(txtEcho, i, 1) = Chr(15) Then
            txtFrm.txtView.SelBold = False
            txtFrm.txtView.SelUnderline = False
            txtFrm.txtView.SelColor = vbBlack
            GoTo Done2
        End If
        If Mid(txtEcho, i, 1) = Chr(2) And txtFrm.txtView.SelBold = False Then txtFrm.txtView.SelBold = True: GoTo Done
        If Mid(txtEcho, i, 1) = Chr(2) And txtFrm.txtView.SelBold = True Then txtFrm.txtView.SelBold = False: GoTo Done
        If Mid(txtEcho, i, 1) = Chr(31) And txtFrm.txtView.SelUnderline = False Then txtFrm.txtView.SelUnderline = True: GoTo Done
        If Mid(txtEcho, i, 1) = Chr(31) And txtFrm.txtView.SelUnderline = True Then txtFrm.txtView.SelUnderline = False: GoTo Done
        If Mid(txtEcho, i, 3) = Chr(3) + "15" Then Color = True: txtFrm.txtView.SelColor = Col15: i = (i + 2): GoTo Done
        If Mid(txtEcho, i, 3) = Chr(3) + "14" Then Color = True: txtFrm.txtView.SelColor = Col14: i = (i + 2): GoTo Done
        If Mid(txtEcho, i, 3) = Chr(3) + "13" Then Color = True: txtFrm.txtView.SelColor = Col13: i = (i + 2): GoTo Done
        If Mid(txtEcho, i, 3) = Chr(3) + "12" Then Color = True: txtFrm.txtView.SelColor = Col12: i = (i + 2): GoTo Done
        If Mid(txtEcho, i, 3) = Chr(3) + "11" Then Color = True: txtFrm.txtView.SelColor = Col11: i = (i + 2): GoTo Done
        If Mid(txtEcho, i, 3) = Chr(3) + "10" Then Color = True: txtFrm.txtView.SelColor = Col10: i = (i + 2): GoTo Done
        If Mid(txtEcho, i, 2) = Chr(3) + "9" Then Color = True: txtFrm.txtView.SelColor = Col9: i = (i + 1): GoTo Done
        If Mid(txtEcho, i, 2) = Chr(3) + "8" Then Color = True: txtFrm.txtView.SelColor = Col8: i = (i + 1): GoTo Done
        If Mid(txtEcho, i, 2) = Chr(3) + "7" Then Color = True: txtFrm.txtView.SelColor = Col7: i = (i + 1): GoTo Done
        If Mid(txtEcho, i, 2) = Chr(3) + "6" Then Color = True: txtFrm.txtView.SelColor = Col6: i = (i + 1): GoTo Done
        If Mid(txtEcho, i, 2) = Chr(3) + "5" Then Color = True: txtFrm.txtView.SelColor = Col5: i = (i + 1): GoTo Done
        If Mid(txtEcho, i, 2) = Chr(3) + "4" Then Color = True: txtFrm.txtView.SelColor = Col4: i = (i + 1): GoTo Done
        If Mid(txtEcho, i, 2) = Chr(3) + "3" Then Color = True: txtFrm.txtView.SelColor = Col3: i = (i + 1): GoTo Done
        If Mid(txtEcho, i, 2) = Chr(3) + "2" Then Color = True: txtFrm.txtView.SelColor = Col2: i = (i + 1): GoTo Done
        If Mid(txtEcho, i, 2) = Chr(3) + "1" Then Color = True: txtFrm.txtView.SelColor = Col1: i = (i + 1): GoTo Done
        If Mid(txtEcho, i, 2) = Chr(3) + "0" Then Color = True: txtFrm.txtView.SelColor = Col0: i = (i + 1): GoTo Done
    txtFrm.txtView.SelText = Mid(txtEcho, i, 1)
Done:
    If Mid(txtEcho, (i + 1), 1) = "," Then
       If Mid(txtEcho, (i + 2), 2) = "15" Then i = (i + 3)
       If Mid(txtEcho, (i + 2), 2) = "14" Then i = (i + 3)
       If Mid(txtEcho, (i + 2), 2) = "13" Then i = (i + 3)
       If Mid(txtEcho, (i + 2), 2) = "12" Then i = (i + 3)
       If Mid(txtEcho, (i + 2), 2) = "11" Then i = (i + 3)
       If Mid(txtEcho, (i + 2), 2) = "10" Then i = (i + 3)
       If Mid(txtEcho, (i + 2), 1) = "9" Then i = (i + 2)
       If Mid(txtEcho, (i + 2), 1) = "8" Then i = (i + 2)
       If Mid(txtEcho, (i + 2), 1) = "7" Then i = (i + 2)
       If Mid(txtEcho, (i + 2), 1) = "6" Then i = (i + 2)
       If Mid(txtEcho, (i + 2), 1) = "5" Then i = (i + 2)
       If Mid(txtEcho, (i + 2), 1) = "4" Then i = (i + 2)
       If Mid(txtEcho, (i + 2), 1) = "3" Then i = (i + 2)
       If Mid(txtEcho, (i + 2), 1) = "2" Then i = (i + 2)
       If Mid(txtEcho, (i + 2), 1) = "1" Then i = (i + 2)
       If Mid(txtEcho, (i + 2), 1) = "0" Then i = (i + 2)
    End If
Done2:
    Next i
End Function
