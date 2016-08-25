VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StyleDocument 
   Caption         =   "Style Document"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   OleObjectBlob   =   "StyleDocument.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StyleDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal wClassName As Any, ByVal wWindowName As String) As Long
                                
                                ''''''''''''''''''''''''
Dim trackedChanges As Boolean   ' Save setting for user.
                                ''''''''''''''''''''''''
                            ''''''''''''''''''''''''
Dim hiddenText As Boolean   ' Save setting for user.
                            ''''''''''''''''''''''''
                                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim lastLetter As String, lastRoman As String   ' Global for assigning upon bookmark creation (BookmarkGW()). Will be used as test in DifferentiateSelection().
                                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim letterFound As String   ' Global for access by DoubleLetter function and DifferentiateSelection sub.
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ''''''''''''''''
Dim lstBx() As Variant  ' ListBox Array.
                        ''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''e''''''''''''''''''''''''''''''''''''''''
Dim bkmrks() As Variant ' This array will store the refs to created bookmarks so they can be browsed through (and later
                        ' set selections for deleting numbers/highlighted text). I'll ReDim with initial size of 50 entries.
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''
Dim bkmrkNum As Integer ' This will just count up to name new bookmarks.
                        ''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub MultiPage1_Change()

'''''''''
' Resize.
'''''''''
If Me.MultiPage1.Value = 0 Then
    Me.Height = 130
    Me.Width = 208

ElseIf Me.MultiPage1.Value = 1 Then
    Me.Height = 239
    Me.Width = 315

ElseIf Me.MultiPage1.Value = 2 Then
    Me.Height = 107
    Me.Width = 216
    ComboBox5.Clear

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This section populates styles in use dropdown on second tab (title marks):
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                            '''''''''''''''''''''''''''''''''''''''''''''''''
    For Each Stz In ActiveDocument.Styles   ' For every style in all styles in this document.
        For x = 1 To 9                      ' If style is list template style
            If Stz.ListLevelNumber = x Then ' (numbered hierarchy usually), then:
                                            '''''''''''''''''''''''''''''''''''''''''''''''''
                With ActiveDocument.Content.Find    '''''''''''''''''''''''''''''''''''''
                    .ClearFormatting                ' Use find to see if style is applied
                    .Text = ""                      ' (search for style in find/replace).
                    .Style = Stz                    '''''''''''''''''''''''''''''''''''''
                    .Execute Format:=True
                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If .Found = True Then       ' If style is found (and is list template style from before).
                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    With ComboBox5      ''''''''''''''''''''''''''''''''''''''
                        .AddItem Stz    ' Add to dropdown of 'heading styles'.
                    End With            ''''''''''''''''''''''''''''''''''''''
                
                End If
                End With
            End If  '''''''''''''''''''''''''''''''''''''''''''''''''''''''
         Next x     ' Test next list level 2-9 if necessary for each style.
    Next Stz        ' Test next style in all styles in document.
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''

ElseIf Me.MultiPage1.Value = 3 Then
    Me.Height = 94
    Me.Width = 208
End If

End Sub

Public Sub UserForm_Initialize()

On Error GoTo Error

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This section prepares doc and populates combobox on first tab, Style Document:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Me.MultiPage1.Value = 0

Selection.GoTo wdMainTextStory
ActiveDocument.ActiveWindow.View.ShowFieldCodes = True

If ActiveDocument.ActiveWindow.View.ShowHiddenText = False Then
    ActiveDocument.ActiveWindow.View.ShowHiddenText = True
    hiddenText = True
Else
    hiddenText = False
End If

If ActiveDocument.TrackRevisions = True Then
    ActiveDocument.TrackRevisions = False
    trackedChanges = True
Else
    trackedChanges = False
End If

CommandButton3.Visible = False  '''''''''''''''''''''''''''''''''''''
ComboBox3.Visible = False       ' Hide Suspect options unless needed.
Label4.Visible = False          '''''''''''''''''''''''''''''''''''''
'CommandButton4.Enabled = False  ' Disable options until needed.
'CommandButton2.Enabled = False  '''''''''''''''''''''''''''''''''''''
'CommandButton1.Enabled = False
'ComboBox2.Enabled = False
'ComboBox1.Enabled = False
CheckBox4.Enabled = False
MultiPage1.Pages("Page1").Enabled = False
MultiPage1.Pages("Page2").Enabled = False
MultiPage1.Pages("Page3").Enabled = False

'''''''''''''''''''''''''''''''''''''''''''''''
' Pull in selection options for style dropdown.
'''''''''''''''''''''''''''''''''''''''''''''''
For Each Stz In ActiveDocument.Styles   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For x = 1 To 9                      ' If style is list template style (numbered hierarchy usually).
        If Stz.ListLevelNumber = x Then '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StylesToAdd (Stz)
        End If  ''''''''''''''''''''''''''''''''''''''''''''
     Next x     ' Test next style in all styles in document.
Next Stz        ''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This section populates a dropdown on second tab (title marks):
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
With ComboBox4      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    .AddItem "."    ' Now add preselected title ending punctuations to second dropdown:
    .AddItem ","    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    .AddItem ":"
    .AddItem ";"    ''''''''''''''''''''''''''''''''''''''
End With            ' Done generating dropdown selections.
                    ''''''''''''''''''''''''''''''''''''''

Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     
If CloseMode = 0 Then
    DestroyProgram
End If

End Sub
 
Sub DestroyProgram()
                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If CommandButton1.Visible = True Then   ' This will check if numbers have been scanned for: if no, unload macro on X click.
        Selection.Find.ClearFormatting      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Unload Me
    ElseIf bkmrks(0)(0) = "Empty" Then
        Selection.Find.ClearFormatting
        Unload Me
    ElseIf bkmrkNum = 1 Then
        ReDim Preserve bkmrks(0 To 0)
        GoTo Continue
    ElseIf bkmrkNum = 2 Then
        ReDim Preserve bkmrks(0 To 1)
        GoTo Continue
    ElseIf bkmrkNum = 3 Then
        ReDim Preserve bkmrks(0 To 2)
        GoTo Continue
    Else
Continue:
        For Each bkmrk In bkmrks()                              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Selection.GoTo What:=wdGoToBookmark, Name:=bkmrk(1) ' This will remove highlighting, any bookmarks created and unload on X click.
                With Selection                                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .Range.HighlightColorIndex = wdNoHighlight
                End With
            ActiveDocument.bookmarks(bkmrk(1)).Delete
        Next bkmrk
        End If

ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
If hiddenText = True Then
    ActiveDocument.ActiveWindow.View.ShowHiddenText = False
End If
If trackedChanges = True Then
    ActiveDocument.TrackRevisions = True
End If

Selection.GoTo wdMainTextStory

Selection.Find.Text = ""
Selection.Find.ClearFormatting
Selection.Find.MatchWildcards = False
            ''''''''''''''
Unload Me   ' Goodbye, me!
            ''''''''''''''
End Sub
 
Public Sub CommandButton1_Click()

On Error GoTo Error
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim patterns()  ' Don't even bother looking at next few lines, it is a 2d array for find/replace pattern searches...
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
patterns = Array(Array("000", "^p^#^#.^#^#.^#^#.^#^#", "1.1.1.1"), Array("000", "^p^#.^#^#.^#^#.^#^#", "1.1.1.1"), Array("000", "^p^#.^#.^#.^#", "1.1.1.1"), _
                Array("001", "^p^#^#.^#^#.^#^#", "1.1.1"), Array("001", "^p^#^#.^#^#.^#", "1.1.1"), Array("001", "^p^#^#.^#.^#^#", "1.1.1"), Array("001", "^p^#.^#^#.^#^#", "1.1.1"), Array("001", "^p^#^#.^#.^#", "1.1.1"), Array("001", "^p^#.^#.^#^#", "1.1.1"), Array("001", "^p^#.^#^#.^#", "1.1.1"), Array("001", "^p^#.^#.^#", "1.1.1"), _
                Array("002", "^p^#^#.^#^#", "1.1"), Array("002", "^p^#^#.^#", "1.1"), Array("002", "^p^#.^#^#", "1.1"), Array("002", "^p^#.^#", "1.1"), _
                Array("004", "^pArticle ^#^#.^#^#", "Article 1.1"), Array("004", "^pArticle ^#.^#^#", "Article 1.1"), Array("004", "^pArticle ^#^#.^#", "Article 1.1"), Array("004", "^pArticle ^#.^#", "Article 1.1"), Array("005", "^pArticle ^#^#", "Article 1"), Array("005", "^pArticle ^#", "Article 1"), _
                Array("006", "^pArticle ^$^$^$^$^$", "Article I"), Array("006", "^pArticle ^$^$^$^$", "Article I"), Array("006", "^pArticle ^$^$^$", "Article I"), Array("006", "^pArticle ^$^$", "Article I"), Array("006", "^pArticle ^$", "Article I"), _
                Array("007", "^pSection ^#^#.^#^#", "Section 1.1"), Array("007", "^pSection ^#.^#^#", "Section 1.1"), Array("007", "^pSection ^#^#.^#", "Section 1.1"), Array("007", "^pSection ^#.^#", "Section 1.1"), Array("008", "^pSection ^#^#", "Section 1"), Array("008", "^pSection ^#", "Section 1"), Array("009", "^pSection ^$^$^$", "Section I"), Array("009", "^pSection ^$^$", "Section I"), Array("009", "^pSection ^$", "Section I"), _
                Array("010", "^p(00^#)", "(001)"), Array("010", "^p(0^#^#)", "(001)"), Array("011", "^p(^#^#^#)", "(1)"), Array("011", "^p(^#^#)", "(1)"), Array("011", "^p(^#)", "(1)"), _
                Array("012", "^p^#^#^#)", "1)"), Array("021", "^p^#^#)", "1)"), Array("012", "^p^#)", "1)"), _
                Array("013", "^p^#^#^#^#.^t", "1."), Array("013", "^p^#^#^#.^t", "1."), Array("013", "^p^#^#.^t", "1."), Array("013", "^p^#.^t", "1."), Array("013", "^p^#^#^#^#. ", "1."), Array("013", "^p^#^#^#. ", "1."), Array("013", "^p^#^#. ", "1."), Array("013", "^p^#. ", "1."), _
                Array("014", "^p^#^#^#^#^t", "1"), Array("014", "^p^#^#^#^t", "1"), Array("013", "^p^#^#^t", "1"), Array("014", "^p^#^t", "1"), _
                Array("039", "^p•^t", "•"), Array("040", "^p›^t", "›"), Array("041", "^p-^t", "-"))         ' Array("014", "^p^#^#^#^# ", "1"), Array("014", "^p^#^#^# ", "1"), Array("014", "^p^#^# ", "1"), Array("014", "^p^# ", "1"),
                                                                                                            ' The above are problematic and probably unnecessary anyway.
                    '''''''''''''''''''''''''''''''''''''''''                                               ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim patternsGW()    ' ...and a second for WildCard searches.
                    '''''''''''''''''''''''''''''''''''''''''
patternsGW = Array(Array("015", "016", "[^13][^040][A-Z]{1,9}[^041]", "(A)", "(I)", "Suspect"), Array("017", "018", "[^13][^040][a-z]{1,9}[^041]", "(a)", "(i)", "Suspect"), _
                Array("019", "020", "[^13][^040][A-Z]{1,9}[.][^041]", "(A.)", "(I.)", "Suspect"), Array("021", "022", "[^13][^040][a-z]{1,9}[.][^041]", "(a.)", "(i.)", "Suspect"), _
                Array("023", "024", "[^13][A-Z]{1,9}[^041]", "A)", "I)", "Suspect"), Array("025", "026", "[^13][a-z]{1,9}[^041]", "a)", "i)", "Suspect"), _
                Array("027", "028", "[^13][A-Z]{1,9}[.][^041]", "A.)", "I.)", "Suspect"), Array("029", "030", "[^13][a-z]{1,9}[.][^041]", "a.)", "i.)", "Suspect"), _
                Array("031", "032", "[^13][A-Z]{1,9}[.]", "A.", "I.", "Suspect"), Array("033", "034", "[^13][a-z]{1,9}[.]", "a.", "i.", "Suspect"), _
                Array("035", "036", "[^13][A-Z]{1,9}[^9]", "A", "I", "Suspect"), Array("037", "038", "[^13][a-z]{1,9}[^9]", "a", "i", "Suspect", "046", "o"))

                '''''''''''''''''''''
Dim bullets()   ' ..last for bullets.
                '''''''''''''''''''''
bullets = Array(Array("042", "-3880", "Arrow Bullet"), Array("043", "-3929", "Square Bullet"), Array("044", "-3937", "Small Bullet"), Array("045", "-3988", "Large Bullet"), Array("046", "-3913", "Symbol Bullet"), Array("047", "-3920", "Symbol O Bullet"))

ReDim bkmrks(0 To 200) As Variant
bkmrks(0) = Array("Empty", "", "", "")
bkmrks(1) = Array("Empty", "", "", "")
bkmrks(3) = Array("Empty", "", "", "")

                        ''''''''''''''''''''''''''''''''''''''''''''''
bkmrkNum = 0            ' Start bookmarks at 1 in array bkmrks().
Dim bkmrkName As String ' This will just be name of new bookmarks.
Dim ptrnFound As String ' This will just be the actually found number.
                        ''''''''''''''''''''''''''''''''''''''''''''''
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This part for Patterns searches (will contain all logic because it only handles numbered cases)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.HomeKey Unit:=wdStory 'Just set selection to top of doc for simplicity; later I'll change it to only work within a
                                'user defined selection so we can leave out Exhibits, etc.
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                            ''''''''''''''''''''''''''''''''''''
For Each ptrn In patterns() ' Each pattern in array patterns()
                            ''''''''''''''''''''''''''''''''''''
    With Selection.Find
        .ClearFormatting    '''''''''''''''''''
        .Text = ptrn(1)     ' Find the pattern
        .Forward = True     '''''''''''''''''''
        .Wrap = wdFindContinue  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .MatchWildcards = False ' Cause .find settings are sticky!
        .Highlight = False      ' So long as it hasn't been found already (highlight false)
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While Selection.Find.Execute    ' This portion will add the associated pattern to ComboBox2 so they can be selected to apply and browsed in the ListBox
        If .Found = True Then       ' (bookmarking the pattern found, minus leading carriage return, plus any ending pattern or periods, spaces, tabs).
        With ComboBox2              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim dupe                ' Make sure each option is only listed once:
            dupe = False            '''''''''''''''''''''''''''''''''''''''''''''
            For i = 0 To ComboBox2.ListCount - 1
                If ptrn(2) = ComboBox2.List(i) Then dupe = True
            Next i                                          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If dupe = False Then ComboBox2.AddItem ptrn(2)  ' Add to dropdown of 'patterns' found if isn't a duplicate.
        End With                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    '''''''''''''''''''''''''''''''''''''''''''''
        With ComboBox3              ' Same deal for ComboBox3 for Suspects.
            Dim dupl                ' Make sure each option is only listed once:
            dupl = False            '''''''''''''''''''''''''''''''''''''''''''''
            For k = 0 To ComboBox3.ListCount - 1
                If ptrn(2) = ComboBox3.List(k) Then dupl = True
            Next k                                               '''''''''''''''''''''''''''''''''''''''
                If dupl = False Then ComboBox3.AddItem ptrn(2)   ' Add to dropdown of 'Suspects' found.
        End With                                                 '''''''''''''''''''''''''''''''''''''''
                                                            
                                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        With Selection                                      ' This portion will set up bookmarks, highlight 'numbering' for review
            .MoveStart Unit:=wdCharacter, Count:=1          ' Deselect the paramarker at left of selection (which was found by pattern search from arrays above)
            .MoveEndWhile Cset:=".,:; " & Chr(9), Count:=4  ' Add ending periods, tabs, whatever to number selection (not bothering to differentiate except where needed)
            .Range.HighlightColorIndex = wdTurquoise        ' Highlight selection turquoise (to tell code it's been bookmarked, for user reference)
        ptrnFound = Selection                               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        bkmrkName = "skdn" & ptrn(0) & "x" & bkmrkNum                               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ActiveDocument.bookmarks.Add Name:=bkmrkName, Range:=Selection.Range    ' Bookmark selection for later review/deletion (with unique name)
                                                                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If bkmrkNum > UBound(bkmrks) - 1 Then ReDim Preserve bkmrks(0 To UBound(bkmrks) + 50)   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        bkmrks(bkmrkNum) = Array(ptrn(0), bkmrkName, ptrn(2), ptrnFound)                        ' Pattern code, bookmark name, bookmark pattern style, actual character.
                                                                                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        bkmrkNum = bkmrkNum + 1     ' Add 1 to bookmark name number so they don't duplicate.
            .Collapse wdCollapseEnd ' This collapses selection so that next array .find won't fail
        End With                    ' because selection was same size as searched for pattern (buggy .find)
                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If      '''''''''''''''''''''''''''''''''''''''''''''''''
    Wend            ' Keep going until all entries searched for/found.
        End With    '''''''''''''''''''''''''''''''''''''''''''''''''
Next ptrn
                                ''''''''''''''''''''''''''''''
Selection.HomeKey Unit:=wdStory ' Move selection back for now.
                                ''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This part for patternsGW searches (Contains logic for letter/Roman numeral checking,
' see Sub BookmarkGW() for full bookmarking, DoubleLetter() and DifferentiateSelect() also).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                '''''''''''''''''''''''''''''''''''''
For Each ptrnGW In patternsGW() ' Each pattern in array patternsGW().
    With Selection.Find         '''''''''''''''''''''''''''''''''''''
        .ClearFormatting    '''''''''''''''''''
        .Text = ptrnGW(2)   ' Find the pattern.
        .Forward = True     '''''''''''''''''''
        .Wrap = wdFindContinue  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .Highlight = False      ' So long as it hasn't been found already (highlights when found).
        .MatchWildcards = True  ' Added for second round of searches that are Wildcard.
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While Selection.Find.Execute            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .Found = True Then               ' This section differentiates Roman numerals and letters, lots of repetition.
            DifferentiateSelection (ptrnGW) '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        End If
    Wend        ''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Keep going until all entries searched for/found
    End With    ''''''''''''''''''''''''''''''''''''''''''''''''''
    
Selection.HomeKey Unit:=wdStory
Next ptrnGW
                
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This part for all bullet searches remaining. All logic contained.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Selection.Collapse wdCollapseEnd

For Each blt In bullets()

    With Selection.Find
        .ClearFormatting
        .Text = "^p" & ChrW(blt(1))
        .Forward = True
        .Wrap = wdFindContinue
        .Highlight = False
        .MatchWildcards = False
    
        While Selection.Find.Execute
            If .Found = True Then
                With ComboBox2         ''''''''''''''''''''''''''''''''''''''''''''
                    Dim duped          ' Make sure each option is only listed once:
                    duped = False      ''''''''''''''''''''''''''''''''''''''''''''
                    For i = 0 To ComboBox2.ListCount - 1
                        If ChrW(blt(1)) = ComboBox2.List(i) Then duped = True
                    Next i                                                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If duped = False Then ComboBox2.AddItem ChrW(blt(1))    ' Add to dropdown of 'patterns' found if isn't a duplicate.
                End With                                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                With Selection                                                  ' This portion will set up bookmarks, highlight 'numbering' for review
                    .MoveStart Unit:=wdCharacter, Count:=1                      ' Deselect the paramarker at left of selection (which was found by pattern search from arrays above)
                    .MoveEndWhile Cset:=".,:; " & Chr(9) & Chr$(160), Count:=4  ' Add ending periods, tabs, whatever to number selection (not bothering to differentiate except where needed)
                    .Range.HighlightColorIndex = wdTurquoise                    ' Highlight selection turquoise (to tell code it's been bookmarked, for user reference)
                ptrnFound = Selection                                           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                bkmrkName = "skdn" & blt(0) & "x" & bkmrkNum                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ActiveDocument.bookmarks.Add Name:=bkmrkName, Range:=Selection.Range    ' Bookmark selection for later review/deletion (with unique name)
                                                                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                If bkmrkNum > UBound(bkmrks) - 1 Then ReDim Preserve bkmrks(0 To UBound(bkmrks) + 50)   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                bkmrks(bkmrkNum) = Array(blt(0), bkmrkName, ChrW(blt(1)), ptrnFound)                    ' Pattern code, bookmark name, bookmark pattern style, actual character.
                                                                                                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                            
                bkmrkNum = bkmrkNum + 1
                    .Collapse wdCollapseEnd
                End With
            End If      '''''''''''''''''''''''''''''''''''''''''''''''''
        Wend            ' Keep going until all entries searched for/found.
    End With            '''''''''''''''''''''''''''''''''''''''''''''''''
Next blt
                                        
                    ''''''''''''''''''''''''''''''''''''''''''''''''
BkmrksReassignSus   ' Use BkmrksReassignSus Sub to rework ComboBox2.
                    ''''''''''''''''''''''''''''''''''''''''''''''''
                                ''''''''''''''''''''''''''''''
Selection.HomeKey Unit:=wdStory ' Move selection back for now.
                                ''''''''''''''''''''''''''''''
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CommandButton1.Visible = False  ' Hide Scan button so no rescanning happens accidentally.
CommandButton4.Visible = True   ' Show Remove button after scan only so it doesn't error.
Label5.Visible = True           ' Also show remove label.
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub

Public Sub CommandButton2_Click()   ''''''''''''''''''''''''''''
                                    ' Apply styles click script.
                                    ''''''''''''''''''''''''''''
On Error GoTo Error
                                                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If ComboBox1 = "" Or ComboBox2 = "" Then                            ' Make sure they make selections from dropdowns (typing disabled).
    MsgBox ("Please make a selection from both dropdown menus.")    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    '''''''''''''''''''''''''''''''''''''''''''''''''''
Else                                ' Let them run it if they made dropdown selections:
    Selection.Find.ClearFormatting  '''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each bkmrk In bkmrks()                                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If bkmrk(2) = ComboBox2.Text Then                       ' Check if bookmark matches pattern selected, then apply.
            Selection.GoTo What:=wdGoToBookmark, Name:=bkmrk(1) '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim oldbkmrk As Range
                Set oldbkmrk = ActiveDocument.bookmarks(bkmrk(1)).Range ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                oldbkmrk.Text = ""                                      ' Capture old bookmark location, delete text and old bookmark, remake bookmark..
                ActiveDocument.bookmarks.Add bkmrk(1), oldbkmrk         ' ..can't remove bookmarked text without removing bookmark.
            With Selection                                              ' Style with 'heading style' picked in dropdown and move to next.
                .Style = ComboBox1.Text                                 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End With
        End If
    Next bkmrk
End If

Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub

Public Sub CommandButton3_Click()   '''''''''''''''''''''''''''''''
                                    ' Assign Suspects click script.
                                    '''''''''''''''''''''''''''''''
On Error GoTo Error

Dim foundSuspect As Boolean
foundSuspect = False
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim StupidArray()   ' This will let me get correct pattern code to update bkmrks()
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
StupidArray = Array(Array("015", "016", "(A)", "(I)"), Array("017", "018", "(a)", "(i)"), _
                Array("019", "020", "(A.)", "(I.)"), Array("021", "022", "(a.)", "(i.)"), _
                Array("023", "024", "A)", "I)"), Array("025", "026", "a)", "i)"), _
                Array("027", "028", "A.)", "I.)"), Array("029", "030", "a.)", "i.)"), _
                Array("031", "032", "A.", "I."), Array("033", "034", "a.", "i."), _
                Array("035", "036", "A", "I"), Array("037", "038", "a", "i"))

If ComboBox2.Text = "Suspect" Then
    If ListBox1.Text = "" Then
        MsgBox ("Please select a Suspect from the list.")
    Else
    For Each stpd In StupidArray
        If ComboBox3.Text = stpd(2) Then                                ''''''''''''''''''''''''''''''''''
            bkmrks(lstBx(ListBox1.ListIndex)(2))(2) = ComboBox3.Text    ' This reassigns bkmrks() entries.
            bkmrks(lstBx(ListBox1.ListIndex)(2))(0) = stpd(0)           ''''''''''''''''''''''''''''''''''
        ElseIf ComboBox3.Text = stpd(3) Then
            bkmrks(lstBx(ListBox1.ListIndex)(2))(2) = ComboBox3.Text
            bkmrks(lstBx(ListBox1.ListIndex)(2))(0) = stpd(1)
        End If
    Next stpd
    End If
End If

For Each bkmrk In bkmrks()          '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If bkmrk(0) = "Suspect" Then    ' Run through bkmrks(), see if any Suspects are left.
        ComboBox2_Click             '''''''''''''''''''''''''''''''''''''''''''''''''''''
        foundSuspect = True
        Exit For
    End If
Next bkmrk

If foundSuspect = False Then                    ''''''''''''''''''''''''''''''''''''''''''''''''
    ComboBox2.RemoveItem ComboBox2.ListIndex    ' If no Suspects found, disable Suspect options.
    ComboBox2 = ComboBox3.Text                  ''''''''''''''''''''''''''''''''''''''''''''''''
    ComboBox2_Click
    CommandButton3.Visible = False
    ComboBox3.Visible = False
    Label4.Visible = False
End If

Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub
                                    ''''''''''''''''''''''''''''''''''''''''
Private Sub CommandButton4_Click()  ' Delete Unwanted Patterns click script.
                                    ''''''''''''''''''''''''''''''''''''''''
On Error GoTo Error
                                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim dltArray() As Variant, dltCount As Integer  ' Setting up array to keep indices of items to delete.
dltCount = 0                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''

If ComboBox2 <> "" And ListBox1.Text <> "" Then
    ReDim dltArray(0 To ListBox1.ListCount - 1)
    
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then ''''''''''''''''''''''''''''''''''''''''''
            dltArray(dltCount) = i          ' This catches selected items in ListBox1,
            dltCount = dltCount + 1         ' and adds them to dltArray.
        End If                              ''''''''''''''''''''''''''''''''''''''''''
    Next i
    
    For Each q In dltArray()            '''''''''''''''''''''''''''''''''''''''''''''''''
        If IsEmpty(q) Then              ' This will remove empty entries from dltArray().
            nullCount = nullCount + 1   '''''''''''''''''''''''''''''''''''''''''''''''''
        End If
    Next q
    ReDim Preserve dltArray(LBound(dltArray) To UBound(dltArray) - nullCount)
    
    If dltCount > 0 And dltCount < ListBox1.ListCount Then
        For Each dlt In dltArray                    ''''''''''''''''''''''''''''''''''
            bkmrks(lstBx(dlt)(2))(2) = "Removed"    ' This reassigns bkmrks() entries.
            bkmrks(lstBx(dlt)(2))(0) = "Removed"    ''''''''''''''''''''''''''''''''''
        Next dlt
        ComboBox2_Click
    ElseIf dltCount = ListBox1.ListCount Then
        ComboBox2.RemoveItem ComboBox2.ListIndex
        ComboBox2.ListIndex = -1
    Else
        MsgBox ("Please select a listed Bookmark to remove it from the list.")
    End If
End If

Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub

Private Sub CommandButton5_Click()

On Error GoTo Error

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Box checked for converting autonumbering to text.
'''''''''''''''''''''''''''''''''''''''''''''''''''

If CheckBox1 = True Then
    For Each para In ActiveDocument.Paragraphs
        If para.IsStyleSeparator = True Then
            para.Range.Select
            With Selection
                .Font.Hidden = False
                '.Range.ListFormat.ConvertNumbersToText
                '.ClearCharacterDirectFormatting
                '.ClearParagraphDirectFormatting
                '.ClearParagraphStyle
                '.Collapse (wdCollapseEnd)
                '.TypeBackspace
            End With
        End If
        
    Next para
    
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Box checked for removing leading tabs and spaces.
'''''''''''''''''''''''''''''''''''''''''''''''''''

If CheckBox2 = True Then
    With Selection.Find     '''''''''''''''''''''''''''''''''''''
        .ClearFormatting    ' Replace tab-space-tab with tab-tab.
        .Text = "^t ^t"     '''''''''''''''''''''''''''''''''''''
        .Replacement.Text = "^t^t"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
    While .Execute
        If .Found = True Then
            .Execute Replace:=wdReplaceAll
        End If
    Wend
    End With
    
    With Selection.Find     '''''''''''''''''''''''''
        .ClearFormatting    ' Remove leading periods.
        .Text = "^p "       '''''''''''''''''''''''''
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
    While .Execute
        If .Found = True Then
            .Execute Replace:=wdReplaceAll
        End If
    Wend
    End With
    
    With Selection.Find         '''''''''''''''''''''''''''''''''''''''''''''''''''
            .ClearFormatting    ' Remove leading tabs, change to first line indent.
            .Text = "^p^t"      '''''''''''''''''''''''''''''''''''''''''''''''''''
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWildcards = False
        While Selection.Find.Execute
            If .Found = True Then
                Selection.MoveStart Unit:=wdCharacter, Count:=1
                Selection.MoveEndWhile Cset:=Chr(9), Count:=12
            End If
            With Selection.ParagraphFormat
                .FirstLineIndent = InchesToPoints(0.5 * Len(Selection))
                Selection.Delete
            End With
        Wend
    End With
End If

'''''''''''''''''''''''''''''''''''''''''
' Box checked for removing graphic lines.
'''''''''''''''''''''''''''''''''''''''''

If CheckBox3 = True Then                                '''''''''''''''''''''''''''''''''''''''''''''''''
    For i = ActiveDocument.shapes.Count To 1 Step -1    ' Count backwards as you delete, to not miss any.
        If ActiveDocument.shapes.Range(i).Type = msoLine Then
            ActiveDocument.shapes.Range(i).Delete
        End If
    Next
End If

'For Each ln In ActiveDocument.Shapes
'    For i = ActiveDocument.Shapes.Count To 1 Step -1
'        If ReturnSelection(ln.Name) = "Straig" Then
'            ln.Delete
'        End If
'    Next i
'Next ln


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Box checked for remove applied styles, retain formatting.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If CheckBox4 = True Then
    GoAheadAndRemoveThen
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Box checked to reflow document (close split paragraphs).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If CheckBox5 = True Then

Dim spltRange As Range
Dim spltStyle As String, spltString As String



MovingOn:
    With Selection.Find
        .ClearFormatting        '''''''''''''''''''
        .Text = "([!.;:])(^13)" ' Find this pattern.
        .Forward = True         '''''''''''''''''''
        .Wrap = wdFindAsk  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .MatchWildcards = True  ' Cause .find settings are sticky!
        .Highlight = False      ' So long as it hasn't been found already (highlight false).
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    ''''''''''''''''''''''''''''
    While .Execute                  ' Find possible split paras.
        If .Found = True Then       ''''''''''''''''''''''''''''
            spltString = Selection
            If spltString = "d" & vbCr Then
                Selection.MoveStart Unit:=wdCharacter, Count:=-3
                spltString = Selection
                If spltString = " and" & vbCr Then
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                    Selection.Collapse
                    GoTo MovingOn
                Else: GoTo OtherCase
                End If
            ElseIf spltString = "r" & vbCr Then
                Selection.MoveStart Unit:=wdCharacter, Count:=-2
                spltString = Selection
                If spltString = " or" & vbCr Then
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                    Selection.Collapse
                    GoTo MovingOn
                Else: GoTo OtherCase
                End If
            Else
OtherCase:
                With Selection
                    .MoveStartWhile Cset:=" .:;", Count:=-5
                    spltString = Selection
                    If Mid(spltString, 1, 1) = "." Or Mid(spltString, 1, 1) = ":" Or Mid(spltString, 1, 1) = ";" Then
                        Selection.MoveRight Unit:=wdCharacter, Count:=1
                        Selection.Collapse
                        GoTo MovingOn
                    Else
                        Selection.MoveStartUntil Cset:=vbCr, Count:=wdForward
                        Selection.MoveEndWhile Cset:=vbCr, Count:=wdForward
                        Selection.MoveEnd Unit:=wdCharacter, Count:=1
                        spltString = Selection
                        Select Case Asc(Mid(spltString, Len(spltString), 1))
                            Case 65 To 90
                                Selection.MoveEndUntil Cset:=" ", Count:=9
                                Selection.Range.HighlightColorIndex = wdPink
                                Selection.Collapse
                                GoTo MovingOn
                            Case 97 To 122
                                With Selection
                                    .MoveEnd Unit:=wdCharacter, Count:=-1
                                    .Delete Unit:=wdCharacter, Count:=1
                                    .InsertBefore " "
                                    .MoveStart Unit:=wdCharacter, Count:=-1
                                    .MoveStartUntil Cset:=" ", Count:=wdBackward
                                    .MoveEndUntil Cset:=" ", Count:=wdForward
                                    .Range.HighlightColorIndex = wdGreen
                                End With
                            Case Else
                                Selection.MoveRight Unit:=wdCharacter, Count:=1
                                Selection.Collapse
                                GoTo MovingOn
                        End Select
                    End If
                End With
            End If
        End If
    Wend
    End With
End If


'Label7.Visible = False          ''''''''''''''''''''''''''''''''
'Label6.Visible = False          ' Hide Options to free up space.
'Label8.Visible = False          ''''''''''''''''''''''''''''''''
'Label9.Visible = False
'Label10.Visible = False
'CheckBox1.Visible = False
'CheckBox2.Visible = False
'CheckBox3.Visible = False
'CheckBox4.Visible = False
'CommandButton5.Visible = False
'
'                               '''''''''''''''''''''''''''
'CommandButton4.Enabled = True  ' Enable Options as needed.
'CommandButton2.Enabled = True  '''''''''''''''''''''''''''
'CommandButton1.Enabled = True
'ComboBox2.Enabled = True
'ComboBox1.Enabled = True
MultiPage1.Pages("Page1").Enabled = True
MultiPage1.Pages("Page2").Enabled = True
MultiPage1.Pages("Page3").Enabled = True


Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub

Private Sub CommandButton6_Click()
                                            
On Error GoTo Error                         ''''''''''''''''''''''
                                            ' Button click script.
If ComboBox4 = "" Or ComboBox5 = "" Then    ''''''''''''''''''''''
    MsgBox ("Please select from the dropdown menus.")       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                            ' Make sure they make selections from dropdowns (stuff can be typed in as well).
Else                                                        ' Let them run it if they made dropdown selections:
    Selection.HomeKey Unit:=wdStory                         ' Move cursor to beginning of document.
    Selection.Find.ClearFormatting                          ' Find first title ending pattern in styled paragraph selected:
    Selection.Find.Style = ActiveDocument.Styles(ComboBox5) ' Find by style name line (this line could've been without.
                                                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With Selection.Find     ''''''''''''''''''''''''''''
        .Text = ComboBox4   ' Find title ending pattern.
        .Forward = True     ''''''''''''''''''''''''''''
        .Wrap = wdFindContinue  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End With                    ' Find title ending pattern from top to bottom, wrapping back around if end is reached.
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While Selection.Find.Execute                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Selection = ComboBox4 Then                                       ' If selected style applied and pattern found:
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove  ' Deselect title ending pattern, move insertion point left one space.
                Application.Run MacroName:="TOCMark"                        ' Insert title mark with already created firm macro. :)
                                                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Wend        ' While/Wend statement to find each pattern & style selected.
End If          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub

Private Sub CommandButton7_Click()
                    ''''''''''''''''''''''''''''''''
On Error Resume Next 'GoTo Error ' Footnotes button click script.
                    ''''''''''''''''''''''''''''''''
Dim ftntRange As Range
Dim ftntString As String
Dim srchArray()
    srchArray = Array("^#^#^#^#", "^#^#^#", "^#^#", "^#")

Selection.HomeKey Unit:=wdStory

For Each srch In srchArray

HopeThisWorks:

    With Selection.Find
        .ClearFormatting    '''''''''''''''''''
        .Text = srch        ' Find the pattern.
        .Forward = True     '''''''''''''''''''
        .Font.Superscript = True
        .Wrap = wdFindContinue  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .MatchWildcards = False ' Cause .find settings are sticky!
        .Highlight = False      ' So long as it hasn't been found already (highlight false).
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                        '''''''''''''''''''''''''''''''''''''''''''
        .Execute                        ' Find the superscripted footnote location.
            If .Found = True Then       '''''''''''''''''''''''''''''''''''''''''''
                Set ftntRange = Selection.Range
                ftntString = Selection
                Selection.Collapse
                
                With Selection.Find
                    .ClearFormatting                            '''''''''''''''''''''''''''
                    .Text = "(^13)(" & ftntString & ")([ ^t])"  ' Find the actual footnote.
                    .Forward = True                             '''''''''''''''''''''''''''
                    .Wrap = wdFindStop      ''''''''''''''''''''''''''''''''''
                    .MatchWildcards = True  ' Cause .find settings are sticky!
                    .Highlight = False      ''''''''''''''''''''''''''''''''''
                                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .Execute                        ' Find the footnote para and cut it, paste as footnote:
                        If .Found = True Then       '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            With Selection                                      '''''''''''''''''''''''''''''''''''''''''''''''
                                .MoveStart Unit:=wdCharacter, Count:=1          ' Deselect the paramarker at left of selection,
                                .MoveEndWhile Cset:=" " & Chr(9)                ' Take out leading spaces,
                                .Delete                                         ' Delete number,
                                .Collapse                                       ' Collapse selection,
                                .MoveEndUntil Cset:=vbCr, Count:=wdForward      ' Include full para in selection,
                                .Cut                                            ' Cut.
                            End With                                            ''''''''''''''''''''''''''''''''''''''''''''''''
                            ftntRange.Select
                                Selection.Delete Unit:=wdCharacter, Count:=1
                                With Selection
                                    'With .FootnoteOptions
                                    '    .Location = wdBottomOfPage
                                    '    .NumberingRule = wdRestartContinuous
                                    '    .StartingNumber = 1
                                    '    .NumberStyle = wdNoteNumberStyleArabic
                                    'End With
                                    .Footnotes.Add Range:=Selection.Range, Reference:=""
                                End With
                                Selection.PasteAndFormat (wdFormatOriginalFormatting)
                            ftntRange.Select
                        Else
                            ftntRange.HighlightColorIndex = wdDarkYellow
                        End If
                        GoTo HopeThisWorks
                    End With
            
            End If
        End With
Next

With Selection.Find
    .ClearFormatting
    .MatchWildcards = False
    .Text = "^p^p"
    .Wrap = wdFindContinue
    .Replacement.ClearFormatting
    .Replacement.Text = ""
    
    While .Execute
        If .Found = True Then
            With Selection
                .MoveEndWhile Cset:=vbCr, Count:=wdForward
                .MoveStartWhile Cset:=vbCr, Count:=wdBackward
                .Delete Unit:=wdCharacter, Count:=1
            End With
        End If
    Wend
End With
    
Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub

Public Sub ComboBox2_Click()

On Error GoTo Error

ReDim lstBx(0 To UBound(bkmrks)) As Variant     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim lstCnt As Integer                           ' lstBx() array populates ListBox1, lstCnt counts so add items to lstBx() properly
Dim nullCount As Long                           ' nullCount counts empty items in lstBx() for removal before display.
Dim lstIndex As Integer                         ' lstIndex counts Index from bkmrks() to add to lstBx(), for ref later.
lstCnt = 0 And lstIndex = 0 And nullCount = 1   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                
                                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ListBox1.Clear                                  ' Display list of bookmarked numbering under each pattern found.
                                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For Each bkmrk In bkmrks()
    If bkmrk(2) = ComboBox2.Text Then
        If bkmrk(3) = "(" & Chr(9) Then
            If bkmrk(0) = "042" Then                                    '''''''''''''''''''''''''''''
                lstBx(lstCnt) = Array(bkmrk(1), ChrW(-3880), lstIndex)  ' Special cases for bullets..
                lstCnt = lstCnt + 1                                     '''''''''''''''''''''''''''''
            ElseIf bkmrk(0) = "043" Then
                lstBx(lstCnt) = Array(bkmrk(1), ChrW(-3929), lstIndex)
                lstCnt = lstCnt + 1
            ElseIf bkmrk(0) = "044" Then
                lstBx(lstCnt) = Array(bkmrk(1), ChrW(-3937), lstIndex)
                lstCnt = lstCnt + 1
            ElseIf bkmrk(0) = "045" Then
                lstBx(lstCnt) = Array(bkmrk(1), ChrW(-3988), lstIndex)
                lstCnt = lstCnt + 1
            End If
        Else                                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            lstBx(lstCnt) = Array(bkmrk(1), bkmrk(3), lstIndex) ' Adds info from bkmrks() to ListBox1 builder array lstBx().
            lstCnt = lstCnt + 1                                 ' (Bookmark name, actual character string, bkmrks() index.
        End If                                                  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
    lstIndex = lstIndex + 1
Next bkmrk

For Each i In lstBx()               '''''''''''''''''''''''''''''''''''''''''''''''
    If IsEmpty(i) Then              ' This will remove empty entries from lstBx().
        nullCount = nullCount + 1   '''''''''''''''''''''''''''''''''''''''''''''''
    End If
Next i

ReDim Preserve lstBx(LBound(lstBx) To UBound(lstBx) - nullCount)

For Each lst In lstBx()     '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ListBox1.AddItem lst(1) ' This generates ListBox1 from lstBx() builder array.
Next lst                    '''''''''''''''''''''''''''''''''''''''''''''''''''''

If ComboBox2.Text = "Suspect" Then
    CommandButton3.Visible = True   '''''''''''''''''''''''''''''''''''''''
    ComboBox3.Visible = True        ' Display Suspect options if necessary.
    Label4.Visible = True           '''''''''''''''''''''''''''''''''''''''
End If

Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub

Private Sub ListBox1_Click()
                                                                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
Selection.GoTo What:=wdGoToBookmark, Name:=lstBx(ListBox1.ListIndex)(0) ' Allow navigation of bookmarked numbers for review.
                                                                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Sub BkmrksReassignSus()
                        
On Error GoTo Error
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim nullCount As Long   ' This will count uninitialized members of bkmrks().
susFound = False        ' This will check for Suspects after bkrmks() generation.
Dim noRemoval()         ' This will hold ComboBox2 entries not to remove.
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
For Each i In bkmrks()                  '''''''''''''''''''''''''''''''''''''''''''''''
    If IsEmpty(i) Then                  ' This will remove empty entries from bkmrks().
        nullCount = nullCount + 1       '''''''''''''''''''''''''''''''''''''''''''''''
    End If
Next i
ReDim Preserve bkmrks(LBound(bkmrks) To UBound(bkmrks) - nullCount)

If bkmrks(0)(0) = "Empty" Then GoTo NoEntry
    For Each bkmrk In bkmrks()
        If bkmrk(0) = "Suspect" Then    '''''''''''''''''''''''''''''''''
            susFound = True             ' This will find Suspect entries.
        End If                          '''''''''''''''''''''''''''''''''
    Next bkmrk
NoEntry:

For j = 0 To (ComboBox2.ListCount - 1)                          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If ComboBox2.List(j) = "Suspect" And susFound = False Then  ' This will remove Suspect from ComboBox2 if Suspects have been reassigned already.
        ComboBox2.RemoveItem (j)                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
Next

Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
'Open "C:\ErrorLog.txt" For Output As #n
'Print #n, "Error: " & Err.Number & vbCr & "Description: " & Err.Description & vbCr & "Report errors to GitHub repository."
'Close #n
DestroyProgram

End Sub

Sub BkmrksRedo(crntRedo As String)

On Error GoTo Error

Dim crntCheck As Boolean
crntCheck = False

For c = 0 To (bkmrkNum - 1)
    If bkmrks(c)(2) = crntRedo Then
        crntCheck = True
        Exit Sub
    End If
Next

For j = 0 To (ComboBox2.ListCount - 1)
    If ComboBox2.List(j) = crntRedo And crntCheck = False Then
        ComboBox2.RemoveItem (j)    ' This will remove Suspect from ComboBox2 if Suspects have been reassigned already.
        Exit Sub                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
Next

Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub

Function DoubleLetter(strValue As String) As Boolean
    
On Error GoTo Error
    
    letterFound = ""                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strtPos As Integer              ' This goes through, finds a letter, then goes back through and sees if next letter matches.
    For strtPos = 1 To Len(strValue)    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case Asc(Mid(strValue, strtPos, 1))  ' See if character found is Ascii 65-90, 97-122 (A-Z, a-z).
            Case 65 To 90, 97 To 122                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Mid(strValue, strtPos, 1) = letterFound Then ' Mid function pulls out part of a text string (in this case the bookmark is the string).
                    DoubleLetter = True                         '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Exit For
                Else
                    DoubleLetter = False
                letterFound = Mid(strValue, strtPos, 1)
                End If
            Case Else
                DoubleLetter = False
        End Select
    Next

Exit Function

Error:
MsgBox ("Error: " & Err.Number & vbCr & "Description: " & Err.Description & vbCr & "Report errors to GitHub repository.")
DestroyProgram

End Function

Function TripleLetter(strValue As String) As Boolean

On Error GoTo Error

Dim strtPos As Integer, letterFoundOne As String, letterFoundTwo As String, letterFoundThree As String
Dim foundTwo As Boolean, foundThree As Boolean
letterFoundTwo = ""
letterFoundThree = ""
foundTwo = False
foundThree = False

For strtPos = 1 To Len(strValue)
    If foundTwo = False And foundThree = False Then '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case Asc(Mid(strValue, strtPos, 1))  ' See if character found is Ascii 65-90, 97-122 (A-Z, a-z).
            Case 65 To 90, 97 To 122                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Mid(strValue, strtPos, 1) = letterFoundOne Then  ' Mid function pulls out part of a text string (in this case the bookmark is the string).
                    foundTwo = True                                 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    letterFoundTwo = letterFoundOne
                    GoTo Again
                Else                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    TripleLetter = False                        ' Once letter is found, loops to next char in string. If it matches, foundTwo = True,
                    letterFoundOne = Mid(strValue, strtPos, 1)  ' loops again. If foundTwo = True, check if next is also matching: TripleLetter = True.
                End If                                          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case Else
                TripleLetter = False
        End Select
    ElseIf foundTwo = True And foundThree = False Then  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case Asc(Mid(strValue, strtPos, 1))      ' See if character found is Ascii 65-90, 97-122 (A-Z, a-z).
            Case 65 To 90, 97 To 122                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Mid(strValue, strtPos, 1) = letterFoundTwo Then  ' Mid function pulls out part of a text string (in this case the bookmark is the string).
                    TripleLetter = True                             '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    letterFoundThree = letterFoundTwo
                    Exit Function
                Else                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Exit Function   ' Once letter is found, loops to next char in string. If it matches, foundThree = True,
                End If              ' loops again. If foundThree = True, TripleLetter = True.
            Case Else               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                TripleLetter = False
                Exit Function
        End Select
    End If
Again:
Next

Exit Function

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Function

Function QuadLetter(strValue As String) As Boolean      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                        ' This goes through, finds a letter, then goes back through and sees if next letter matches.
                                                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo Error

Dim strtPos As Integer, letterFoundOne As String, letterFoundTwo As String, letterFoundThree As String, letterFoundFour As String
Dim foundTwo As Boolean, foundThree As Boolean
letterFoundTwo = ""
letterFoundThree = ""
letterFoundFour = ""
foundTwo = False
foundThree = False

For strtPos = 1 To Len(strValue)
    If foundTwo = False And foundThree = False Then '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case Asc(Mid(strValue, strtPos, 1))  ' See if character found is Ascii 65-90, 97-122 (A-Z, a-z).
            Case 65 To 90, 97 To 122                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Mid(strValue, strtPos, 1) = letterFoundOne Then  ' Mid function pulls out part of a text string (in this case the bookmark is the string).
                    foundTwo = True                                 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    letterFoundTwo = letterFoundOne
                    GoTo Again
                Else                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    QuadLetter = False                          ' Once letter is found, loops to next char in string. If it matches, foundTwo = True,
                    letterFoundOne = Mid(strValue, strtPos, 1)  ' loops again. If foundTwo = True, check if next is also matching: TripleLetter = True.
                End If                                          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case Else
                QuadLetter = False
        End Select
    ElseIf foundTwo = True And foundThree = False Then  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case Asc(Mid(strValue, strtPos, 1))      ' See if character found is Ascii 65-90, 97-122 (A-Z, a-z).
            Case 65 To 90, 97 To 122                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Mid(strValue, strtPos, 1) = letterFoundTwo Then  ' Mid function pulls out part of a text string (in this case the bookmark is the string).
                    foundThree = True                               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    letterFoundThree = letterFoundTwo
                    GoTo Again
                Else                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Exit Function   ' Once letter is found, loops to next char in string. If it matches, foundThree = True,
                End If              ' loops again. If foundThree = True, check if next is also matching: QuadLetter = True.
            Case Else               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                QuadLetter = False
        End Select
    ElseIf foundThree = True Then
        If Mid(strValue, strtPos, 1) = letterFoundThree Then
            QuadLetter = True
            letterFoundFour = letterFoundThree
            Exit Function
        Else                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            QuadLetter = False  ' Once letter is found, loops to next char in string. If it matches, foundTwo = True,
            Exit Function       ' loops again. If foundTwo = True, check if next is also matching: TripleLetter = True.
        End If                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
Again:
Next

Exit Function

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Function

Public Function StylesToAdd(listStyle As String)    ''''''''''''''''''''''''''''
                                                    ' StylesToNOTAdd more like..
On Error GoTo Error                                 ''''''''''''''''''''''''''''

Dim stzNobodyCaresAbout()
stzNobodyCaresAbout = Array("ShortOutlineList", "ArticleList", "Hidden_text", "List Double Para,ldp", "List Number", "List Number 2", "List Number 3", "List Number 4", "List Number 5", "List Single Para,lsp", "Ordinal Para,op", "Parties", "Table: Footnote Line,tfl", "Table: Footnote,tf", "TableFootnotesList", "TOC Heading", "FlatBulletsList", "FlatNumbersList", "Default1", "Default2", "Default3", "Default4", "Default5", "Def", "Def 1", "Def 2", "Body Text Center I", "Article / Section", "1 / a / i", "1 / 1.1 / 1.1.1", "(i)(ii)(iii)", "(a)(b)(c)")

For Each sty In stzNobodyCaresAbout
    If sty = listStyle Then         '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Exit Function               ' If style found is one of the list styles, make sure it's not one of the styles no one uses and shouldn't even be in the template anyway (seriously who made these things?).
    End If                          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Next sty

With ComboBox1          ''''''''''''''''''''''''''''''''''''''
    .AddItem listStyle  ' Add to dropdown of 'heading styles'.
End With                ''''''''''''''''''''''''''''''''''''''

Exit Function

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Function

Public Function ReturnSelection(strSelection As Variant) As String

On Error GoTo Error

Dim strtPos As Integer, ltr1 As String, ltr2 As String, ltr3 As String, ltr4 As String, ltr5 As String, ltr6 As String
ltr1 = ""
ltr2 = ""
ltr3 = ""
ltr4 = ""
ltr5 = ""
ltr6 = ""

For strtPos = 1 To Len(strSelection)                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case Asc(Mid(strSelection, strtPos, 1))  ' See if character found is Ascii 65-90, 97-122 (A-Z, a-z).
        Case 65 To 90, 97 To 122                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If ltr1 = "" Then
                ltr1 = Mid(strSelection, strtPos, 1)    ' Mid function pulls out part of a text string (in this case the Selection is the string).
                GoTo Again                              ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ElseIf ltr2 = "" Then
                ltr2 = Mid(strSelection, strtPos, 1)
                GoTo Again
            ElseIf ltr3 = "" Then
                ltr3 = Mid(strSelection, strtPos, 1)
                GoTo Again
            ElseIf ltr4 = "" Then
                ltr4 = Mid(strSelection, strtPos, 1)
                GoTo Again
            ElseIf ltr5 = "" Then
                ltr5 = Mid(strSelection, strtPos, 1)
                GoTo Again
            ElseIf ltr6 = "" Then
                ltr6 = Mid(strSelection, strtPos, 1)
                GoTo Again
            End If
        Case Else
            GoTo Again
    End Select
Again:
Next

ReturnSelection = ltr1 & ltr2 & ltr3 & ltr4 & ltr5 & ltr6

Exit Function

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Function

Public Function IsEven(strBkmrk As Variant) As Boolean  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                        ' Odd will be False, indicate letter. Even will be True, indicate Roman numeral.
On Error GoTo Error                                     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If strBkmrk Mod 2 = 1 Then
    IsEven = False
    Exit Function
Else
    IsEven = True
    Exit Function
End If

Exit Function

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Function

Public Function JustAWord(strValue As String) As Boolean

On Error GoTo Error
                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strtPos As Integer              ' This goes through, finds a letter, checks if it's Roman numeral possible or not.
For strtPos = 1 To Len(strValue)    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case Asc(Mid(strValue, strtPos, 1))                              ' See if character found is ASCII number for (I, V, X, L, C, D, M, i, v, x, l, c, d, m).
        Case 73, 86, 88, 76, 67, 68, 77, 105, 118, 120, 108, 99, 100, 109   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            JustAWord = False
        Case Else
            JustAWord = True
            Exit For
    End Select
Next

Exit Function

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Function

Public Sub BookmarkGW(ptrnGW As Variant, ptrnCode As Integer, ptrnName As Integer)
          
On Error GoTo Error
          
With Selection
    If Right(Selection, 1) <> Chr(9) Then
        .MoveEnd Unit:=wdCharacter, Count:=1
        If Right(Selection, 1) = "." Or Right(Selection, 1) = "," Or Right(Selection, 1) = ":" Or Right(Selection, 1) = ";" Or Right(Selection, 1) = " " Or Right(Selection, 1) = Chr(9) Or Right(Selection, 1) = Chr$(160) Then
            .MoveEnd Unit:=wdCharacter, Count:=-1
        Else
            .Collapse wdCollapseEnd
            Exit Sub
        End If
    End If
End With
          
With ComboBox2
    Dim dup
    dup = False
    For i = 0 To ComboBox2.ListCount - 1
    If ptrnGW(ptrnName) = ComboBox2.List(i) Then dup = True
    Next i                                                  ''''''''''''''''''''''''''''''''''''''
    If dup = False Then ComboBox2.AddItem ptrnGW(ptrnName)  ' Add to dropdown of 'Patterns' found.
End With                                                    ''''''''''''''''''''''''''''''''''''''

With ComboBox3
    Dim dupl
    dupl = False
    For i = 0 To ComboBox3.ListCount - 1
    If ptrnGW(ptrnName) = ComboBox3.List(i) Then dupl = True
    Next i
    If dupl = False Then
        If ptrnGW(ptrnName) <> "Suspect" Then   ''''''''''''''''''''''''''''''''''''''
            ComboBox3.AddItem ptrnGW(ptrnName)  ' Add to dropdown of 'Suspects' found.
        End If                                  ''''''''''''''''''''''''''''''''''''''
    End If
End With

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This portion will set up bookmarks, highlight 'numbering' for review
' Bookmarking the pattern found, minus leading carriage return, plus any ending pattern or periods, spaces, tabs.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
With Selection                                      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    .MoveStart Unit:=wdCharacter, Count:=1          ' Deselect the paramarker at left of selection
    .MoveEndWhile Cset:=".,:; " & Chr$(160) & Chr(9), Count:=4  ' Add ending periods, tabs, whatever to number selection
    .Range.HighlightColorIndex = wdTurquoise        ' Highlight selection turquoise
ptrnFound = Selection                               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
bkmrkName = "skdn" & ptrnGW(ptrnCode) & "x" & bkmrkNum                      '''''''''''''''''''''''''''''''''''''''''''''''
    ActiveDocument.bookmarks.Add Name:=bkmrkName, Range:=Selection.Range    ' Bookmark selection for later review/deletion
                                                                            '''''''''''''''''''''''''''''''''''''''''''''''
                                                                                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If bkmrkNum > UBound(bkmrks) - 1 Then ReDim Preserve bkmrks(0 To UBound(bkmrks) + 100)  ' Resizes array bkmrks() as number of patterns is found, adds entries.
bkmrks(bkmrkNum) = Array(ptrnGW(ptrnCode), bkmrkName, ptrnGW(ptrnName), ptrnFound)      ' Pattern code, bookmark name, bookmark pattern style, actual number.
                                                                                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
bkmrkNum = bkmrkNum + 1         ' Add 1 to bookmark name number so they don't duplicate.
    .Collapse wdCollapseEnd     ' Collapse selection after each pattern is found so next can be.
End With                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If IsEven(ptrnGW(ptrnCode)) = True Then
    lastRoman = ReturnSelection(ptrnFound)  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Else                                        ' Note what last letter/Roman numeral was, for testing. This is best test.
    lastLetter = ReturnSelection(ptrnFound) ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub
                                                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DifferentiateSelection(ptrnGW As Variant)    ' This is a nightmare. Checks whether Selection is Roman numeral or letter.
                                                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo Error
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' (I, V, X, L, C, D, M, i, v, x, l, c, d, m) in ASCII    |    0, 4, 9, 49, 99, 499, 999 in Roman numerals.
Dim romNum()    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
romNum = Array(Array(73, "nulla"), Array(86, "IV"), Array(88, "IX"), Array(76, "XLIX"), Array(67, "XCIX"), Array(68, "CDXCIX"), Array(77, "CMXCIX"), _
            Array(105, "nulla"), Array(118, "iv"), Array(120, "ix"), Array(108, "xlix"), Array(99, "xcix"), Array(100, "cdxcix"), Array(109, "cmxcix"))
            
Dim oneAgo3 As String, oneAgo0 As String, twoAgo0 As String, twoAgo3 As String, crntDbl As String
Dim fourAgo0 As String, fourAgo3 As String, crntRedo As String, ltrSelection As String

ltrSelection = ReturnSelection(Selection)

If bkmrkNum = 0 Then
    oneAgo3 = "Empty"   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    oneAgo0 = "Empty"   ' Set a fake first bookmark because I need to rely on previous bookmarks to determine whether Roman numeral or letter,
    twoAgo3 = "Empty"   ' but rarely there may be no previous bookmark to test (which would throw error). A better way exists, but I'm leaving this as is.
    twoAgo0 = "Empty"   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    fourAgo3 = "Empty"
    fourAgo0 = "Empty"
    crntRedo = "Empty"
ElseIf bkmrkNum = 1 Or bkmrkNum = 2 Then
    oneAgo3 = ReturnSelection(bkmrks(bkmrkNum - 1)(3))
    oneAgo0 = bkmrks(bkmrkNum - 1)(0)
    twoAgo3 = "Empty"
    twoAgo0 = "Empty"
    fourAgo3 = "Empty"
    fourAgo0 = "Empty"
    crntRedo = bkmrks(bkmrkNum - 1)(2)
ElseIf bkmrkNum = 3 Then
    oneAgo3 = ReturnSelection(bkmrks(bkmrkNum - 1)(3))
    oneAgo0 = bkmrks(bkmrkNum - 1)(0)
    twoAgo3 = ReturnSelection(bkmrks(bkmrkNum - 2)(3))
    twoAgo0 = bkmrks(bkmrkNum - 2)(0)
    fourAgo3 = "Empty"
    fourAgo0 = "Empty"
    crntRedo = bkmrks(bkmrkNum - 1)(2)
ElseIf bkmrkNum > 3 Then
    oneAgo3 = ReturnSelection(bkmrks(bkmrkNum - 1)(3))
    oneAgo0 = bkmrks(bkmrkNum - 1)(0)
    twoAgo3 = ReturnSelection(bkmrks(bkmrkNum - 2)(3))
    twoAgo0 = bkmrks(bkmrkNum - 2)(0)
    fourAgo3 = ReturnSelection(bkmrks(bkmrkNum - 4)(3))
    fourAgo0 = bkmrks(bkmrkNum - 4)(0)
    crntRedo = bkmrks(bkmrkNum - 1)(2)
End If
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This whole section looks at numbering patterns, decides if it's a Roman numeral
' or a letter based on length, what surrounds it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Select Case Len(ltrSelection) ' Based on how long Selection is (how many characters), do the following:
    Case Is > 3               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If QuadLetter(ltrSelection) = True Then ' If it's 4 or more letters that are all the same, they're letters.
            BookmarkGW ptrnGW, "0", "3"         '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Exit Sub
        Else
            If JustAWord(ltrSelection) = True Then
                Exit Sub
            Else                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                BookmarkGW ptrnGW, "1", "4" ' If it's 4 or more letters that aren't all the same, must be Roman Numerals.
                Exit Sub                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End If
        End If
    Case Is = 2, 3                                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If DoubleLetter(ltrSelection) = False Then  ' If it's 2 or 3 letters, test if there are double letters (see Sub DoubleLetter() for details). Word only does
            If JustAWord(ltrSelection) = True Then  ' repeating letters for letters past Z, starting at AA; if it's not same two letters in a row, so long as it isn't
                Exit Sub                            ' just a random word (test if Selection contains only Roman numeral possible characters) it must be Roman numerals.
            Else                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                BookmarkGW ptrnGW, "1", "4"
                Exit Sub
            End If
        ElseIf Len(ltrSelection) = 3 And TripleLetter(ltrSelection) = False Then
            BookmarkGW ptrnGW, "1", "4"
            Exit Sub                                                                                '''''''''''''''''''''''''''''''''''''''''''''
        ElseIf DoubleLetter(ltrSelection) = True And letterFound <> "J" And letterFound <> "j" Then ' Special case for double and triple J below.
            crntDbl = letterFound               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            For Each rom In romNum              ' DoubleLetters return true, could be II, III, XX Roman numerals, or any double letter otherwise, so:
                If crntDbl = Chr(rom(0)) Then   ' If last letter in DoubleLetter() found is a Roman numeral possible number, continue testing:
                    If oneAgo3 = Chr(rom(0) - 1) & Chr(rom(0) - 1) Or oneAgo3 = Chr(rom(0) - 1) & Chr(rom(0) - 1) & Chr(rom(0) - 1) Then
                        BookmarkGW ptrnGW, "0", "3"     ' If last identified bookmark is previous letter of alphabet, must be a letter.
                        Exit Sub                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ElseIf crntDbl = "x" Or crntDbl = "X" Then              ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If Len(ltrSelection) = 2 And Len(oneAgo3) = 3 Then  ' If Selection is 2 characters and last was 3, might be XIX, XX.
                            If oneAgo3 = "xix" Or oneAgo3 = "XIX" Then      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                BookmarkGW ptrnGW, "1", "4"
                                Exit Sub
                            Else
                                BookmarkGW ptrnGW, "5", "5" ' 0, 3?
                                Exit Sub
                            End If                                              ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ElseIf Len(ltrSelection) = 3 And Len(oneAgo3) = 4 Then  ' If Selection is 3 characters and last was 4, might be XXIX, XXX.
                            If oneAgo3 = "xxix" Or oneAgo3 = "XXIX" Then        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                BookmarkGW ptrnGW, "1", "4"
                                Exit Sub
                            Else
                                BookmarkGW ptrnGW, "5", "5" ' 0, 3?
                                Exit Sub
                            End If
                        Else
                            BookmarkGW ptrnGW, "0", "3"
                            Exit Sub
                        End If
                    ElseIf crntDbl = "i" Or crntDbl = "I" Then                      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If Len(ltrSelection) = 2 And IsEven(oneAgo0) = True Then    ' If Selection is 2 characters AND if last identified bookmark is a Roman numeral..
                            If oneAgo3 = "i" Or oneAgo3 = "I" Then                  ' ..and if last is same character..
                                BookmarkGW ptrnGW, "1", "4"                         ' ..must be Roman numerals II, I.
                                Exit Sub                                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            Else
                                BookmarkGW ptrnGW, "0", "3"
                                Exit Sub
                            End If
                        ElseIf Len(ltrSelection) = 2 And IsEven(oneAgo0) = False Then   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            If oneAgo3 = "i" Or oneAgo3 = "I" Then                      ' If Selection is 2 letters AND if last identified bookmark is same char AND a letter..
                                bkmrks(bkmrkNum - 1) = Array(ptrnGW(1), bkmrks(bkmrkNum - 1)(1), ptrnGW(4), bkmrks(bkmrkNum - 1)(3)) ' ..must be Roman numerals II, I; correcting.
                                BookmarkGW ptrnGW, "1", "4"                                                                          '''''''''''''''''''''''''''''''''''''''''''''
                                BkmrksRedo (crntRedo)
                                lastLetter = twoAgo3
                                Exit Sub
                            ElseIf oneAgo3 = "hh" Or oneAgo3 = "HH" Then
                                BookmarkGW ptrnGW, "0", "3"
                                Exit Sub
                            Else
                                BookmarkGW ptrnGW, "1", "4"
                                Exit Sub
                            End If
                        ElseIf Len(ltrSelection) = 3 And IsEven(oneAgo0) = True Then
                            If oneAgo3 = "ii" Or oneAgo3 = "II" Then
                                BookmarkGW ptrnGW, "1", "4"     ' If Selection is 3 letters AND last was same char AND a Roman numeral, must be Roman III.
                                Exit Sub                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            End If
                        ElseIf Len(ltrSelection) = 3 And IsEven(oneAgo0) = False Then
                            If lastRoman = "ii" Or lastRoman = "II" Then
                                If lastLetter = "hhh" Or lastLetter = "HHH" Then
                                    If oneAgo3 = "hhh" Or oneAgo3 = "HHH" Then
                                        BookmarkGW ptrnGW, "0", "3"
                                        Exit Sub
                                    End If
                                Else
                                    BookmarkGW ptrnGW, "1", "4"
                                    Exit Sub
                                End If
                            ElseIf lastLetter = "ii" Or lastLetter = "II" Then
                                If oneAgo3 = "ii" Or oneAgo3 = "II" Then
                                    bkmrks(bkmrkNum - 1) = Array(ptrnGW(1), bkmrks(bkmrkNum - 1)(1), ptrnGW(4), bkmrks(bkmrkNum - 1)(3)) ' ..must be Roman numerals III, II; correcting.
                                    BookmarkGW ptrnGW, "1", "4"                                                                          '''''''''''''''''''''''''''''''''''''''''''''''
                                    BkmrksRedo (crntRedo)
                                    Exit Sub
                                End If
                            End If
                        End If
                    ElseIf crntDbl = "c" Or crntDbl = "C" Or crntDbl = "M" Or crntDbl = "m" Then
                        If oneAgo3 <> rom(1) Then       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            BookmarkGW ptrnGW, "0", "3" ' Check the romNum() Array and see if last bookmark is previous Roman numeral.
                            Exit Sub                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Else
                            BookmarkGW ptrnGW, "1", "4"
                            Exit Sub
                        End If
                    Else                            '''''''''''''''''''''''''''''''
                        BookmarkGW ptrnGW, "0", "3" ' Letter if not one of above..
                        Exit Sub                    '''''''''''''''''''''''''''''''
                    End If  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                End If      ' If DoubleLetter() is True, but not one of the above, letter.
            Next rom        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            BookmarkGW ptrnGW, "0", "3"
        ElseIf DoubleLetter(ltrSelection) = True And Len(ltrSelection) = 2 And letterFound = "j" Or DoubleLetter(ltrSelection) = True And Len(ltrSelection) = 2 And letterFound = "J" Then
            If oneAgo3 = "II" Or oneAgo3 = "ii" Then
                If twoAgo3 = "I" Or twoAgo3 = "i" Then
                    If fourAgo3 = "GG" Or fourAgo3 = "gg" Then
                        bkmrks(bkmrkNum - 1) = Array(ptrnGW(0), bkmrks(bkmrkNum - 1)(1), ptrnGW(3), bkmrks(bkmrkNum - 1)(3))
                        BookmarkGW ptrnGW, "0", "3" ' If this is JJ or jj, and if last was II or ii but was Roman numeral, and four ago was not ii or II and two ago was i or ii, last was letter, correcting.
                        BkmrksRedo (crntRedo)       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Exit Sub
                    End If
                End If
            End If
            BookmarkGW ptrnGW, "0", "3"
        Else
            If oneAgo3 = "III" Or oneAgo3 = "iii" Then
                If IsEven(oneAgo0) = True Then
                    If twoAgo3 <> "III" And twoAgo3 <> "iii" Then
                        If fourAgo3 <> "III" And fourAgo3 <> "iii" Then
                            bkmrks(bkmrkNum - 1) = Array(ptrnGW(0), bkmrks(bkmrkNum - 1)(1), ptrnGW(3), bkmrks(bkmrkNum - 1)(3))
                            BookmarkGW ptrnGW, "0", "3" ' If this is JJJ or jjj, and if last was III or iii but was Roman numeral, but four ago was not iii or III and two ago wasn't ii or iii, last was letter, correcting.
                            BkmrksRedo (crntRedo)       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            Exit Sub
                        ElseIf fourAgo3 = "III" Or fourAgo3 = "iii" Then
                            BookmarkGW ptrnGW, "0", "3" ' If this is JJJ or jjj, and if last was III or iii but was Roman numeral and four ago was iii or III, but two ago wasn't ii or iii, last was Roman, don't change.
                            Exit Sub                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        End If
                    End If
                ElseIf oneAgo0 = "Suspect" Then
                    bkmrks(bkmrkNum - 1) = Array(ptrnGW(0), bkmrks(bkmrkNum - 1)(1), ptrnGW(3), bkmrks(bkmrkNum - 1)(3))
                    BookmarkGW ptrnGW, "0", "3"     ' If this is JJJ or jjj, and if last was Suspect, last should have been letter, correcting.
                    BkmrksRedo (crntRedo)           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Exit Sub
                Else                            '''''''''''''''''''''''''''''''''''''
                    BookmarkGW ptrnGW, "0", "3" ' If nothing else, make JJJ a letter.
                    Exit Sub                    '''''''''''''''''''''''''''''''''''''
                End If
            Else                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                BookmarkGW ptrnGW, "0", "3"     ' Else, if this is jjj or JJJ, last was a letter, just a letter.
                Exit Sub                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End If
        End If                      ''''''''''''''''''''''
    Case Is = 1                     ' If it's 1 character.
        For Each rom In romNum      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If ltrSelection = Chr(rom(0)) Then     ' If Selection is a Roman numeral character then..
                If oneAgo3 = Chr(rom(0) - 1) Then  ' ..if last bookmark is previous letter in alphabet..
                    BookmarkGW ptrnGW, "0", "3"    ' ..must be a letter.
                    Exit Sub                       '''''''''''''''''''''''''''''''''''''''''''''''''''''
                Else                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If oneAgo3 = rom(1) Then        ' But if last bookmark does equal previous Roman numeral..
                        BookmarkGW ptrnGW, "1", "4" ' ..must be Roman numeral
                        Exit Sub                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Else                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If rom(1) <> "nulla" Then       ' ..if previous Roman numeral does NOT equal 'nulla'..
                            BookmarkGW ptrnGW, "0", "3" ' ..must be letter.
                            Exit Sub                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Else                                    ' If previous Roman numeral is 'nulla' (if we're looking at an I basically)..
                            If (oneAgo0) = ptrnGW(1) Then       ' ..if last pattern found was also a Roman numeral..
                                If lastLetter = "h" Or lastLetter = "H" Then    ' ..If last letter found was H..
                                    BookmarkGW ptrnGW, "0", "3"                 ' ..must be letter I.
                                    Exit Sub                                    '''''''''''''''''''''
                                Else                            '''''''''''''''''''''''''''''''''''''''''''''''''''''
                                    BookmarkGW ptrnGW, "1", "4" ' ..If last letter wasn't H, must be another Roman I.
                                    Exit Sub                    '''''''''''''''''''''''''''''''''''''''''''''''''''''
                                End If                          '''''''''''''''''''''''''''''''''''''''
                            ElseIf oneAgo0 = ptrnGW(5) Then     ' ..if last pattern found was Suspect..
                                BookmarkGW ptrnGW, "1", "4"     ' ..PROBABLY a Roman numeral.
                                Exit Sub                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            Else                            ' If last pattern found was a letter..
                                BookmarkGW ptrnGW, "1", "4" ' ..Roman numeral (the case where H, I, I, II, J occurs).
                                Exit Sub                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            End If
                        End If
                    End If
                End If
            End If
        Next rom
    If oneAgo3 = "v" Or oneAgo3 = "V" Then
        If lastLetter = "U" And lastRoman = "V" Or lastLetter = "u" And lastRoman = "v" Or lastLetter = "u" And lastRoman = "V" Or lastLetter = "U" And lastRoman = "v" Then
            bkmrks(bkmrkNum - 1) = Array(ptrnGW(0), bkmrks(bkmrkNum - 1)(1), ptrnGW(3), bkmrks(bkmrkNum - 1)(3))
            BookmarkGW ptrnGW, "0", "3"
            BkmrksRedo (crntRedo)
            Exit Sub
        Else
            BookmarkGW ptrnGW, "0", "3"
            Exit Sub
        End If
    ElseIf oneAgo3 = "x" Or oneAgo3 = "X" Then
        If lastLetter = "w" And lastRoman = "x" Or lastLetter = "W" And lastRoman = "X" Or lastLetter = "w" And lastRoman = "X" Or lastLetter = "W" And lastRoman = "x" Then
            bkmrks(bkmrkNum - 1) = Array(ptrnGW(0), bkmrks(bkmrkNum - 1)(1), ptrnGW(3), bkmrks(bkmrkNum - 1)(3))
            BookmarkGW ptrnGW, "0", "3"
            BkmrksRedo (crntRedo)
            Exit Sub
        Else
            BookmarkGW ptrnGW, "0", "3"
            Exit Sub
        End If
    ElseIf ltrSelection = "o" Then
        If lastLetter = "n" Then
            BookmarkGW ptrnGW, "0", "3"
            Exit Sub
        Else
            BookmarkGW ptrnGW, "6", "7"
            Exit Sub
        End If
    Else                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    BookmarkGW ptrnGW, "0", "3" ' If Selection isn't a Roman numeral character, obviously a letter.
    Exit Sub                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
End Select

Exit Sub

Error:
MsgBox ("Error: " & Err.Number & "." & vbCr & "Description: " & Err.Description & "." & vbCr & "Report errors to GitHub repository." & "." & vbCr & "Unloading program.")
DestroyProgram

End Sub

Sub GoAheadAndRemoveThen()
    
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Dim rngCounter As Integer, frmt As ParagraphFormat ', fnt As Font
    Dim rngArray()
    ReDim rngArray(0 To 10000) As Variant
    
    Selection.HomeKey Unit:=wdStory
    
    For Each par In ActiveDocument.Paragraphs
        If IsStyleSeparator = True Then
        With Selection.Find
            .ClearFormatting
            .Text = "^p"
            .Forward = True
            .Font.Hidden = True
            .Wrap = wdFindContinue
        While Selection.Find.Execute
            If .Found = True Then
                Selection.Delete
            End If
        Wend
        End With
        End If
    Next
    
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Font.Bold = True
        '.Font.Italic = True
        '.Font.Underline = wdUnderlineSingle
        '.Font.Superscript = True
        '.Font.Subscript = True
        '.Font.StrikeThrough = True
        '.Font.SmallCaps = True
        '.Font.AllCaps = True
        '.Font.Hidden = True
        
    While Selection.Find.Execute                                    '''''''''''''''''''''''''''''''
        If .Found = True Then                                       ' This tracks bold font ranges.
            rngArray(rngCounter) = Array(Selection.Range, "Bold")   '''''''''''''''''''''''''''''''
            rngCounter = rngCounter + 1
        End If
    Wend        ''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Keep going until all entries searched for/found
    End With    ''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        '.Font.Bold = True
        .Font.Italic = True
        '.Font.Underline = wdUnderlineSingle
        '.Font.Superscript = True
        '.Font.Subscript = True
        '.Font.StrikeThrough = True
        '.Font.SmallCaps = True
        '.Font.AllCaps = True
        '.Font.Hidden = True
        
    While Selection.Find.Execute                                    ''''''''''''''''''''''''''''
        If .Found = True Then                                       ' Tracks italic font ranges.
            rngArray(rngCounter) = Array(Selection.Range, "Italic") ''''''''''''''''''''''''''''
            rngCounter = rngCounter + 1
        End If
    Wend        ''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Keep going until all entries searched for/found
    End With    ''''''''''''''''''''''''''''''''''''''''''''''''''

    
    For Each prgrph In ActiveDocument.Paragraphs
        With prgrph
            If .Style <> ActiveDocument.Styles("Normal") Then
                'Set fnt = .Style.Font
                Set frmt = .Style.ParagraphFormat
                .Style = ActiveDocument.Styles("Normal")
                '.Range.Font = fnt
                .Range.ParagraphFormat = frmt
            End If
        End With
    Next prgrph
        
    ReDim Preserve rngArray(0 To rngCounter - 1)
    
    For Each rng In rngArray
        If rng(1) = "Bold" Then
            With rng(0)
                .Font.Bold = True
            End With
        ElseIf rng(1) = "Italic" Then
            With rng(0)
                .Font.Italic = True
            End With
        End If
    Next rng


End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Send all questions/comments/error reporting to GitHub repository.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Extra cases in case someone is more ambitious than I.
                'Array("^p^#^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#", "01.01.01.01.01.01.01.01.01"), Array("^p^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#", "1.01.01.01.01.01.01.01.01"), Array("^p^#.^#.^#.^#.^#.^#.^#.^#.^#", "1.1.1.1.1.1.1.1.1"), _
                'Array("^p^#^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#", "01.01.01.01.01.01.01.01"), Array("^p^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#", "1.01.01.01.01.01.01.01"), Array("^p^#.^#.^#.^#.^#.^#.^#.^#", "1.1.1.1.1.1.1.1"), _
                'Array("^p^#^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#", "01.01.01.01.01.01.01"), Array("^p^#.^#^#.^#^#.^#^#.^#^#.^#^#.^#^#", "1.01.01.01.01.01.01"), Array("^p^#.^#.^#.^#.^#.^#.^#", "1.1.1.1.1.1.1"), _
                'Array("^p^#^#.^#^#.^#^#.^#^#.^#^#.^#^#", "01.01.01.01.01.01"), Array("^p^#.^#^#.^#^#.^#^#.^#^#.^#^#", "1.01.01.01.01.01"), Array("^p^#.^#.^#.^#.^#.^#", "1.1.1.1.1.1"), _
                'Array("^p^#^#.^#^#.^#^#.^#^#.^#^#", "01.01.01.01.01"), Array("^p^#.^#^#.^#^#.^#^#.^#^#", "1.01.01.01.01"), Array("^p^#.^#.^#.^#.^#", "1.1.1.1.1"), _
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Extra cases in case it becomes important to differentiate between 1.1 and 1.01 to users (this is from before I came to my senses; I don't think it is).
                'Array("001", "^p^#^#.^#^#.^#^#.^#^#", "01.01.01.01"), Array("002", "^p^#.^#^#.^#^#.^#^#", "1.01.01.01"), Array("003", "^p^#.^#.^#.^#", "1.1.1.1"), _
                'Array("003", "^p^#^#.^#^#.^#^#", "01.01.01"), Array("004", "^p^#.^#^#.^#^#", "1.01.01"), Array("005", "^p^#.^#.^#", "1.1.1"), _
                'Array("006", "^p^#^#.^#^#", "01.01"), Array("007", "^p^#.^#^#", "1.01"), Array("008", "^p^#.^#", "1.1"), _
                'Array("009", "^p^#^#", "01"), Array("010", "^p^#", "1"), _
                'Array("^pArticle ^#^#.^#^#", "Article 01.01"), Array("^pArticle ^#.^#^#", "Article 1.01"), Array("^pArticle ^#^#.^#", "Article 01.1"), Array("^pArticle ^#.^#", "Article 1.1"), Array("^pArticle ^#^#", "Article 01"), Array("^pArticle ^#", "Article 1"), _
                'Array("^pArticle ^$^$^$^$^$", "Article I"), Array("^pArticle ^$^$^$^$", "Article I"), Array("^pArticle ^$^$^$", "Article I"), Array("^pArticle ^$^$", "Article I"), Array("^pArticle ^$", "Article I"), _
                'Array("^pSection ^#^#.^#^#", "Section 01.01"), Array("^pSection ^#.^#^#", "Section 1.01"), Array("^pSection ^#^#.^#", "Section 01.1"), Array("^pSection ^#.^#", "Section 1.1"), Array("^pSection ^#^#", "Section 01"), Array("^pSection ^#", "Section 1"), Array("^pSection ^$^$", "Section I"), Array("^pSection ^$", "Section I"), _
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'Sub testDialogOpen()
'    Dim wHandle As Long
'    Dim wName As String
'
'    wName = "Find and Replace"
'    wHandle = FindWindow(0&, wName)
'    If wHandle <> 0 Then
'        MsgBox "Dialog window is open, closing.."
'    End If
'End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''
