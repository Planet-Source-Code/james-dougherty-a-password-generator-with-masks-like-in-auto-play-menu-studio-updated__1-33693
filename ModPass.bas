Attribute VB_Name = "PassGen"
Option Explicit

Public Function GeneratePasswords(List As ListBox, NumPasswords As Long, Mask As String, Optional ClearList As Boolean = True)
Dim i As Integer
Dim j As Integer
Dim z As Integer
Dim tmpMask As String
Dim hMask As String

'Good idea to clear list so we dont get some kind of
'overflow error
If ClearList Then List.Clear

'Get a random seed value
Randomize

'How many passwords do we want to create?
For j = 0 To NumPasswords - 1
 'Reset this to "" or it will double like, Password1Password2
 'we want Password1
 '        Password2
 hMask = vbNullString
 'Step through every character in the string
 For i = 1 To (Len(Mask))
  'We step through the characters according to where we are in the loop
  tmpMask$ = Mid$(Mask, i, 1)
  'If the character is "#" then we want a random number
  If tmpMask$ = "#" Then
   'Create a random number
   tmpMask$ = CInt(Rnd * 9)
   'Add to our hold mask(just a temp string)
   hMask$ = hMask$ & tmpMask$
  'If the character is "X" then we want a random character
  ElseIf tmpMask$ = "X" Then
   'Create a random character
   tmpMask$ = Chr((Int((90 - 65 + 1) * Rnd + 65)))
   'Add to our hold mask
   hMask$ = hMask$ & tmpMask$
  Else
   'If another character "-" "CDKEY" whatever ignore it
   tmpMask$ = tmpMask$
   'Just add it to our hold mask
   hMask$ = hMask$ & tmpMask$
  End If
 Next i
 'After the loop has went through every character our hold mask
 'should now contain our full password so add it the list
 
 'UPDATED - NOW WE ONLY ADD THE PASSWORD IF IT IS UNIQUE.
 
 'START FROM 0 TO OUR CURRENT INDEX AND MAKE SURE IT IS UNIQUE.
 For z = 0 To j
  If hMask$ <> CStr(List.List(z)) Then
   'Its unique keep it the same
   hMask$ = hMask$
  Else
   'Not unique make it null ""
   hMask$ = ""
  End If
 Next z
 'If the mask was unique and not null we add it
 If hMask$ <> "" Then List.AddItem hMask$
Next j
End Function
