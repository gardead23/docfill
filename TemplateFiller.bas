Attribute VB_Name = "TemplateFiller"
' ============================================================
' Template Filler — Word VBA Macro
' ============================================================
' Usage:
'   1. Open a Word document with {{placeholder}} fields
'   2. Run this macro (Developer > Macros > FillTemplate > Run)
'   3. Fill in each field when prompted
'   4. Done — all placeholders are replaced in the document
' ============================================================

Sub FillTemplate()
    Dim doc As Document
    Set doc = ActiveDocument

    ' --- Find all {{placeholder}} keys in the document ---
    Dim placeholderList() As String
    Dim placeholderCount As Integer
    placeholderCount = 0

    Dim fullText As String
    fullText = doc.Content.Text

    Dim i As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim key As String

    i = 1
    Do While i <= Len(fullText) - 3
        If Mid(fullText, i, 2) = "{{" Then
            startPos = i + 2
            endPos = InStr(startPos, fullText, "}}")
            If endPos > startPos Then
                key = Mid(fullText, startPos, endPos - startPos)
                If IsValidKey(key) And Not AlreadyFound(placeholderList, placeholderCount, key) Then
                    ReDim Preserve placeholderList(placeholderCount)
                    placeholderList(placeholderCount) = key
                    placeholderCount = placeholderCount + 1
                End If
                i = endPos + 2
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop

    ' --- Nothing found ---
    If placeholderCount = 0 Then
        MsgBox "No {{placeholders}} found in this document." & Chr(13) & Chr(13) & _
               "Add fields like {{client_name}} or {{date}} to your document, then run this macro again.", _
               vbInformation, "Template Filler"
        Exit Sub
    End If

    ' --- Prompt user for each value ---
    Dim values() As String
    ReDim values(placeholderCount - 1)

    Dim j As Integer
    For j = 0 To placeholderCount - 1
        Dim label As String
        label = FriendlyLabel(placeholderList(j))

        Dim prompt As String
        prompt = label & Chr(13) & Chr(13) & _
                 "Field " & (j + 1) & " of " & placeholderCount & "   [Cancel to abort]"

        Dim val As String
        val = InputBox(prompt, "Template Filler")

        ' Empty + Cancel are both "" in Mac VBA — we just allow empty values
        values(j) = val
    Next j

    ' --- Confirm before replacing ---
    Dim summary As String
    summary = "Ready to fill " & placeholderCount & " field(s):" & Chr(13) & Chr(13)
    For j = 0 To placeholderCount - 1
        Dim displayVal As String
        displayVal = values(j)
        If Len(displayVal) = 0 Then displayVal = "(empty — will be skipped)"
        summary = summary & "  " & FriendlyLabel(placeholderList(j)) & ": " & displayVal & Chr(13)
    Next j
    summary = summary & Chr(13) & "Proceed?"

    If MsgBox(summary, vbYesNo + vbQuestion, "Template Filler") = vbNo Then
        MsgBox "Cancelled. No changes were made.", vbInformation, "Template Filler"
        Exit Sub
    End If

    ' --- Replace all placeholders ---
    Dim totalReplaced As Long
    totalReplaced = 0

    For j = 0 To placeholderCount - 1
        If Len(values(j)) > 0 Then
            With doc.Content.Find
                .ClearFormatting
                .Text = "{{" & placeholderList(j) & "}}"
                .Replacement.ClearFormatting
                .Replacement.Text = values(j)
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = True
                .Execute Replace:=wdReplaceAll
            End With
            totalReplaced = totalReplaced + 1
        End If
    Next j

    MsgBox "Done! " & totalReplaced & " field(s) filled." & Chr(13) & Chr(13) & _
           "Use Cmd+Z to undo if needed.", _
           vbInformation, "Template Filler"
End Sub

' -------------------------------------------------------
' Check if a key has already been found (deduplication)
' -------------------------------------------------------
Private Function AlreadyFound(arr() As String, count As Integer, key As String) As Boolean
    If count = 0 Then
        AlreadyFound = False
        Exit Function
    End If
    Dim k As Integer
    For k = 0 To count - 1
        If arr(k) = key Then
            AlreadyFound = True
            Exit Function
        End If
    Next k
    AlreadyFound = False
End Function

' -------------------------------------------------------
' Validate that a key contains only letters, numbers, _
' -------------------------------------------------------
Private Function IsValidKey(s As String) As Boolean
    If Len(s) = 0 Or Len(s) > 60 Then
        IsValidKey = False
        Exit Function
    End If
    Dim i As Integer
    For i = 1 To Len(s)
        If Not (Mid(s, i, 1) Like "[A-Za-z0-9_]") Then
            IsValidKey = False
            Exit Function
        End If
    Next i
    IsValidKey = True
End Function

' -------------------------------------------------------
' Convert snake_case to Title Case for display
' -------------------------------------------------------
Private Function FriendlyLabel(s As String) As String
    FriendlyLabel = StrConv(Replace(s, "_", " "), vbProperCase)
End Function
