<div align="center">

## rtf2html\-2\.1


</div>

### Description

This code recieves RTF code as output by a Rich Text Box in VB or MS Word. It outputs the equivalent in HTML. It's in a somewhat BETA form in that it handles a number of but not all of the possible codes. If you encounter a code it doesn't properly convert just send it to me and I'll try to fix the function within 24 hours. I think it does a better job on uncomplicated text than MS Word's HTML conversion.
 
### More Info
 
String containing rich text to convert. Note: Currently the input must include the Rich-text header codes otherwise the function will return an empty string.

This function may get updated fairly regularly for awhile. Please download the file at the URL below for the latest version:

<A href="http://www2.bitstream.net/~bradyh/downloads/rtf2html.zip">rtf2html.zip</a>

Here's an example of how to use the function with a rich text box (Note that the function also be used with rich text files.)

TextBoxHTML.Text = (RTF2HTML(TextBoxRTF.TextRTF))

String containing HTML code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brady Hegberg](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brady-hegberg.md)
**Level**          |Unknown
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brady-hegberg-rtf2html-2-1__1-1545/archive/master.zip)

### API Declarations

None


### Source Code

```
Function RTF2HTML(strRTF As String) As String
  'Version 2.1 (3/30/99)
  'The most current version of this function is available at
  'http://www2.bitstream.net/~bradyh/downloads/rtf2html.zip
  'Converts Rich Text encoded text to HTML format
  'if you find some text that this function doesn't
  'convert properly please email the text to
  'bradyh@bitstream.net
  Dim strHTML As String
  Dim l As Long
  Dim lTmp As Long
  Dim lRTFLen As Long
  Dim lBOS As Long         'beginning of section
  Dim lEOS As Long         'end of section
  Dim strTmp As String
  Dim strTmp2 As String
  Dim strEOS            'string to be added to end of section
  Const gHellFrozenOver = False  'always false
  Dim gSkip As Boolean       'skip to next word/command
  Dim strCodes As String      'codes for ascii to HTML char conversion
  strCodes = "&nbsp; {00}&copy; {a9}&acute; {b4}&laquo; {ab}&raquo; {bb}&iexcl; {a1}&iquest;{bf}&Agrave;{c0}&agrave;{e0}&Aacute;{c1}"
  strCodes = strCodes & "&aacute;{e1}&Acirc; {c2}&acirc; {e2}&Atilde;{c3}&atilde;{e3}&Auml; {c4}&auml; {e4}&Aring; {c5}&aring; {e5}&AElig; {c6}"
  strCodes = strCodes & "&aelig; {e6}&Ccedil;{c7}&ccedil;{e7}&ETH;  {d0}&eth;  {f0}&Egrave;{c8}&egrave;{e8}&Eacute;{c9}&eacute;{e9}&Ecirc; {ca}"
  strCodes = strCodes & "&ecirc; {ea}&Euml; {cb}&euml; {eb}&Igrave;{cc}&igrave;{ec}&Iacute;{cd}&iacute;{ed}&Icirc; {ce}&icirc; {ee}&Iuml; {cf}"
  strCodes = strCodes & "&iuml; {ef}&Ntilde;{d1}&ntilde;{f1}&Ograve;{d2}&ograve;{f2}&Oacute;{d3}&oacute;{f3}&Ocirc; {d4}&ocirc; {f4}&Otilde;{d5}"
  strCodes = strCodes & "&otilde;{f5}&Ouml; {d6}&ouml; {f6}&Oslash;{d8}&oslash;{f8}&Ugrave;{d9}&ugrave;{f9}&Uacute;{da}&uacute;{fa}&Ucirc; {db}"
  strCodes = strCodes & "&ucirc; {fb}&Uuml; {dc}&uuml; {fc}&Yacute;{dd}&yacute;{fd}&yuml; {ff}&THORN; {de}&thorn; {fe}&szlig; {df}&sect; {a7}"
  strCodes = strCodes & "&para; {b6}&micro; {b5}&brvbar;{a6}&plusmn;{b1}&middot;{b7}&uml;  {a8}&cedil; {b8}&ordf; {aa}&ordm; {ba}&not;  {ac}"
  strCodes = strCodes & "&shy;  {ad}&macr; {af}&deg;  {b0}&sup1; {b9}&sup2; {b2}&sup3; {b3}&frac14;{bc}&frac12;{bd}&frac34;{be}&times; {d7}"
  strCodes = strCodes & "&divide;{f7}&cent; {a2}&pound; {a3}&curren;{a4}&yen;  {a5}"
  strHTML = ""
  lRTFLen = Len(strRTF)
  'seek first line with text on it
  lBOS = InStr(strRTF, vbCrLf & "\deflang")
  If lBOS = 0 Then GoTo finally Else lBOS = lBOS + 2
  lEOS = InStr(lBOS, strRTF, vbCrLf & "\par")
  If lEOS = 0 Then GoTo finally
  While Not gHellFrozenOver
    strTmp = Mid(strRTF, lBOS, lEOS - lBOS)
    l = lBOS
    While l <= lEOS
      strTmp = Mid(strRTF, l, 1)
      Select Case strTmp
      Case "{"
        l = l + 1
      Case "}"
        strHTML = strHTML & strEOS
        l = l + 1
      Case "\"  'special code
        l = l + 1
        strTmp = Mid(strRTF, l, 1)
        Select Case strTmp
        Case "b"
          If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
            strHTML = strHTML & "<B>"
            strEOS = "</B>" & strEOS
            If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
          ElseIf (Mid(strRTF, l, 7) = "bullet ") Then
            strHTML = strHTML & "•"  'bullet
            l = l + 6
          Else
            gSkip = True
          End If
        Case "e"
          If (Mid(strRTF, l, 7) = "emdash ") Then
            strHTML = strHTML & "—"
            l = l + 6
          Else
            gSkip = True
          End If
        Case "i"
          If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
            strHTML = strHTML & "<I>"
            strEOS = "</I>" & strEOS
            If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
          Else
            gSkip = True
          End If
        Case "l"
          If (Mid(strRTF, l, 10) = "ldblquote ") Then
            strHTML = strHTML & "“"
            l = l + 9
          ElseIf (Mid(strRTF, l, 7) = "lquote ") Then
            strHTML = strHTML & "‘"
            l = l + 6
          Else
            gSkip = True
          End If
        Case "p"
          If ((Mid(strRTF, l, 6) = "plain\") Or (Mid(strRTF, l, 6) = "plain ")) Then
            strHTML = strHTML & strEOS
            strEOS = ""
            If Mid(strRTF, l + 5, 1) = "\" Then l = l + 4 Else l = l + 5  'catch next \ but skip a space
          Else
            gSkip = True
          End If
        Case "r"
          If (Mid(strRTF, l, 7) = "rquote ") Then
            strHTML = strHTML & "’"
            l = l + 6
          ElseIf (Mid(strRTF, l, 10) = "rdblquote ") Then
            strHTML = strHTML & "”"
            l = l + 9
          Else
            gSkip = True
          End If
        Case "t"
          If (Mid(strRTF, l, 4) = "tab ") Then
            strHTML = strHTML & Chr$(9)  'tab
            l = l + 3
          Else
            gSkip = True
          End If
        Case "'"
          strTmp2 = "{" & Mid(strRTF, l + 1, 2) & "}"
          lTmp = InStr(strCodes, strTmp2)
          If lTmp = 0 Then
            strHTML = strHTML & Chr("&H" & Mid(strTmp2, 2, 2))
          Else
            strHTML = strHTML & Trim(Mid(strCodes, lTmp - 8, 8))
          End If
          l = l + 2
        Case "~"
          strHTML = strHTML & " "
        Case "{", "}", "\"
          strHTML = strHTML & strTmp
        Case vbLf, vbCr, vbCrLf  'always use vbCrLf
          strHTML = strHTML & vbCrLf
        Case Else
          gSkip = True
        End Select
        If gSkip = True Then
          'skip everything up until the next space or "\"
          While ((Mid(strRTF, l, 1) <> " ") And (Mid(strRTF, l, 1) <> "\"))
            l = l + 1
          Wend
          gSkip = False
          If (Mid(strRTF, l, 1) = "\") Then l = l - 1
        End If
        l = l + 1
      Case vbLf, vbCr, vbCrLf
        l = l + 1
      Case Else
        strHTML = strHTML & strTmp
        l = l + 1
      End Select
    Wend
    lBOS = lEOS + 2
    lEOS = InStr(lEOS + 1, strRTF, vbCrLf & "\par")
    If lEOS = 0 Then GoTo finally
    strHTML = strHTML & "<br>"
  Wend
finally:
  RTF2HTML = strHTML
End Function
```

