<%
'============================================================
' MODULE:    Captcha.asp
' AUTHOR:    © www.u229.no
' CREATED:  July 2005
' HOME PAGE: http://www.u229.no/stuff/Captcha/
' LICENSE:    http://www.u229.no/stuff/license/
'============================================================
' COMMENT: This is a CAPTCHA made with Classic ASP, some CSS and some javascript.
'                  You may want to move the user checking to other parts of your own code. See
'                  info somewhere near line 150.
'                  Save this file as Captcha.asp
'============================================================
' TODO:
'    1) Include numbers?
'    2) If above = True Then: Add more human questions like doing some basic math on the numbers?
'    3) Limit the number of log in attemps based on the visitors IP Number?
'    4) The yellow color might be hard to read. It could have been gray instead.
'============================================================
' ROUTINES:

' - Function CreateCAPTCHA()
' - Sub InitArrays()
' - Sub CreateStyleSheet()
' - Sub CreateJavascript()
' - Function RandomizeArrayUnique(arr, arrNew)
' - Function RandomizeArray(arr, arrNew)
' - Function RandomNumber(iMax)
' - Function RandomString(iMax)
'============================================================

'Option Explicit
'On Error Resume Next

Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

'// YOUR PREFERENCES
Const MAX_NUMBER_OF_CHARACTERS = 3    '// How many characters in our CAPTCHA?
Const MAX_LENGTH_CSS_CLASSES = 12       '// How many characters in the CSS class names?
Const CAPTCHA_CHARACTER_FACTOR = 40  '// How many pixels are we moving each new character from left?
Const CAPTCHA_BOX_BORDER = "border: 2px solid #ccc;"   '// Style the div box holding the CAPTCHA.
Const CAPTCHA_BOX_WIDTH = 160              '// Width. This value should balance the number of characters and size.
Const CAPTCHA_BOX_HEIGHT = 80               '// Same as above.
Const NAME_OF_CAPTCHA_TEXTBOX = "txtCaptchaBox"    '// Name of CAPTCHA text box. Rename this!!


Dim m_arrCaptcha()             '// Array holding our CAPTCHA charaters. Hold in session variable.
Dim m_arrCaptchaScreen()   '// Parallell array where some items migth be hex/decimal encoded for display on screen.
Dim m_sCSS                        '// Our CSS
Dim m_sJavascript                '// Our Javascript
Dim m_sUserResult               '// Return a response to client/demo if success or failure

Dim m_sNameOfWrapperDiv  '// Holding the id name attribute for the div wrapping the CAPTCHA?
Dim m_arrColor(4)               '// Array of colors for the characters
Dim m_arrColorNew(4)         '// Same colors randomized
Dim m_arrFontFamily(4)        '// Array of font family strings
Dim m_arrFontFamilyNew(4)  '// Same fonts randomized
Dim m_arrFontSize(4)           '// Array of font sizes
Dim m_arrFontSizeNew(4)     '// Same font sizes now randomized
Dim m_arrTopPosition(4)       '// Array of top position values
Dim m_arrTopPositionNew(4) '// Same values randomized
Dim m_arrClassNames()        '// Array of names for the CSS classes
Dim m_arrQuestions(3)          '// Array of questions for the human visitor
Dim m_lngQuestionIndex       '// This number between 0 - 3 defines what question to ask the human visitor
Dim m_arrCaptchaColor(2)     '// Array holding the color of the character we are asking the visitor for
Dim m_arrCSSStrings(4)        '// Array holding our CSS elements/strings
Dim m_arrCSSStringsNew(4)  '// Same strings now randomly and uniquely sorted

'// START UP THE MODULE ARRAYS
m_arrColor(0) = "green"
m_arrColor(1) = "blue"
m_arrColor(2) = "red"
m_arrColor(3) = "black"
m_arrColor(4) = "yellow"

m_arrFontFamily(0) = "Verdana"
m_arrFontFamily(1) = "Arial"
m_arrFontFamily(2) = "Tahoma"
m_arrFontFamily(3) = "Courier"
m_arrFontFamily(4) = "Georgia"

m_arrFontSize(0) = 24
m_arrFontSize(1) = 50
m_arrFontSize(2) = 60
m_arrFontSize(3) = 40
m_arrFontSize(4) = 70

m_arrTopPosition(0) = 5
m_arrTopPosition(1) = 10
m_arrTopPosition(2) = 15
m_arrTopPosition(3) = 5
m_arrTopPosition(4) = 10

m_arrQuestions(0) = "Before submitting this form, please type the characters displayed above:"
m_arrQuestions(1) = "Before submitting this form, please type the color of the first character:"
m_arrQuestions(2) = "Before submitting this form, please type the color of the second character:"
m_arrQuestions(3) = "Before submitting this form, please type the color of the third character:"

m_arrCSSStrings(0) = "position: absolute;"
m_arrCSSStrings(1) = "top: "
m_arrCSSStrings(2) = "left: "
m_arrCSSStrings(3) = "color: "
m_arrCSSStrings(4) = "font: bold "


'------------------------------------------------------------------------------------------------------------
' Comment: Call this function from where you want to include the CAPTCHA.
'------------------------------------------------------------------------------------------------------------
Function CreateCAPTCHA()
    On Error Resume Next

    Dim i, iTmp, sTmp

'---------------------------- Create our CAPTCHA!

    '// This holds plain text characters. They are stored in a session variable and compared with the user input.
    ReDim m_arrCaptcha(MAX_NUMBER_OF_CHARACTERS - 1)
    '// This holds the decimal and hexified characters displayed on screen.
    ReDim m_arrCaptchaScreen(MAX_NUMBER_OF_CHARACTERS - 1)

    For i = 0 To (MAX_NUMBER_OF_CHARACTERS - 1)

        sTmp = UCase(RandomString(1))
        iTmp = RandomNumber(101)

        m_arrCaptcha(i) = sTmp
        
        If iTmp < 33 Then m_arrCaptchaScreen(i) = "&#" & Asc(UCase(sTmp)) & ";"              '// Decimal
        If iTmp > 66 Then m_arrCaptchaScreen(i) = "&#x" & Hex(Asc(UCase(sTmp))) & ";"    '// Hexify
        If iTmp < 67 And iTmp > 32 Then m_arrCaptchaScreen(i) = UCase(sTmp)                   '// Plain Ascii

    Next

'---------------------------- What question will we ask the human visitor?

    m_lngQuestionIndex = RandomNumber(MAX_NUMBER_OF_CHARACTERS + 1)

    '// Default max number of questions is 4
    If m_lngQuestionIndex > 4 Then m_lngQuestionIndex = RandomNumber(4)

'---------------------------- Create CSS and javascript

    Call CreateStyleSheet
    Call CreateJavascript

'---------------------------- Check to see if someone submitted CAPTCHA, machine or human
'// You may want to move this code to another part of your own application and do the testing there.

    If Len(Request.Form(NAME_OF_CAPTCHA_TEXTBOX)) > 0 Then

        If UCase(Request.Form(NAME_OF_CAPTCHA_TEXTBOX)) = UCase(Session("CAPTCHA")) Then
            m_sUserResult = "You typed " & Request.Form(NAME_OF_CAPTCHA_TEXTBOX) & " which was correct!"
        Else
            m_sUserResult = "You typed " & Request.Form(NAME_OF_CAPTCHA_TEXTBOX) & " which was wrong!" & _
                            "<br />(Support for cookies must be enabled in your web browser.)"
        End If

    End If
    
    '// Nothing was submitted, so just set a new session value which is our CAPTCHA characters or a color
    Session("CAPTCHA") = Replace(Join(m_arrCaptcha), " ", "")

    '// We will ask visitor for a color! Reduce m_lngQuestionIndex by 1 to match the m_arrCaptchaColor array
    If (m_lngQuestionIndex > 0) Then Session("CAPTCHA") = m_arrCaptchaColor(m_lngQuestionIndex - 1)

'---------------------------- Return the html

    CreateCAPTCHA = m_sCSS & m_sJavascript

End Function

'------------------------------------------------------------------------------------------------------------
' Comment: Randomize our module arrays holding the CSS values.
'------------------------------------------------------------------------------------------------------------
Sub InitArrays()
    On Error Resume Next
    
    '// First 4 arrays are randomly sorted meaning that all characters might have the same color.
    Call RandomizeArray(m_arrColor, m_arrColorNew)
    Call RandomizeArray(m_arrFontFamily, m_arrFontFamilyNew)
    Call RandomizeArray(m_arrFontSize, m_arrFontSizeNew)
    Call RandomizeArray(m_arrTopPosition, m_arrTopPositionNew)
    Call RandomizeArrayUnique(m_arrCSSStrings, m_arrCSSStringsNew)
    
End Sub

'------------------------------------------------------------------------------------------------------------
' Comment: Build the CSS.
'------------------------------------------------------------------------------------------------------------
Sub CreateStyleSheet()
    On Error Resume Next
    
    Dim sCSS, i, l, iLeft, sTmp, sTmpClassName
   
'---------------------------- Create the CSS for the div box

    '// First create a random name for the wrapper div.
    m_sNameOfWrapperDiv = RandomString(MAX_LENGTH_CSS_CLASSES)
    
    sCSS = "<style type=""text/css"">" & vbCrLf
    sCSS = sCSS & "#" & m_sNameOfWrapperDiv & " {" & vbCrLf
    sCSS = sCSS & CAPTCHA_BOX_BORDER & vbCrLf
    sCSS = sCSS & "position: relative;" & vbCrLf
    sCSS = sCSS & "width: " & CAPTCHA_BOX_WIDTH & "px;" & vbCrLf
    sCSS = sCSS & "height: " & CAPTCHA_BOX_HEIGHT & "px;" & vbCrLf
    sCSS = sCSS & "}" & vbCrLf

    '// Array holding the class names we will produce in the next loop
    ReDim m_arrClassNames(MAX_NUMBER_OF_CHARACTERS - 1)

    '// Remember the characters left position. Increase this value in every loop.
    iLeft = 10
    
'---------------------------- Build the CSS for our CAPTCHA

    For i = 0 To (MAX_NUMBER_OF_CHARACTERS - 1)

        '// Initialize our module arrays. We are randomizing them every time we get here.
        Call InitArrays
        
        sTmpClassName = RandomString(MAX_LENGTH_CSS_CLASSES)
        sCSS = sCSS & "." & sTmpClassName & " {" & vbCrLf

'---------------------------- Loop the 5 CSS strings and fill them with values
        For l = 0 To 4
            sTmp = m_arrCSSStringsNew(l)

            If InStr(sTmp, "color") > 0 Then
                sTmp = sTmp & m_arrColorNew(l) & ";" & vbCrLf

                '// We need to remember the colors when asking the visitors to type them so put them in an array.
                If Not i > UBound(m_arrCaptchaColor) Then m_arrCaptchaColor(i) = m_arrColorNew(l)
            End If

            If InStr(sTmp, "font") > 0 Then sTmp = sTmp & m_arrFontSizeNew(l) & "px " & _
                            m_arrFontFamilyNew(l) & ";" & vbCrLf
            If InStr(sTmp, "top") > 0 Then sTmp = sTmp & m_arrTopPositionNew(l) & "px;" & vbCrLf
            If InStr(sTmp, "left") > 0 Then sTmp = sTmp & iLeft & "px;" & vbCrLf
            If InStr(sTmp, "position") > 0 Then sTmp = sTmp & vbCrLf
                        
            sCSS = sCSS & sTmp
        Next
        
        '// Store the CSS class names in array
        m_arrClassNames(i) = sTmpClassName
        '// Calculate the new left position for next CAPTCHA character
        iLeft = (iLeft + CAPTCHA_CHARACTER_FACTOR)
        sCSS = sCSS & "}" & vbCrLf

    Next

    m_sCSS = sCSS & "</style>"

End Sub

'------------------------------------------------------------------------------------------------------------
' Comment: Create the javascript with our unique css class names and the CAPTCHA characters.
'------------------------------------------------------------------------------------------------------------
Sub CreateJavascript()
    On Error Resume Next

    Dim i, sJScript

    sJScript = "<script type=""text/javascript"">" & vbCrLf
    sJScript = sJScript & "document.write('<div id=""" & m_sNameOfWrapperDiv & """>');" & vbCrLf

    For i = 0 To (MAX_NUMBER_OF_CHARACTERS - 1)
        sJScript = sJScript & "document.write('<span class=""" & m_arrClassNames(i) & """>" & _
                        m_arrCaptchaScreen(i) & "</span>');" & vbCrLf
    Next

    sJScript = sJScript & "document.write('</div>');" & vbCrLf
    sJScript = sJScript & "</script>" & vbCrLf

   ' sJScript = sJScript & "<p><b>" & m_arrQuestions(m_lngQuestionIndex) & "</b></p>" & vbCrLf

    m_sJavascript = sJScript

End Sub

'------------------------------------------------------------------------------------------------------------
' Comment: Randomize array but make sure all values are present in the new array.
'------------------------------------------------------------------------------------------------------------
Function RandomizeArrayUnique(arr, arrNew)
    On Error Resume Next

    Dim i, l, sBuf, sTmp, iMax

    iMax = UBound(arr)

    ReDim arrNew(iMax)

    For i = 0 To iMax

        '// This should be enough looping
        For l = 1 To (iMax * 20)
            sTmp = arr(RandomNumber(iMax + 1))

            If InStr(sBuf, sTmp) = 0 Then
                sBuf = (sBuf & sTmp)
                arrNew(i) = sTmp
                Exit For
            End If

        Next
    Next

End Function

'------------------------------------------------------------------------------------------------------------
' Comment: Randomize our module arrays holding the CSS. One value might appear several times.
'------------------------------------------------------------------------------------------------------------
Function RandomizeArray(arr, arrNew)
    On Error Resume Next

    Dim i

    ReDim arrNew(UBound(arr))

    For i = LBound(arr) To UBound(arr)
        arrNew(i) = arr(RandomNumber(UBound(arr) + 1))
    Next

End Function

'------------------------------------------------------------------------------------------------------------
' Comment: Return a random number not bigger than the input parameter.
'------------------------------------------------------------------------------------------------------------
Function RandomNumber(iMax)
    On Error Resume Next

    Randomize
    RandomNumber = Int(iMax * Rnd)

End Function

'------------------------------------------------------------------------------------------------------------
' Comment: Create a random string of lower case letters [a-z] for the css class names.
'------------------------------------------------------------------------------------------------------------
Function RandomString(iMax)
    On Error Resume Next

    Dim i, sTmp

    For i = 1 To iMax
        sTmp = sTmp & Chr(97 + RandomNumber(26))    '// Return a random number between 97 and 122, ascii values for [a-z]
    Next

    RandomString = sTmp

End Function

'============================================================ END OF ASP CODE
%>
<!--

<html>
<head>
<title>A CAPTCHA Solution For Classic ASP</title>
</head>
<body>

<h3>A CAPTCHA Solution Built With Classic ASP, CSS And Javascript.</h3>

<form method="post" action="" name="form1">
<%Response.Write CreateCAPTCHA%>
<input type="text" name="CaptchaBox" />
<input type="submit" value="Submit" /> 
</form>

<p style="font-weight: bold; color: red;"><%=m_sUserResult%>&nbsp;</p>


</body>
</html>
-->