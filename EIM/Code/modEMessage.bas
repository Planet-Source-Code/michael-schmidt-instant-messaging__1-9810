Attribute VB_Name = "modEMessage"
Option Explicit
'====================================
'   API
'====================================
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Const EIMPort = 4564 ' Port our program uses.

'====================================
'   Get User Name
'====================================
Function GetUser()
Dim usr As String
Dim aa As String

    usr = Space(256)
    aa = GetUserName(usr, 256)
    GetUser = Left(RTrim(usr), Len(RTrim(usr)) - 1)

End Function


'====================================
'   Normal Message
'====================================
Public Sub SendEMessage()
Dim SIM As New frmEMessage ' Send Instant MessageLoad SIM
Dim pPacket    As String ' Packet To Send
Dim MaxAddr    As String ' Max IP In LAN Scan
Dim MinAddr    As String ' Min IP In LAN Scan
Dim LocAddr    As String ' Our Address
Dim LocPort    As String ' Our Port
Dim RemoteAddress As String

    ' Load Form
    SIM.txtPORT.Text = EIMPort
    SIM.txtSUBJECT.Text = "(No Subject)"
    SIM.Show

    ' Winsock Data (Bind to random port)
    SIM.sckListen.Bind SIM.sckListen.LocalPort
    SIM.sckSM.Bind SIM.sckSM.LocalPort
    
    ' Parse Data Add Min & Max To End
    RemoteAddress = SIM.sckListen.LocalIP
    MinAddr = Word(RemoteAddress, 1, ".") & "." & _
              Word(RemoteAddress, 2, ".") & "." & _
              Word(RemoteAddress, 3, ".") & ".1"
    MaxAddr = Word(RemoteAddress, 1, ".") & "." & _
              Word(RemoteAddress, 2, ".") & "." & _
              Word(RemoteAddress, 3, ".") & ".254"
    
    ' Build Hello Packet
    LocPort = SIM.sckListen.LocalPort
    LocAddr = SIM.sckListen.LocalIP
    pPacket = "030" & Chr(1) & LocAddr & Chr(1) & LocPort & Chr(1)

    ' Send Packet
    PingLAN SIM.sckSM, SIM.txtPORT.Text, MinAddr, MaxAddr, pPacket

End Sub


'====================================
'   Server
'====================================
Public Sub StartEServer()
    Load frmEServer
    frmEServer.sckSYS.Bind EIMPort ' Port To Listen
End Sub


'====================================
'   UDP Packet
'====================================
Public Sub UDPPACKET(WINSCK As Object, ByVal Port As Integer, ByVal IP As String, ByVal Packet As String)
On Error GoTo ErrHandle
    
    ' Send Packet
    
    WINSCK.RemotePort = Port
    WINSCK.RemoteHost = IP
    DoEvents
    WINSCK.SendData (Packet)
    
Exit Sub

ErrHandle:
If Err.Number = 10014 Then MsgBox "Invalid Address", vbCritical, App.Title _
Else MsgBox Err.Number & " : " & Err.Description, vbCritical, "Error"
Err.Clear

End Sub


'====================================
'   Ping Lan
'====================================
Public Sub PingLAN(WINSCK As Object, Port As Integer, MinIP, MaxIP, Packet As String)
Dim LANLoop As Integer

Dim Amax, Bmax, Cmax, Dmax As Integer
Dim Amin, Bmin, Cmin, Dmin As Integer
Dim a, b, C, D As Integer

' PARSE MAX IP
Amax = Left(MaxIP, InStr(1, MaxIP, ".") - 1)
Bmax = Mid(MaxIP, Len(Amax) + 2, InStr(Len(Amax) + 2, MaxIP, ".") - (Len(Amax) + 2))
Cmax = Mid(MaxIP, Len(Amax) + Len(Bmax) + 3, InStr(Len(Amax) + Len(Bmax) + 4, MaxIP, ".") - (Len(Amax) + Len(Bmax) + 3))
Dmax = Mid(MaxIP, Len(Amax) + Len(Bmax) + Len(Cmax) + 4, Len(MaxIP) - Len(Amax) + Len(Bmax) + Len(Cmax) + 3)

' PARSE MIN IP
Amin = Left(MinIP, InStr(1, MinIP, ".") - 1)
Bmin = Mid(MinIP, Len(Amin) + 2, InStr(Len(Amin) + 2, MinIP, ".") - (Len(Amin) + 2))
Cmin = Mid(MinIP, Len(Amin) + Len(Bmin) + 3, InStr(Len(Amin) + Len(Bmin) + 4, MinIP, ".") - (Len(Amin) + Len(Bmin) + 3))
Dmin = Mid(MinIP, Len(Amin) + Len(Bmin) + Len(Cmin) + 4, Len(MinIP) - Len(Amin) + Len(Bmin) + Len(Cmin) + 3)

On Error Resume Next

For a = Amin To Amax
 For b = Bmin To Bmax
  For C = Cmin To Cmax
   For D = Dmin To Dmax
     WINSCK.RemoteHost = a & "." & b & "." & C & "." & D
     WINSCK.RemotePort = Port
     DoEvents
     WINSCK.SendData (Packet)
     If Err.Number <> 0 Then
      Err.Clear
     End If
   Next D
  Next C
 Next b
Next a


End Sub


'===============================================
' NOTE: I Added a field, this determines the
' separator for words, which is usually a space
' but may now be changed to whatever you wish.
' I separate words by Chr(1) so when i call
' a procedure, i use words("find 3 in here",Chr(1))
' which uses Chr(1) as a separator.
'===============================================
'Words.bas - string handling functions for words
'Author: Evan Sims         [esims@arcola-il.com]
'Based on a module by Kevin O'Brien
'Version - 1.2 (Sept. 1996 - Dec 1999)
'
'These functions deal with "words".
'Words = blank-delimited strings
'Blank = any combination of one or more spaces,
'        tabs, line feeds, or carriage returns.
'
'Examples:
'      word("find 3 in here", 3)     = "in"      3rd word
' Modified to find given character (chr(1)) instead of hard coded spaces
' separating words...
'     words("find 3 in here")        = 4         number of words
'     split("here's /s more", "/s")  = "more"    Returns words after split identifier (/s)
'   delWord("find 3 in here", 1, 2)  = "in here" delete 2 words, start at 1
'   midWord("find 3 in here", 1, 2)  = "find 3"  return 2 words, start at 1
'   wordPos("find 3 in here", "in")  = 3         word-number of "in"
' wordCount("find 3 in here", "in")  = 1         occurrences of word "in"
' wordIndex("find 3 in here", "in")  = 8         position of "in"
' wordIndex("find 3 in here", 3)     = 8         position of 3rd word
' wordIndex("find 3 in here", "3")   = 6         position of "3"
'wordLength("find 3 in here", 3)     = 2         length of 3rd word
'
'Difference between Instr() and wordIndex():
'     InStr("find 3 in here", "in")   = 2
' wordIndex("find 3 in here", "in")   = 8
'
'     InStr("find 3 in here", "her")  = 11
' wordIndex("find 3 in here", "her")  = 0
'===============================================

Public Function Word(ByVal sSource As String, n As Long, SP As String) As String
'=================================================
' Word retrieves the nth word from sSource
' Usage:
'    Word("red blue green ", 2)   "blue"
'=================================================
Dim pointer As Long   'start parameter of Instr()
Dim pos     As Long   'position of target in InStr()
Dim X       As Long   'word count
Dim lEnd    As Long   'position of trailing word delimiter

sSource = CSpace(sSource)

'find the nth word
X = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If X = n Then                               'the target word-number
      lEnd = InStr(pointer, sSource, SP)       'pos of space at end of word
      If lEnd = 0 Then lEnd = Len(sSource) + 1 '   or if its the last word
      Word = Mid$(sSource, pointer, lEnd - pointer)
      Exit Do                                  'word found, done
   End If
  
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   X = X + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop
  
End Function

Public Function DelWord(ByVal sSource As String, _
                                    n As Long, SP As String, _
                                    Optional vWords As Variant) As String
'===========================================================
' DelWord deletes from sSource, starting with the
' nth word for a length of vWords words.
' If vWords is omitted, all words from the nth word on are
' deleted.
' Usage:
'   DelWord("now is not the time", 3)     "now is"
'   DelWord("now is not the time", 3, 1)  "now is the time"
'===========================================================

Dim lWords  As Long    'length of sTarget
Dim lSource As Long    'length of sSource
Dim pointer As Long    'start parameter of InStr()
Dim pos     As Long    'position of target in InStr()
Dim X       As Long    'word counter
Dim lStart  As Long    'position of word n
Dim lEnd    As Long    'position of space after last word

lSource = Len(sSource)
DelWord = sSource
sSource = CSpace(sSource)
If IsMissing(vWords) Then
   lWords = -1
ElseIf IsNumeric(vWords) Then
   lWords = CLng(vWords)
Else
   Exit Function
End If

If n = 0 Or lWords = 0 Then Exit Function      'nothing to delete

'find position of n
X = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If X = n Then                               'the target word-number
      lStart = pointer
      If lWords < 0 Then Exit Do
   End If
   
   If lWords > 0 Then                          'lWords was provided
      If X = n + lWords - 1 Then               'find pos of last word
         lEnd = InStr(pointer, sSource, SP)    'pos of space at end of word
         Exit Do                               'word found, done
      End If
   End If
   
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   X = X + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop
If lStart = 0 Then Exit Function
If lEnd = 0 Then
   DelWord = Trim$(Left$(sSource, lStart - 1))
Else
   DelWord = Trim$(Left$(sSource, lStart - 1) & Mid$(sSource, lEnd + 1))
End If
End Function

Public Function CSpace(sSource As String) As String
'==================================================
'CSpace converts blank characters
'(ascii: 9,10,13,160) to space (32)
'
'  cSpace("a" & vbTab   & "b")  "a b"
'  cSpace("a" & vbCrlf  & "b")  "a  b"
'==================================================
Dim pointer   As Long
Dim pos       As Long
Dim X         As Long
Dim iSpace(3) As Integer

' define blank characters
iSpace(0) = 9    'Horizontal Tab
iSpace(1) = 10   'Line Feed
iSpace(2) = 13   'Carriage Return
iSpace(3) = 160  'Hard Space

CSpace = sSource
For X = 0 To UBound(iSpace) ' replace all blank characters with space
   pointer = 1
   Do
      pos = InStr(pointer, CSpace, Chr$(iSpace(X)))
      If pos = 0 Then Exit Do
      Mid$(CSpace, pos, 1) = " "
      pointer = pos + 1
   Loop
Next X

End Function
