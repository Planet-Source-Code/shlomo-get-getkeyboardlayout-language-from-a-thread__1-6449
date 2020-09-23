<div align="center">

## get GetKeyboardLayout language from a thread


</div>

### Description

hi!this is my first submit

finely i think i can put something usefull

for other users.

this code read the keyboard language from

another application all you need is to send

the handle of the thread window.
 
### More Info
 
winHWND is the thread window Handle

i think NON i tested it on my mechine


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shlomo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shlomo.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shlomo-get-getkeyboardlayout-language-from-a-thread__1-6449/archive/master.zip)

### API Declarations

```
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
```


### Source Code

```
'find out what keyboard language a theard is
Public Sub FindTheardlanguage ()
Dim TheardId As Long
Dim TheardLang As Long
  TheardId = get_threadId 'call function
  TheardLang = GetKeyboardLayout(ByVal TheardId)
  TheardLang = TheardLang Mod 10000
 Select Case TheardLang
  Case 9721 'english
  'do your stuff
  Case 1869 'hebrew
   'do your stuff
 End Select
End Sub
Public Function get_threadId() As Long
Dim threadid As Long, processid As Long
get_threadId = GetWindowThreadProcessId(winHWND, processid)
End Function
```

