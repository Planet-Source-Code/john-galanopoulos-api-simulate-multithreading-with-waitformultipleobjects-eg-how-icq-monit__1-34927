<div align="center">

## API \- Simulate multithreading with WaitForMultipleObjects \(eg\. How ICQ monitors connection state\)


</div>

### Description

In this article, we are going to see how to use WaitForSingleObject, WaitForMultipleObjects, RasConnectionNotification and many other commands with Visual Basic. We are also going to see how to monitor multiple events without the need of multithreading. I have included two examples : how to monitor when a shelled application has ended and how ICQ monitors connection state (that little flower that gets green when we dial-up and establish a connection). If you like it, post a comment. I ll be happy to read your thoughts or suggestions. (

----

A Special "Thank you" goes to all of you who spent a few secs to rate this article :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-07-05 20:13:46
**By**             |[John Galanopoulos](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-galanopoulos.md)
**Level**          |Intermediate
**User Rating**    |5.0 (194 globes from 39 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[API\_\-\_Simu102436752002\.zip](https://github.com/Planet-Source-Code/john-galanopoulos-api-simulate-multithreading-with-waitformultipleobjects-eg-how-icq-monit__1-34927/archive/master.zip)





### Source Code

<p><font face="Monotype Corsiva" size="4" color="#800000"><b>Simulate
multithreading with WaitForMultipleObjects </b></font></p>
<p><font color="#800000" face="Monotype Corsiva"><b>              
</b></font><font face="Monotype Corsiva"><b><font color="#800000">(eg. How ICQ
monitors connection state)</font></b></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">I
have used extensivly the event driven mechanism that Windows provide in many
different programming aspects</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">(RDO,
ADO, ODBC, Windows Sockets, Winlogon, mutexes, semaphores etc) and used
WaitForSingleObject when</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">i
was in need of an event monitor API command. </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">The
WaitForSingleObject is located in kernel32.dll and waits until a specific event
objects gets signaled or when a time limit is</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">reached.
It accepts two parameters; a handle to the event object and a time-out interval. </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><i>**
</i><i>The main benefit of this function is that it uses no processor time while waiting for the object state </i></font></p> <p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><i> </i><i>to become signaled or the time-out interval to elapse.</i></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"> </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">Here
is the declaration :</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">Public Declare Function
</font> WaitForSingleObject<font color="#0000FF"> Lib </font> "kernel32" <font color="#0000FF"> Alias </font> "WaitForSingleObject"<font color="#0000FF"> _</font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">(ByVal
</font> hHandle<font color="#0000FF"> As Long, ByVal </font> dwMilliseconds<font color="#0000FF"> As Long) As Long</font><br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">Let's
see an example of this command's usage :</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">In
this example we are going to run the Windows calculator.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">We
will open this shelled process and we will monitor the process handle; </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">if
it gets 0 then the process was ended.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">Public Const</font> WAIT_FAILED = &HFFFFFFFF              
<font color="#008000">'Our WaitForSingleObject failed to wait and returned -1</font><br>
<font color="#0000FF">Public Const </font> WAIT_OBJECT_0 = &H0&         <font color="#008000">              
'The waitable object got signaled</font><br>
<font color="#0000FF">Public Const</font> WAIT_ABANDONED = &H80&               
<font color="#008000">'We got out of the waitable object</font><br>
<font color="#0000FF">Public Const</font> WAIT_TIMEOUT = &H102&                     
<font color="#008000">'the interval we used, timed out.<br>
</font><font color="#0000FF">Public Const</font> STANDARD_RIGHTS_ALL = &H1F0000 <font color="#008000">
'No special user rights needed to open this process</font><br>
<br>
<font color="#0000FF">Public Declare Function</font> OpenProcess <font color="#0000FF"> Lib</font> "kernel32" (<font color="#0000FF">ByVal</font> dwDesiredAccess
<font color="#0000FF"> As Long</font>, <font color="#0000FF"> ByVal</font> bInheritHandle
<font color="#0000FF"> As Long</font>, <font color="#0000FF"> ByVal</font> dwProcessId
<font color="#0000FF"> As Long</font>)<font color="#0000FF"> As Long</font><br>
<font color="#0000FF">Public Declare Function</font> WaitForSingleObject <font color="#0000FF">Lib</font>
"kernel32" (<font color="#0000FF">ByVal</font> hHandle <font color="#0000FF">As
Long</font>, <font color="#0000FF">ByVal</font> dwMilliseconds <font color="#0000FF">As
Long</font>) <font color="#0000FF">As Long</font><br>
<font color="#0000FF">Public Declare Function</font> CloseHandle <font color="#0000FF"> Lib</font> "kernel32" (<font color="#0000FF">ByVal</font> hObject
<font color="#0000FF"> As Long</font>) <font color="#0000FF"> As Long</font><br>
<br>
<font color="#0000FF">Public Sub </font> ShelledAPP()<font color="#0000FF"><br>
Dim </font> shProcID<font color="#0000FF"> As Long<br>
Dim </font> hProcess<font color="#0000FF"> As Long<br>
Dim </font> WaitRet<font color="#0000FF"> As Long<br>
<br>
</font>shProcID =<font color="#0000FF"> </font>Shell("calc.exe", vbNormalFocus)<font color="#0000FF"><br>
</font>hProcess = OpenProcess(STANDARD_RIGHTS_ALL, <font color="#0000FF">False</font>, shProcID)<font color="#0000FF"><br>
</font><font color="#008000"><br>
'This is the proper and optimized way to use the WaitForSingleObject
function. </font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#008000">'</font><font color="#008000">I
saw many programmers use the INFINITE constant as </font><font color="#008000">for
the dwMilliseconds field. </font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font color="#008000" size="2" face="Tahoma">'If
<i>dwMilliseconds</i> is INFINITE, the function's time-out interval never
elapses.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font color="#008000" size="2" face="Tahoma">'That's
wrong 'cause the program won't refresh thus giving the impression that is a hung
application.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font color="#008000" size="2" face="Tahoma">'In
Windows XP specially you might see a popup screen informing you about this.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#008000">'The
problem also appears when you apply WaitForSingleObject with </font><font color="#008000">INFINITE
</font><font color="#008000">in an application that</font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font color="#008000" size="2" face="Tahoma">'uses
windows. </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font color="#008000" size="2" face="Tahoma">'Always
use a reasonable number of milliseconds and always use DoEvents to refresh the
program's message queue</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF"> Do  <br>
   </font>WaitRet = WaitForSingleObject(hProcess, 10)   <font color="#008000">'
wait for 10ms to see if the hProcess was signaled</font><font color="#0000FF"><br>
           Select Case </font> WaitRet<font color="#0000FF"><br>
  </font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF"> </font><font color="#0000FF">                  
Case </font> WAIT_TIMEOUT   <font color="#008000">'The first case must
always be WAIT_TIMEOUT 'cause it is the most used option</font><font color="#0000FF"><br>
                             
</font>DoEvents               
<font color="#008000">'until the shelled process terminates </font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">                    </font><font color="#0000FF"><br>
                   
Case </font> WAIT_FAILED or WAIT_ABANDONED<font color="#0000FF"><br>
                             
</font>MsgBox "Wait failed or abandoned"<font color="#0000FF"><br>
                             
Exit Do<br>
<br>
                   
Case </font> WAIT_OBJECT_0 <font color="#008000">'The object got signaled so
inform user and get out of the loop</font><font color="#0000FF"><br>
                            
</font>MsgBox "The shelled application has ended"<font color="#0000FF"><br>
                            
Exit Do<br>
<br>
          End Select<br>
 Loop<br>
</font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">Call
</font>CloseHandle(hProcess)     <font color="#008000">'Close
the process handle</font><font color="#0000FF"><br>
Call </font>CloseHandle(shProcID)    <font color="#008000">'Close
the process id handle</font><br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"> DoEvents      
<font color="#008000">'free any pending messages from the message queue</font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF"><br>
End Sub<br>
<br>
</font>Now what if we had to monitor two or more shelled applications? are we
going to use multithreading?</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">I
haven't yet implemented multithreading api in a vb.net project of mine but as
you most know, </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">multithreading
is lethal (basically for those who will implement the CreateThread API function) when used within Visual Basic 6 (or prior).</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">Crashes,
unexpected terminations, exceptions</font> <font face="Tahoma" size="2">and many
other "beautifull" encounters are some of the experiences a programmer
can get.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">The
answer comes from WaitForMultipleObjects API function which is also included in
kernel32.dll</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">Here
is the declaration :</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">Public Declare Function</font> WaitForMultipleObjects
<font color="#0000FF"> Lib</font> "kernel32" <font color="#0000FF"> Alias</font> "WaitForMultipleObjects"
(<font color="#0000FF">ByVal</font> nCount <font color="#0000FF"> As Long</font>, lpHandles
<font color="#0000FF"> As Long</font>, <font color="#0000FF"> ByVal</font> bWaitAll
<font color="#0000FF"> As Long</font>, ByVal dwMilliseconds <font color="#0000FF"> As
Long</font>) <font color="#0000FF"> As Long</font><br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">it
accepts four values :</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><b>nCount</b>
as the maximum number of events to monitor,</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><b>lpHandles</b>
as the array of different event handles (not multiple copies of the same one),</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><b>bWaitAll</b> 
(True/False) True if it must return when the state of all objects is signaled,</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">                                  
False if it must return when the state of any one of these objects gets
signaled,</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><b>dwMilliseconds</b>
as a maximum time-out interval</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">Like
WaitForSingleObject, WaitForMultipleObjects can accept event handles of any of
the following object types </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">in
the <i>lpHandles</i> : Change notification, Console input, Waitable timmer,
Event, Job, Mutex, Process, Semaphore</font></p>
<font face="Tahoma" size="2">and Threads</font>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">In
the following example we are going to try something else than monitoring
multiple shelled apps; </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">Those
of you that have ICQ installed, have noticed that "red flower" icon,
placed on the system tray.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">When
you are not connected on the internet, ICQ makes this icon look like inactive.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">Now
when you connect, it suddently starts to get one by one of it's leaf green,
meaning that it tries to</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">connect
to it's main server and when the connection completes, the flower get's green.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">How
do they do it? I mean. do they have an IsConnected() function on a timer with
some interval?</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">Definetly
no!</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">What
they do is take advantage of WaitForMultipleObjects with another function
located in rasapi32.dll; RasConnectionNotification </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">The
RasConnectionNotification function specifies an event object that the system
sets to the signaled state when </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">a
RAS connection is created or terminated.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">The
function accepts three values :</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><b>hrasconn</b>
as the handle to a RAS connection </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><b>hEvent</b> 
as the  handle to an event object </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><b>dwFlags</b>
as the type of event to receive notifications for (RASCN_Connection or RASCN_Disconnection)</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">Now
we are going to use WaitForMultipleObjects  to monitor both events</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">Public Const</font> RASCN_Connection = &H1      <font color="#008000">
'Our two flags</font><br>
<font color="#0000FF">Public Const</font> RASCN_Disconnection = &H2<br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">Public Const</font> WAIT_FAILED = &HFFFFFFFF<br>
<font color="#0000FF">Public Const</font> WAIT_OBJECT_0 = &H0&<br>
<font color="#0000FF">Public Const</font> WAIT_ABANDONED = &H80&<br>
<font color="#0000FF">Public Const</font> WAIT_TIMEOUT = &H102&<br>
<br>
<font color="#0000FF">Public Type</font> SECURITY_ATTRIBUTES<br>
          nLength <font color="#0000FF"> As Long</font><br>
          lpSecurityDescriptor <font color="#0000FF"> As Long</font><br>
          bInheritHandle <font color="#0000FF"> As Long</font><br>
<font color="#0000FF">End Type</font><br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">Public Declare Function
</font> CreateEvent <font color="#0000FF"> Lib</font> "kernel32" <font color="#0000FF"> Alias</font>
"CreateEventA" (lpEventAttributes <font color="#0000FF"> As</font>
SECURITY_ATTRIBUTES, <font color="#0000FF"> ByVal</font> bManualReset <font color="#0000FF"> As
Long</font>, <font color="#0000FF"> ByVal</font> bInitialState <font color="#0000FF"> As
Long</font>, <font color="#0000FF"> ByVal</font> lpName <font color="#0000FF"> As
String</font>) <font color="#0000FF"> As Long</font><br>
<font color="#0000FF">Public Declare Function</font> RasConnectionNotification <font color="#0000FF"> Lib</font> "rasapi32.dll"
<font color="#0000FF"> Alias</font> "RasConnectionNotificationA" (hRasConn <font color="#0000FF"> As
Long</font>, <font color="#0000FF"> ByVal</font> hEvent <font color="#0000FF"> As
Long</font>, <font color="#0000FF"> ByVal</font> dwFlags <font color="#0000FF"> As
Long</font>) <font color="#0000FF"> As Long</font><br>
<font color="#0000FF">Public Declare Function</font> WaitForMultipleObjects <font color="#0000FF"> Lib</font> "kernel32" (<font color="#0000FF">ByVal</font> nCount
<font color="#0000FF"> As Long</font>, lpHandles <font color="#0000FF"> As Long</font>,
<font color="#0000FF"> ByVal</font> bWaitAll <font color="#0000FF"> As Long</font>,
<font color="#0000FF"> ByVal</font> dwMilliseconds <font color="#0000FF"> As
Long</font>) <font color="#0000FF"> As Long</font><br>
<font color="#0000FF">Public Declare Function</font> ResetEvent <font color="#0000FF"> Lib</font> "kernel32"
(<font color="#0000FF">ByVal</font> hEvent <font color="#0000FF"> As Long</font>)
<font color="#0000FF"> As Long</font><br>
<font color="#0000FF">Public Declare Function</font> CloseHandle <font color="#0000FF"> Lib</font> "kernel32" (<font color="#0000FF">ByVal</font> hObject
<font color="#0000FF"> As Long</font>) <font color="#0000FF"> As Long</font><br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">Public Sub</font> MonitorRASStatusAsync()<br>
<br>
<font color="#0000FF">Dim</font> hEvents(1) <font color="#0000FF"> As Long       </font><font color="#008000">
'Array of event handles. Since there are two events we'd like to monitor, i have
already dimention it.</font><br>
<font color="#0000FF">Dim</font> RasNotif <font color="#0000FF"> As Long          </font><br>
<font color="#0000FF">Dim</font> WaitRet <font color="#0000FF"> As Long           </font><br>
<font color="#0000FF">Dim</font> sd <font color="#0000FF"> As</font> SECURITY_ATTRIBUTES<br>
<font color="#0000FF">Dim</font> hRasConn <font color="#0000FF"> As Long</font><br>
<br>
hRasConn = 0<br>
<br>
<font color="#008000">'We are going to create and register two event objects
with CreateEvent API function</font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'There
aren't any special treated events that need any kind of security attributes so
we just initialize the structure</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">With</font> sd                                <br>
       .nLength = Len(sd)     <font color="#008000">
'we pass the length of sd </font><br>
       .lpSecurityDescriptor = 0<br>
       .bInheritHandle = 0<br>
<font color="#0000FF">End With</font><br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'We
create the event by passing in CreateEvent any security attributes, </font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'we
want to manually reset the event after it gets signaled,</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'we
also want it's initial state not signaled assuming that we don't have yet any
connection to the internet,</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#008000">'last
but not least we give the event a name (RASStatusNotificationObject1)<br>
</font>hEvents(0) = CreateEvent(sd, <font color="#0000FF">True</font>, <font color="#0000FF">False</font>, "RASStatusNotificationObject1")<br>
<font color="#008000">'If the returned value was zero, something went wrong so
exit the sub</font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">If</font> hEvents(0) = 0
<font color="#0000FF"> Then</font> MsgBox "Couldn't assign an event handle": <font color="#0000FF"> Exit Sub</font><br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'If
we succesfully created the first event object we pass it to
RasConnectionNotification</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'with
the flag RASCN_Connection so that this event will monitor for internet
connection</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">RasNotif = RasConnectionNotification(ByVal hRasConn, hEvents(0), RASCN_Connection)<br>
<font color="#0000FF">If</font> RasNotif <> 0 <font color="#0000FF"> Then</font> MsgBox "Ras Notification failure":
<font color="#0000FF"> GoTo</font> ras_TerminateEvent<br>
<br>
<br>
<font color="#008000">'We create the second event object exactly like the first
one</font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font color="#008000" face="Tahoma" size="2">'but
we name it </font><font face="Tahoma" size="2" color="#008000">RASStatusNotificationObject2</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">hEvents(1) = CreateEvent(sd,
<font color="#0000FF">True</font>, <font color="#0000FF">False</font>, "RASStatusNotificationObject2")<br>
<font color="#0000FF">If</font> hEvents(1) = 0 <font color="#0000FF"> Then</font> MsgBox "Couldn't assign
an event handle": <font color="#0000FF"> Exit Sub</font><br>
<br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'If
we succesfully created the second event object too, we pass it to
RasConnectionNotification</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'with
the flag RASCN_Disconnection. This event will monitor for disconnection</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">RasNotif = RasConnectionNotification(ByVal hRasConn, hEvents(1), RASCN_Disconnection)<br>
<font color="#0000FF">If </font> RasNotif <> 0 <font color="#0000FF"> Then</font> MsgBox "Ras Notification failure":
<font color="#0000FF"> GoTo</font> ras_TerminateEvent<br>
<br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'We
then issue the loop</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'Notice
that we have put hEvents array to it's first array item.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'and
we used False cause we want to get notifications</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#008000">'when
any of the two events occur.<br>
</font> <font color="#0000FF">Do</font><br>
       WaitRet = WaitForMultipleObjects(2, hEvents(0),
<font color="#0000FF">False</font>, 20)<br>
                      
<font color="#0000FF">Select Case</font> WaitRet<br>
                                
<font color="#0000FF">Case</font> WAIT_TIMEOUT<br>
                                         
DoEvents<br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">                                
<font color="#0000FF">Case</font> WAIT_FAILED <font color="#0000FF">Or</font> WAIT_ABANDONED
<font color="#0000FF">Or</font> WAIT_ABANDONED  + 1<br>
                                         
GoTo ras_TerminateEvent<br>
<br>
                                
<font color="#0000FF">Case </font> WAIT_OBJECT_0<br>
                                          
MsgBox "Connected"<br>
                                          
ResetEvent hEvents(0)<font color="#008000"> 'Reset the event to avoid a second
message box </font><br>
                                          
DoEvents    <font color="#008000">'Free any pending messages</font><br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">                                
<font color="#0000FF">Case</font> WAIT_OBJECT_0 + 1 <br>
                                         
MsgBox "Disconnected"<br>
                                         
ResetEvent hEvents(1) <font color="#008000">'Reset the event to place it in no
signal state (Manual reset, remember?)</font><br>
                                         
DoEvents <br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">                       
<font color="#0000FF">End Select</font><br>
                  <br>
 <font color="#0000FF"> Loop<br>
</font><br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">ras_TerminateEvent:<br>
<br>
<font color="#008000">'Close all event handles</font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2" color="#008000">'For
more than two events you could apply  a For.. Next</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"><font color="#0000FF">Call</font>
CloseHandle(hEvents(1))     <br>
<font color="#0000FF">Call</font> CloseHandle(hEvents(0))<br>
</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2"> DoEvents   </font><font face="Tahoma" size="2">
<font color="#008000">'Free any pending messages from the application message
queue</font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma"><font size="2"><br>
<font color="#0000FF">End Sub</font></font><br>
<br>
<font size="2">Now imagine that you could monitor events from different objects
like </font></font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">a
file or folder change, along with connection status, shelled applications,</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">multiple
printer objects, different processes and threads etc etc etc.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">(64
maximum event objects i think)</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">It
will appear that you program is multithreading but the truth behind that, is</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">that
you will be taking advantage of WaitForMultipleObjects internal</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">multithreading
mechanism.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">I
hope i helped with this article, people.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">Feel
free to leave any comments or suggestions.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><font face="Tahoma" size="2">It
will help all of us.</font></p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"> </p>
<p style="text-indent: 0; word-spacing: 0; line-height: 100%; margin: 0" align="left"><b><font face="Monotype Corsiva" color="#0000FF" size="4">John
Galanopoulos</font></b></p>
<br><b>Need Oracle tips? try here : http://aboutoracle.blogspot.com<b><br>

