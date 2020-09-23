Hi all,

Thanks for trying Ventura Dictionary. Please take a moment to look at some installations facts.

General Dependency:
*******************
1)Your machine should have Microsoft Word 97 or above for this application to work.
2) Please registerthe vbsendmail.dll library (using regsvr32), so that the email feature works, and I can
get feedbacks.


For Windows 95/98/NT
********************

This application uses a TTS engine provided by Microsoft. The DLL is called XVoice.dll. The application
will run fine without that, but the TTS functionality will be unavailable. Just do this:

1) Copy the XVoice.dll file from the Dependency Folder and copy it to Windows\System directory or Windows\System32 directory, and
register it with regsvr32 exe.

For Windows 2000/XP/Me
**********************
The XVoice.dll should be available by default in your system. However, if it is not, look at the instructions for Windows 95/98.

Please email all bug reports, suggestions, to soumyas_v@hotmail.com.
Feel free to yell at me too...

Soumya Sengupta.
