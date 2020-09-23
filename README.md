<div align="center">

## Dial up Password recovery


</div>

### Description

The program is a dial-up password Recovery program for all windows version 98/ME/2k/NT/XP.

It displays the connection name , username and the pass of the dialup connection.

If i see interest about my code soon will post code with dialup pass on XP and phone numbers

in XP it has a bug. it displays correct all those except that the pass is 

----

. i'm working on it.Pleaze vote.If i see interest about my code soon will post code with dialup pass on XP and phone numbers.thnks
 
### More Info
 
this prog is made by john papadakhs. I am using windows API'S to access the phone

book of windows and get informations about the RAS CONNECTIONS available

the ras api is RasGetCredentials. this programm shows the connection name of eatch dialup

connection the username and the password. Only in windows XP the password is

shown with "*". I'm working on an other api that solves that problem with XP

and soon will have and XP passwords.also i'm working on getting the dialup number.

all these will be send soon if i see interest on my code.that's a promise.Pleaze vote

return the connection name, the username and the password of all dialup connection in WINDOWS version


<span>             |<span>
---                |---
**Submitted On**   |2004-09-02 12:54:02
**By**             |[John Papadakis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-papadakis.md)
**Level**          |Intermediate
**User Rating**    |4.6 (55 globes from 12 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Dial\_up\_Pa178870922004\.zip](https://github.com/Planet-Source-Code/john-papadakis-dial-up-password-recovery__1-55965/archive/master.zip)

### API Declarations

```
Private Declare Function RasGetCredentials Lib "rasapi32.dll" Alias "RasGetCredentialsA" _
 (ByVal lpcstr As String, ByVal lpcstr As String, ByRef TLPRASCREDENTIALSA As RASCREDENTIALS) _
 As Long
```





