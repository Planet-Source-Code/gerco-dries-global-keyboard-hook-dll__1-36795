<div align="center">

## Global Keyboard Hook Dll


</div>

### Description

This is a dll to intercept all keyboard messages to any program in Windows (except for the ALT-TAB, CTRL-ALT-DEL and CTRL-ESC keycombo's). It uses a C++ dll found right here on PSC, I just wrapped that one into a friendly ActiveX DLL for ease-of-use.
 
### More Info
 
You need to compile the C++ DLL first (With MSVC++ 6), since PSC doesn't allow binary uploads. Then compile the GlobalHook.dll (with VB6). Then you can open and run the sample program.

Be aware that you should NOT try to start 2 hooks at the same time. The code should block it, but just in case it doesn't, I don't know what will happen. You have been warned.


<span>             |<span>
---                |---
**Submitted On**   |2002-07-11 11:45:46
**By**             |[Gerco Dries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gerco-dries.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Global\_Key1048467112002\.zip](https://github.com/Planet-Source-Code/gerco-dries-global-keyboard-hook-dll__1-36795/archive/master.zip)








