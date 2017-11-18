# VBS-Python
A bridge between Python and VBS - call VBS functions from a python file. Just there in case somebody needs it

This is a bridge between Python and VBScript which lets you call VBScript code from python. <br>
<b>Basic Use</b>:
```python
vbs = VBS("script.vbs") #location of the file 
vbs.run("hello()")
 ```
<br>
<b>This is the .vbs file:</b> <br>

```VB
Function hello() 
   MsgBox "Hello, Bridge World!" 
End Function 
WScript.Echo "Hello Echo"
 ```

Ideally, the file should only contain the functions as all code is executed when creating a VBS object. <br>
CScript is used, so <code>Wscript.Echo</code> is redirected to <code>stdout</code> and not displayed in message boxes<br>
Anything printed by the VBS file can thus be redirected to python: <br>
<b>Python file:</b>

```python
vbs = VBS("script.vbs")
string = vbs.run("askQuestion()")
 ```
 
 <b>VBScript file:</b>
 
 ```VB
Function askQuestion() 
    answer = inputbox("What is your name?")
    WScript.Echo answer
End Function 
 ```
 
 
