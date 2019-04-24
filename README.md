# vba-methods
Simple Functions to do things in Excel VBA For Microsoft Office 16. 

Functions are saved as .vba files to ensure the repository language is correct. 

## How to Contribute
PRs and Issues welcome. Use Issues for questions. Use PRs for new functions.
*Each file is a function*

Please remember that beginners will use this code. 
Add comments as needed.

### Conventions
We will try to use [VB Coding Standards](https://en.wikibooks.org/wiki/Visual_Basic/Coding_Standards) for consistency. 
The big picture means to 
- Use verb starts and camelCase for sub and function names. Like "getMoreFood()" and "returnLastCupcake()"
- Use NounDescripter for variable names. 
- there are a whole bunch of others


## How to Use
You can copy and paste functions directly into your Visual Basic Windows.

### Generic use of a Sub or Function
Let's assume that you want to incorporate the Sub printa(), Sub printStr(myStr), and Function getNextInt(myInt) in another Sub. Of course, this is a contrived example.

First, save your project with the new subs and/or functions in a Module

```vba
Sub printa()
   Debug.Print("a") 'prints a in the Immediate Window
End Sub
Sub printStr(myStr As String)
   Debug.Print(myStr) 'prints the passed String in the Immediate Window
End Sub

Function getNextInt(myInt As Integer)
  getNextInt = myInt + 1 ' returns the next Integer
End Function
```

Then, you can call it in your other Sub

```vba
Sub doOtherSub()
   Call printa
   Call printStr("I like Cupcakes")
   
   Debug.Print(getNextInt(4)) ' prints 5
   
   'You can also Dim and use variables
   Dim smyStrVar As String, imyIntVar As Integer
   smyStrVar = "Do you like cupcakes?"
   imyIntVar = 12
   
   Call printStr(smyStrVar)
   Debug.Print(getNextInt(imyIntVar)) 'prints 13
   
   ' You can write to variables with functions
   Dim imyNewInt As Integer
   imyNewInt = getNextInti(myIntVar) 'now, imyNewInt is 13
   
End Sub
```

 

