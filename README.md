# vba-methods
Simple Functions to do things in Excel VBA

Functions are saved as .vba files to ensure the repository language is correct. 

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
Sub otherSub()
   Call printa
   Call printStr("I like Cupcakes")
   
   Debug.Print(getNextInt(4)) ' prints 5
   
   'You can also Dim and use variables
   Dim myStrVar As String, myIntVar As Integer
   myStrVar = "Do you like cupcakes?"
   myIntVar = 12
   
   Call printStr(myStrVar)
   Debug.Print(getNextInt(myIntVar)) 'prints 13
   
   ' You can write to variables with functions
   Dim myNewInt As Integer
   myNewInt = getNextInt(myIntVar) 'now, myNewInt is 13
   
End Sub
```

 

