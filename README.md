Microsoft Access Date Picker
============================

Brendan Kidwell

15 January 2003

Introduction
------------

I've been programming with Microsoft Access for a few years now, and I've found it to be an excellent database development platform, despite its many shortcomings.

One of my pet peeves with Microsoft Access has been its lack of a built-in control or dialog box for graphically picking a date. There are some nice ActiveX controls out there you can reference from your Access programs, but this has proven to be an unreliable solution for me. I have clients running a database of mine at a few different locations, and some of them are running Access 97 while others are running Access 2000 or 2002. I can't guarantee that a particular ActiveX control is installed there, and talking them through an install or automating the install aren't easy to do.

`DatePicker.mdb` contains a module and a form that together implement a date picker function using only intrinsic Access controls and Visual Basic code. It should be compatible with Microsoft Access 97, Microsoft Access 2000, and Microsoft Access 2002.

Adding DatePicker to Your Database
----------------------------------

To add DatePicker to your database, use the `Import` command (File / Get External Data / Import) to copy in the form `DatePicker` and the module `mdlDatePicker` from `DatePicker.mdb`.

The best way to use the date picker dialog box on a date field in your form is to create a text box and a command button for the field. (See the sample form in `DatePicker.mdb`.) Bind the text box to a date field in a table as you normally would. Then add something like the following code to the comand button's `Click` event:

```
Private Sub DateButton_Click()
   InputDateField DateTextBox, "Some prompt"
End Sub
```

(Hint: When you create the button to call `InputDateField()` you might want to give it a blank caption and assign the built-in picture &quot;Calendar&quot; to it.)

Another way to use the module is to call the `InputDate()` method, like this:

```
Private Sub SomeAction()
   Dim d As Variant

   ' Initialize whatever needs initializing
   
   d = InputDate("Some prompt")
   ' d is now a date (IsDate(d) = True) or Null, if the user
   ' cancelled.

   ' Do something with useful with d.
End Sub
```

Please note that I am a lazy American and the code is very American English. It will use English names for months no matter where you run it, and I'm not sure how it will behave on computers configured to use European style dates. If you're using this outside of North America I suggest you inspect and test it carefully before you run with it.

Synopsis
--------

```
' Use this method to prompt for and set a new date on a textbox
Public Sub InputDateField(x As TextBox, Optional Prompt As _
   String = "Select Date")
```

When you call `InputDateField()` you must pass a reference to a `TextBox`, optionally followed by a string to display in the `DatePicker` form's caption. The `DatePicker` form initializes with the date (if any) contained in `x` highlighted.

```
' Use this method to prompt for a date inside a procedure
Public Function InputDate(Optional Prompt As String = "Select Date", _
   Optional InitDate As Variant) As Variant
```

When you call `InputDate()` you may optionally pass a string to display in the `DatePicker` form's caption, followed optionally by a `Variant` specifying the date to hightlight when the form initializes. The method returns a `Variant` containing either the chosen date or, if the user clicked the `Cancel` button, `Null`.
