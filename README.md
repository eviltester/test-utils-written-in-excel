test-utils-written-in-excel
===========================

Testing Utilities written in Excel.

I used to regularly finding myself working on client sites where their machines were locked down and you couldn't install any software on them.

Fortunately the machines usually came with Excel. And since Excel allowed coding in VBA, I would write my own tools to help me.

This spreadsheet contains some of the code I used to use:
- generating counterstrings
- generating random strings
- generating files of certain length with random data
- locking files
- doing a file compare using the dos FC command
- opening command propmts

I haven't used it in years, but the code might be useful to someone.

It uses macros, so check out the code before you use it, so that you are happy with what it does.

Original Release Documentation
------------------------------

Over the years I have written too many spreadsheets, to aid me in my testing, for me to count and list for you. I've also written the same functions too often. So now I'm pulling those functions with simple routines into a single spreadsheet so that I have easy access to it, and so that anyone else who wants easy access to these functions can too.

The spreadsheet in its current form allows you to:

- generate counterStrings to clipboard, using 2 different algorithms
- generate counter strings to file
- lock files
- generate binary files of a defined length
- Provide a small front end for dos fc (file compare) command
- Open Command prompt @
- Generate ascii or unicode strings to the clipboard (within a range)

The spreadsheet is released under GPL.

Because the sheet uses macros you will have to enable them in order to use the functionality.

###Why Excel?

There are utilities to do this stuff already. There are scripts in numerous scripting languages to do this. So why write it in excel?

Well, one of the reasons I recommend to testers that they learn VBA is that every site I have worked on has given me access to different tools. Some give me Perl, some give me Java, some give me Unix. But every single one of them has given me a Microsoft Application that allows me to code in VBA.

Some sites absolutely refuse to allow me to install any software other than the software they give me. And so if I want to create any kind of test utilities, I use VBA.

Excel was the undisputed king of [alternative test tools](http://www.compendiumdev.co.uk/page.php?title=alternative_test_tools_in_action) and so I often created simple macros in Excel to suport me.

###CounterStrings

[James Bach](http://www.satisfice.com/) and [Danny Faught](http://www.tejasconsulting.com/) released a perl application to generate Counter Strings, and I have adapted that code so that I have a VBA implementation. Simply because these strings are too useful for me not to have them at my disposal when I need them.

A counter string, for those that don't know, is a string which is self documenting in terms of its length. So a 10 character string looks like this:

- *3*5*7*10*

I know it is a 10 character string because the last * has the number 10 before it, which tells me that it is in position 10. Easy.

[James has described counter strings on his blog](http://www.satisfice.com/blog/archives/22) and you can download his perl code for generating them to the clipboard.

I use two different algorithms for generating counter strings.

Backwards ( which is a variant of James') and Forwards (without prediction).

Both generate strings of the required length, but may end up with different results. Here is an example of a 15 char string:

- Backwards:
  - *3*5*7*9*12*15*
- Forwards:
  - 2*4*6*8*11*14*1

The backwards algorithm is generally the better output as the last set of strings are guaranteed to correspond to the actual length of the file.

The forwards algorithm is much more suited to writing the strings out to a file as I don't have to buffer the string and reverse it, which is the way James' algorithm works, although since it is written in Perl I imagine that the Perl implementation of the rev functions is actually in the same memory space and is actually very very efficient. Since I'm writing it in VBA, the implementation is pretty much guaranteed to be less efficient.

### Generating Files of Arbitrary Length

There are times when I have wanted random files of a very specific length in order to test limits in programs. And I no longer enjoy hacking files in hex editors.

So here is a routine I have written numerous times. The binary files have random binary data in them, and the text files have a counter string in them using the forward generation algorithm.

### Locking Files

I have previously written about using the dos command more for locking text files. Which is all very well if you are testing text files, but if you are testing binary files how are you going to lock it?

No doubt there is a dos command somewhere that does it, but instead of spending 20 mins looking for it, I spent 20 mins creating a small VBA subroutine to do it for me.

The code is very simple, I simply OPEN the file with a defined Lock mode, and since I didn't want to dictate which Lock mode you have to test with, you can use all 4.

- Shared (which doesn't lock it at all),
- Lock Write (which prevents anyone from opening it in write mode),
- Lock Read (which prevents anyone from opening it in read mode)
- Lock Read Write (which seems to prevent anyone else from using it at all, and is my favourite lock mode for testing)

This function is a handy tool for file exception testing.


Licensed Under GPL
--------------------------

http://www.gnu.org/copyleft/gpl.html