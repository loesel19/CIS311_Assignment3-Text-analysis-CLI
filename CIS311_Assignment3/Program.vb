Imports System.IO
Imports System.Text.RegularExpressions
Module Program
    '------------------------------------------------------------------------
    '-                    File Name: Program                                -
    '-                    Part of Project: Text File Analyzer (CIS 311: HW3)-
    '------------------------------------------------------------------------
    '-                      Written By: Andrew A. Loesel                    -
    '-                      Written On: January 21, 2022                    -
    '------------------------------------------------------------------------
    '- File Purpose:                                                        -
    '-                                                                      -
    '- This file contains the code for the invoice form of the Kustom Karz  -
    '- Order system. This file accesses data that contains the users input  -
    '- from frmMain, and dispays/ performs calculations to display informati-
    '- on such as body type and price, drive-train selected and price, optio-
    '- ns requested and total price for all options, prive per vehicle, numb-
    '- er of vehicles requested, and total price of order. This file also   -
    '- controls the closing and loading of this form, and some control of   -
    '- frmMain's closing operation                                          -
    '-                                                                      -
    '------------------------------------------------------------------------
    '- Program Purpose:                                                     -
    '-                                                                      -
    '- The purpose of this program is to act as an order system for Kustom  -
    '- Karz, an automobile shop. It is made up of two forms, frmMain, and   -
    '- frmInvoice. The two forms will comminucate and be controled in a fash-
    '- ion that allows the User to select options for their order on frmMain-
    '- and review the options on frmInvoice. The user can also navigate betw-
    '- een the two forms without losing the selected options. frmInvoice wil-
    '- l display an invoice of the order to the customer. From there the cus-
    '- tomer can either exit the application, submit the order, or change th-
    '- e order.                                                             -
    '------------------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically):                         -
    '- intOptionExtra - An integer variable which holds the total cost of al-
    '-                  l the extra options the user selected on frmMain.   -
    '- intPerVehicleTotal - An integer which holds the total cost per vehicl-
    '-                    - e.                                              -
    '- intVehicleTypeCost - An integer which holds the cost of the vehicle  -
    '-                    - type the user selected on frmMain.              -
    '- strLine            - A string which contains a line of '='s to act as-
    '-                    - a seporator.                                    -
    '------------------------------------------------------------------------

    'GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS
    'GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS
    Const CLI_CHARACTER_MAX = 100 'this will be the maximum amount of characters any line the application prints out can be
    Const REPORT_FREQUENCY_DIGITS = 4 'the number of digits to be shown for word frequencies in the report file

    'GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES
    'GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES
    Dim strPathToRead As String
    Dim strPathToWrite As String
    Dim dicWords As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)

    'SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS
    'SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS
    Sub Main(args As String())
        '------------------------------------------------------------------------
        '-                      Subprogram Name: fileToReadInteraction          -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 27, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '-                   -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- None                                                                 -
        '------------------------------------------------------------------------
        'set background color and foreground color(text) to desired colors
        Console.BackgroundColor = ConsoleColor.White
        Console.ForegroundColor = ConsoleColor.Blue
        'clear the current console so that nothing from initialization shows up with different colors
        Console.Clear()
        'set the title
        Console.Title = "Word Analysis Profiler Application"
        fileToReadInteraction()
        fileToWriteInteraction()
        analyzeFile()
        displayReport()

    End Sub

    Sub fileToReadInteraction()
        '------------------------------------------------------------------------
        '-                      Subprogram Name: fileToReadInteraction          -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 27, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '-                   -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- None                                                                 -
        '------------------------------------------------------------------------
        Console.WriteLine("Please enter the path and name of the file to process:")
        strPathToRead = Console.ReadLine
        'see if file exists, if it does return true
        If System.IO.File.Exists(strPathToRead) Then
            If strPathToRead.Contains(".txt") Then

            Else
                Console.WriteLine("The file path/name provided was not for a .txt file. Please enter the path and name of a .txt file.")
            End If
        Else
            Console.WriteLine("No such file found, check file path/name.")
            fileToReadInteraction()
        End If
    End Sub

    Sub fileToWriteInteraction()
        '------------------------------------------------------------------------
        '-                      Subprogram Name: fileToReadInteraction          -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 27, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '-                   -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- None                                                                 -
        '------------------------------------------------------------------------
        Console.WriteLine("Please enter the path and name of the report file to generate:")
        strPathToWrite = Console.ReadLine
        'ensure that the user does not try to write to the file that is being analyzed
        If strPathToRead.Equals(strPathToWrite) Then
            Console.WriteLine("Can not write and read to the same file.")
            fileToWriteInteraction()
        End If
        If Not System.IO.File.Exists(strPathToWrite) Then
            'try to create a file with the provided path and name
            Try
                Dim fs As FileStream = System.IO.File.Create(strPathToWrite)
                fs.Close()
            Catch ex As Exception
                'Console.WriteLine(ex.StackTrace)
                Console.WriteLine("Could not create a file with that name.")
                fileToWriteInteraction()
            End Try
        Else
            'if file already exists we want to delete it so that we do not have stuff before our report in the file
            System.IO.File.Delete(strPathToWrite)
        End If

    End Sub

    Function analyzeFile() As Dictionary(Of String, Integer)

        Dim rgWordsFromFile As String()
        Using fileReader As StreamReader = New StreamReader(strPathToRead, True)
            'analyze every line of the file (if it has multiple lines)
            While Not fileReader.EndOfStream
                rgWordsFromFile = Split(fileReader.ReadLine, " ")
                analyzeLine(rgWordsFromFile)
            End While
        End Using
        'now we want to sort our dictionary, we use a LINQ query to get a sorted IEnumerable of our original dictionary,
        'and then we create a new dictionary from the IEnumerable
        Dim dicSortedWords = (From entry In dicWords Order By entry.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        'get information about word frequency
        writeReportFile(dicSortedWords, getStatistics(dicSortedWords), getLongestWordLength())
    End Function

    Sub analyzeLine(rgWordsFromFile As String())
        For i As Integer = 0 To rgWordsFromFile.Count - 1
            Dim strCurrentWord As String = rgWordsFromFile(i)
            'make sure that we are not just looking at an empty string
            If strCurrentWord.Length > 0 Then
                'see if word contains ' ', '.', or ',' and either replace or trim the characters with empty string
                strCurrentWord = strCurrentWord.Replace(",", "")
                strCurrentWord = strCurrentWord.Replace(".", "")
                strCurrentWord = strCurrentWord.Replace("?", "")
                strCurrentWord = strCurrentWord.Replace("!", "")
                strCurrentWord.Trim()
                'now to make it case insensitive we will do ToUpper
                strCurrentWord = strCurrentWord.ToUpper
                'check word length again since we replaced some characters
                If strCurrentWord.Length > 0 Then
                    'check to see if our dictionary already contains the word. if so we increment the integer of the entry
                    If dicWords.ContainsKey(strCurrentWord) Then
                        dicWords.Item(strCurrentWord) += 1
                    Else
                        'else we add the word to the dictionary and set its value to 1
                        dicWords.Add(strCurrentWord, 1)
                    End If
                End If
            End If
        Next

    End Sub

    Function getLongestWordLength() As Integer
        Dim intLongest As Integer = 0
        'look for the longest word in the dictionary, notice we are using the global structure here,
        'because the order of the keys(words) does not matter in this instance
        For Each entry As KeyValuePair(Of String, Integer) In dicWords
            If entry.Key.Length > intLongest Then
                intLongest = entry.Key.Length
            End If
        Next
        Return intLongest
    End Function
    Function getStatistics(dicSortedWords As Dictionary(Of String, Integer)) As Double()
        Dim intMax As Integer = 0 'start max as 0 so any word will have a larger frequency to start
        Dim intMin As Integer = 2147483647 'start min at the maximum value of an int so the first word frequency will be less than this
        Dim intTotal As Integer = 0
        For Each entry As KeyValuePair(Of String, Integer) In dicSortedWords
            If entry.Value > intMax Then
                intMax = entry.Value
            End If
            If entry.Value < intMin Then
                intMin = entry.Value
            End If
            intTotal += entry.Value
        Next
        Dim rgMinMaxAvg As Double() = {intMin, intMax, CDbl(intTotal / dicSortedWords.Count)}
        Return rgMinMaxAvg
    End Function

    Sub writeReportFile(dicSortedWords As Dictionary(Of String, Integer), rgMinMaxAvg As Double(), intLongestWordLength As Integer)
        Dim intNumberOfStars As Integer = CLI_CHARACTER_MAX
        Dim strLowFrequency As String = ""
        Dim strHighFrequency As String = ""
        'the most stars that will be in a line will be the max character for line limit - (digits for frequency + length
        'of the longest word + 4). the 4 is for a space and colon after the word, then a space before and after the frequency
        intNumberOfStars = intNumberOfStars - (REPORT_FREQUENCY_DIGITS + intLongestWordLength + 4)
        'we then divide this by the maximum frequency which will be the second element of the rgMinMaxArray
        'we do integer division since we can not have partial stars
        intNumberOfStars = intNumberOfStars \ rgMinMaxAvg(1)
        'make sure that at least 1 start shows up
        If intNumberOfStars = 0 Then
            intNumberOfStars = 1
        End If
        Using fileWriter As StreamWriter = New StreamWriter(strPathToWrite, True)

            'write each line to the file with the word, frequency, and then number of stars indicating frequency
            For Each entry As KeyValuePair(Of String, Integer) In dicSortedWords
                Dim intSpacesToAdd As Integer = intLongestWordLength - (entry.Key.Length - 1)
                Dim strLine As String = entry.Key & Space(intSpacesToAdd) & ": " & String.Format("{0,4:d4} ", entry.Value)
                strLine = strLine & StrDup((intNumberOfStars * entry.Value), "*")
                fileWriter.WriteLine(strLine)
                'add words that show up with the min or max frequency to the corresponding String
                If entry.Value = rgMinMaxAvg(0) Then
                    'shows up fewest
                    'check if string is empty, if so we will not put a ',' before the first word
                    If strLowFrequency.Length < 1 Then
                        strLowFrequency = strLowFrequency & entry.Key
                    Else
                        'here we add a comma and space before the word
                        strLowFrequency = strLowFrequency & ", " & entry.Key
                    End If

                ElseIf entry.Value = rgMinMaxAvg(1) Then
                    'shows up the most
                    'check if string is empty, if so we will not put a ',' before the first word
                    If strHighFrequency.Length < 1 Then
                        strHighFrequency = strHighFrequency & entry.Key
                    Else
                        'here we add a comma and space before the word
                        strHighFrequency = strHighFrequency & ", " & entry.Key
                    End If
                End If
            Next
            'now we have to write the utilization, and words that showed up the most, and the fewest
            fileWriter.WriteLine("")
            fileWriter.WriteLine("Average Word Utilization: " & rgMinMaxAvg(2))
            fileWriter.WriteLine("Highest Word Utilization: " & rgMinMaxAvg(1) & " on " & strHighFrequency)
            fileWriter.WriteLine("Lowest Word Utilization: " & rgMinMaxAvg(0) & " on " & strLowFrequency)
            'write a message in the report if the max is > CLI_CHAR_LIMIT - (Longest word length + 8)
            If (CLI_CHARACTER_MAX - (intLongestWordLength + 8)) < rgMinMaxAvg(1) Then
                fileWriter.WriteLine("-*-*-*Histogram will be innacurate due to limitations on how long the line can be. Certain words occurred too many times to be accurately represented*-*-*-")
            End If

        End Using
    End Sub

    Sub displayReport()
        Console.WriteLine("Would you like to see the report file? [Y/n]")
        Dim strInput = Console.ReadLine()
        If strInput.Trim.Equals("Y") Or strInput.Trim.Equals("y") Then
            Using fileReader As StreamReader = New StreamReader(strPathToWrite, True)
                While Not fileReader.EndOfStream
                    Console.WriteLine(fileReader.ReadLine)
                End While
            End Using
        ElseIf strInput.Trim.Equals("N") Or strInput.Trim.Equals("n") Then
            'do nothing, console will close itself
        Else
            Console.WriteLine("Please press either Y or N ...")
            displayReport()
        End If
    End Sub
End Module
