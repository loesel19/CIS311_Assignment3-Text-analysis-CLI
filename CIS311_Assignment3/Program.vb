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
    '- This file contains the code for a text-file analysis console Line App-
    '- lication. This is the only Module of the program. All user input and -
    '- file I\O is performed in this file.                                  -
    '-                                                                      -
    '------------------------------------------------------------------------
    '- Program Purpose:                                                     -
    '-                                                                      -
    '- The purpose of this program is to analyze a text file of the users ch-
    '- oosing. The analysis consists of figuring out how many times each wor-
    '- d shows up, what words show up the fewest and most times, and how man-
    '- y times on average does any word show up, as well as generating a his-
    '- togram to represent the distribution of word frequency in the file.  -
    '- all of the analysis data is written out to a file that the user speci-
    '- fies. Lastly, the user may choose to view the analysis of the file in-
    '- the console window.                                                  -
    '------------------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically):                         -
    '- dicWords - a dictionary that will contain the KeyValuePairs of each  -
    '-            each word from the file and the amount of times the word  -
    '-            shows up.                                                 -
    '- strPathToRead - the string containing the file path that the user wan-
    '-                 ts to analyze. (user specified)                      -
    '- strPathToWrite - A string containing the file path of the report file-
    '-                  the user wants to generate. (user specified)        -
    '------------------------------------------------------------------------

    'GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS
    'GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS
    Const CLI_CHARACTER_MAX = 100 'this will be the maximum amount of characters any line the application prints out can be
    Const REPORT_FREQUENCY_DIGITS = 4 'the number of digits to be shown for word frequencies in the report file
    '^^^^^could not find a way to implement this constant since string.format placeholders dont allow 
    '^^^^^variable values
    Const BACKGROUND_COLOR = ConsoleColor.White 'Background color of the console
    Const FOREGROUND_COLOR = ConsoleColor.Blue 'foreground color of the console

    'GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES
    'GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES
    Dim strPathToRead As String
    Dim strPathToWrite As String
    Dim dicWords As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)

    'SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS
    'SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS
    Sub Main(args As String())
        '------------------------------------------------------------------------
        '-                      Subprogram Name: Main                           -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 30, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- This is the main subprogram of the Application. The purpose of this  -
        '- subroutine is to direct the flow of the application by calling the   -
        '- other subprocedures. We also set the background\ foreground colors   -
        '- and the title of the console.                                        -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- args - A string array that holds the command line arguments for appli-
        '- cation start(if any, none for this application)                      -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- None                                                                 -
        '------------------------------------------------------------------------
        'set background color and foreground color(text) to desired colors
        Console.BackgroundColor = BACKGROUND_COLOR
        Console.ForegroundColor = FOREGROUND_COLOR
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
        '-                      Written On: January 30, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subprogram is to have the user input the name and-
        '- path of the file that they want to have analyzed. If the file does   -
        '- exist or does not contain the '.txt' extension the method is called  -
        '- recursively until the user inputs a valid file. The file path is then-
        '- saved to the global string 'strPathToRead'                           -
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
                fileToReadInteraction()
            End If
        Else
            Console.WriteLine("No such file found, check file path/name.")
            fileToReadInteraction()
        End If
    End Sub

    Sub fileToWriteInteraction()
        '------------------------------------------------------------------------
        '-                      Subprogram Name: fileToWriteInteraction         -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 30, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subroutine is to get the file name and path to   -
        '- write the report to from the user. If they enter the same file as in -
        '- fileToReadInteraction the console will tell them that they can not   -
        '- read and write to the same file, and calls this function recursively.-
        '- If the file does not exist we try to create it, if unsuccesful we mak-
        '- e a recursive call to this function again, telling the user that some-
        '- thing went wrong. If the file already exists we delete it so that our-
        '- report does not get written after some other text in the file.       -
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

    Sub analyzeFile()
        '------------------------------------------------------------------------
        '-                      Subprogram Name: analyzeFile                    -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 30, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this function is to handle the analysis of the file.  -
        '- we analyze the file line by line using a streamreader and looping thr-
        '- ough every line of the file and sending an array of the split words  -
        '- to the subroutine analyzeLine(). After each line has gone through ana-
        '- lyzeLine() we perform a LINQ query on the public dictionary of word, -
        '- frequency keyValuePairs to sort them alphabetically by key, and we se-
        '- nd that sorted dictionary as well as a 3-element array with the numbe-
        '- r of occurences of the least popular and most popular words, as well -
        '- as the average appearance of all words in the file, and the length   -
        '- of the longest word in the file.
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- dicSortedWords - A dictionary containing key value pairs of word(stri-
        '-                  ng) and frequency(int), which is assigned to the res-
        '-                  ults of a LINQ query to sort the words alphabeticall-
        '-                  y.                                                  -
        '- rgWordsFromFile - a String array that contains the split words from a-
        '-                   line in the file.                                  -
        '------------------------------------------------------------------------
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
    End Sub

    Sub analyzeLine(rgWordsFromFile As String())
        '------------------------------------------------------------------------
        '-                      Subprogram Name: analyzeLine                    -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 30, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- the purpose of this subroutine is to take an array of strings contain-
        '- ing a line of words from the file, and loop through every string in  -
        '- the array to remove any weird characters or spaces from the string an-
        '- d adds the word to the dictionary with value 1 if the key is not in  -
        '- the dictionary. If the key(word) is in the dictionary it increments  -
        '- the value by 1.
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- rgWordsFromFile - A string array that contains all the words of a lin-
        '-                   e from the file we are analyzing.                  -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- strCuurentWord - A string that contains the current word in the array-
        '-                  that is being processed in the loop.                -
        '------------------------------------------------------------------------
        For i As Integer = 0 To rgWordsFromFile.Count - 1
            Dim strCurrentWord As String = rgWordsFromFile(i)
            'make sure that we are not just looking at an empty string
            If strCurrentWord.Length > 0 Then
                'see if word contains certain characters and replace them with the empty string.
                strCurrentWord = strCurrentWord.Replace(",", "")
                strCurrentWord = strCurrentWord.Replace(".", "")
                strCurrentWord = strCurrentWord.Replace("?", "")
                strCurrentWord = strCurrentWord.Replace("!", "")
                strCurrentWord = strCurrentWord.Replace("(", "")
                strCurrentWord = strCurrentWord.Replace(")", "")
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
        '------------------------------------------------------------------------
        '-                      Subprogram Name: getLongestWordLength           -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 30, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this function is to loop through every keyValuePair in-
        '- our public dictionary to figure out what the longest word length is. -
        '- we then return the length as an integer.                             -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- intLongest - an integer that represents the longest length of any wor-
        '-              d encountered thus far in our dictionary.               -
        '------------------------------------------------------------------------
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
        '------------------------------------------------------------------------
        '-                      Subprogram Name: getStatistics                  -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 30, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this function is to loop through each entry in our dic-
        '- tionary of words to see which word(s) shows up the most, which show  -
        '- up the fewest and what the average utilization of a word is. We then -
        '- place these 3 values into an array of doubles to return.             -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- dicSortedWords - a dictionary containing keyValuePairs of words in al-
        '-                  phabetical order.                                   -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- intMax - an integer that will represent the maximum frequency of a wo-
        '-          rd in the file. Starts at zero so that the first value will -
        '-          set it to a value from the dictionary.                      -
        '- intMin - an integer that will represent the minimum frequency of a wo-
        '-          rd in the file. Initially set to the max value of an int in -
        '-          vb so that after the first iteration of the loop it will hav-
        '-          e a value from the dictionary.                              -
        '- intTotal - an integer that will be a sum of all values from the dicti-
        '-            onary so that the average frequency can be calculated.    -
        '- rgMinMaxAvg - a double array that will have the min, max, and average-
        '-               that this function calculates set to its elements. This-
        '-               will be returned by the function.                      -
        '------------------------------------------------------------------------
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
        '------------------------------------------------------------------------
        '-                      Subprogram Name: writeReportFile                -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 30, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subroutine is to write the reportFile. Before the-
        '- file gets written we have to know how many stars will be used to repr-
        '- esent a single word occurance. To calculate this we subtract the leng-
        '- th of the longest word + 2 (for ' :') and 6 (space before and after  -
        '- four digits) from the constant CLI_CHARACTER_MAX, and then we divide -
        ' this number by the maximum word frequency. There is a problem here, if-
        ' we have a word that shows up more times than whatever the that number -
        '- is we will end up with 0 for the number of stars since we use integer-
        '- division. To handle this I added an if statement that sets numberOfSt-
        '- ars to 1 if it is 0 so that our application doesnt just have a list o-
        '- f words. However, if this happens the histogram will be messed up if -
        '- the user decides to see it in the console, so a message is printed at-
        '- the end of the file stating so. After this we use the StreamWriter an-
        '- d loop through every word in the dictionary, spacing it properly, wit-
        '- h a calculation for how many spaces should follow the word, and then -
        '- we write the frequency followed by the numberOfStars * frequency. We -
        '- then write that string to the file. While in the loop we also see if -
        '- the frequency = min or max frequency, if so we add the word to a stri-
        '- g containing the fewest and most frequent words. After the loop we wr-
        '- ite the min, max and average utilization. Min and Max are followed by-
        '- the words that appeared that many times, and after that we print out -
        '- the messy histogram message if necessary.                            -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- dicSortedWords - a dictionary containing keyValuePairs of the words  -
        '-                  from the input file.                                -
        '- rgMinMaxAvg - a double array countaining the minimum\ maximum word   -
        '-               frequencies and the average frequency.                 -
        '- intLongestWord - an integer representing the longest length of any wo-
        '-                  rd from the input file.                             -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- intNumberOfStars - an integer which represents the number of stars to-
        '-                    be associated with a single word occurance.       -
        '- intSpacesToAdd - an integer representing the number of spaces to add -
        '-                  for each word in the dictionary, set programatically-
        '-                  in the for each loop.                               -
        '- strHighFrequency - a string containing a comma delimited list of all -
        '-                    words that show up with the maximum frequency in  -
        '-                    the file.                                         -
        '- strLine - a string which is used to hold the contents of the current -
        '-           line that we wish to write out to the file.                -
        '- strLowFrequency - a string containing a comma delimited list of all  -
        '-                   words that show up the fewest times in the file.   -
        '------------------------------------------------------------------------
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
            fileWriter.WriteLine("Highest Word Utilization: " & CInt(rgMinMaxAvg(1)) & " on " & strHighFrequency)
            fileWriter.WriteLine("Lowest Word Utilization: " & CInt(rgMinMaxAvg(0)) & " on " & strLowFrequency)
            'write a message in the report if the max is > CLI_CHAR_LIMIT - (Longest word length + 8)
            If (CLI_CHARACTER_MAX - (intLongestWordLength + 8)) < rgMinMaxAvg(1) Then
                fileWriter.WriteLine("-*-*-*Histogram will be innacurate due to limitations on how long the line can be. Certain words occurred too many times to be accurately represented*-*-*-")
            End If

        End Using
    End Sub

    Sub displayReport()
        '------------------------------------------------------------------------
        '-                      Subprogram Name: displayReport                  -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: January 30, 2022                    -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- the purpose of this subroutine is to see whether or not the user woul-
        '- d like to see the report file displayed on the console. We check the -
        '- input by catching it in strInput. If the user says 'y' we loop throug-
        '- h every line of the file and simply print it out to the console. If  -
        '- the user says no we do nothing(the program will terminate if displayR-
        '- eport is the last subprogram called from main), and if the user enter-
        '- ed anything else we ask them to either press y or n again.           -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- strInput - A string which holds the user input.                      -
        '------------------------------------------------------------------------
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
