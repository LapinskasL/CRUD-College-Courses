'******************************************************
'* Name:       Lukas Lapinskas
'* Class:      CIS-1510
'* Assignment: Project #1 Fall 2018
'* File:       frmCourses.vb
'* Purpose:    This program allows you to add, change,
'*             and view college courses. The program
'*             creates a CSV file called "data.csv" in
'*             debug folder and writes the data about
'*             a course on each row. Then, the data is
'*             read into the program, where it can be 
'*             manipulated.
'******************************************************


'*********************************************************************************************************
'* DEFENSIVE PROGRAMMING RULES:
'* 
'*                          Example course format -- CIS 1510-001 2018FA
'*
'*
'* Subject: this is the shortened version of the name of the course (ex: CIS, CIT, CHEM, BIOLO, etc.). It
'*          can be from 1 to 5 letters long. The letters are stored as all caps. This is a required field.
'*
'* Course #: this is the course number, like the number 1510 in the example course above. It can be any
'*          number from 0001 to 9999. If the user enters 1 or 450, it will be stored as 0001 and 0450,
'*          respectively. This is a required field.
'*
'* Section: the section is the number 001 in the example course. I can be any number from 001 to 999.
'*          Similarly to Course #, if the user enters only one digit, zeros will be added in front of
'*          the number to make it 3 digits long. This is a required field.
'*
'* Year: the year is the number 2018 in the example course. It is the year the course is being taught.
'*          Can be any number between 1000 and 9999. This is a required field.
'*
'* Faculty email: this is the email of the lecturer who will lecture the class. The email is an optional
'*          field, so it can be left empty. However, the email must end with either "@college.edu" or 
'*          "@county.edu". The local-part section (the part before @ symbol) can have any letter, digit,
'*          or an underscore. The email is limited to 30 characters in length.
'*
'*
'* What is a "course"?
'*
'* I defined a course to be the combination of the subject, course number, section, year, and term. 
'* This means that any change to either of those parts will constitute a different course. In other words,
'* if you have the course in the example above, any difference in the five parts mentioned will make it a 
'* new course, and it be allowed to be added to the list. Attempting to add a course with the same exact 
'* information will result in an error and it will not be added to the list.
'*********************************************************************************************************


Imports System.IO 'imported so I wouldn't need to type "IO." for IO.File

Public Class frmCourses

    Private Const FILENAME As String = "data.csv" 'holds the file name

    Private Const MAX_COURSES As Integer = 1000 'the max amount of courses possible
    Private num_of_courses As Integer = 0   'holds current number of courses in the file



    '@ form Load event handler
    Private Sub frmCourses_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        updateListBox() 'updates the list box with data from the file
    End Sub


    '@ mnuView mode Click event handler
    Private Sub mnuView_Click(sender As Object, e As EventArgs) Handles mnuView.Click
        'if the previous mode was New, focus on the group box. This is so that when a user adds a new course, the
        'focus doesn't go to the Exit button when going to Edit mode.
        If mnuNew.Checked = True Then
            gbxGroupBox.Select()
        End If

        mnuView.Checked = True 'checks the View menu item

        'if button Delete is disabled, reset all the fields. The reason for this if statement is so
        'that if no list item is selected in edit mode (ex, after deleting an item), the fields are
        'cleared if the user changes to the View mode.
        If btnDelete.Enabled = False Then
            resetAllFields()
        End If

        'if the mode is switched from New mode with fields filled to View mode, the fields are reset
        If btnSave.Enabled = True Then
            resetAllFields()
        End If

        'if in the Edit mode an item was selected and altered but not saved, and if the user then goes
        'to view mode, the fields are reloaded.
        If btnReloadOrClear.Enabled = True Then
            If btnReloadOrClear.Text = "Re&load" Then
                reloadData()
            End If
        End If

        changeMode("Courses (Lukas Lapinskas #32) - View Courses", "View Course Details",
                    "View") 'a call to a procedure that changes the form components based on the mode
    End Sub


    '@ mnuEdit mode Click event handler
    Private Sub mnuEdit_Click(sender As Object, e As EventArgs) Handles mnuEdit.Click
        'if the previous mode was New, focus on the group box. This is so that when a user adds a new course, the
        'focus doesn't go to the Exit button when going to Edit mode.
        If mnuNew.Checked = True Then
            gbxGroupBox.Select()
        End If

        'if the mode is switched from New mode with fields filled to Edit mode, the fields are reset
        If btnSave.Enabled = True Then
            If btnReloadOrClear.Text = "C&lear" Then
                resetAllFields()
            End If

        End If
        changeMode("Courses (Lukas Lapinskas #32) - Edit Courses", "Edit Course Details",
                    "Edit") 'a call to a procedure that changes the form components based on the mode

        'if no list item was selected and the user goes into view mode, all fields are disabled,
        'the delete button is hidden, and the reload/clear button is hidden
        If lstCourses.SelectedIndex = -1 Then
            enableAllFields(False)
            btnDelete.Enabled = False
            btnReloadOrClear.Enabled = False
        End If

        'if the Reload button is enabled then reload the data. The reason for this if statement is if the user
        'is in Edit mode, makes some adjustments to an item, doesnt save it, and then goes to edit mode again,
        'the data of the item selected doesn't get erased. Instead, the data is reloaded.
        If btnReloadOrClear.Enabled Then
            reloadData()
        End If

    End Sub


    '@ mnuNew mode Click event handler
    Private Sub mnuNew_Click(sender As Object, e As EventArgs) Handles mnuNew.Click
        changeMode("Courses (Lukas Lapinskas #32) - New Course", "New Course Details",
                    "New") 'a call to a procedure that changes the form components based on the mode

        'clears any selected items from list box to avoid confusion when returning back to other modes
        lstCourses.ClearSelected()

        'resets all fields and disables the appropriate buttons
        resetAllFields()
        btnSave.Enabled = False
        btnReloadOrClear.Enabled = False
    End Sub


    '@ mnuAbout button Click event handler
    Private Sub mnuAbout_Click(sender As Object, e As EventArgs) Handles mnuAbout.Click
        'displays an About message box
        MsgBox("This program allows you to create a list of courses." & ControlChars.NewLine & ControlChars.NewLine &
               "You can add a course in the New mode, change its data in the Edit mode, and display the data in " &
               "View mode." & ControlChars.NewLine & ControlChars.NewLine &
               "©2018 Lukas Lapinskas", MsgBoxStyle.Information, "About")
    End Sub


    '@ Save button Click event handler
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Dim courseForList As String = "" 'holds the course information to put into list
        Dim lengthOfID As Integer   'holds the length of a fieldID
        Dim fieldIDNumber As Integer 'holds fieldID used to add a new item

        Dim outFile As StreamWriter 'used to store a StreamWriter object 

        'if there was an error in any textbox filled out by the user, this procedure stops executing
        If errorExists() Then
            Return
        End If

        'formats course information into sometihng that looks like the ones in the list box, but without the field
        'number in parentheses. Ex: "CIS 1510-001 2018FA"
        courseForList = txtSubject.Text.ToUpper & " " & txtCourseNum.Text.PadLeft(4, "0") & "-" &
                        txtSection.Text.PadLeft(3, "0") & " " & txtYear.Text & getCheckedRad().Substring(0, 2).ToUpper

        '@ if the user is in Edit mode
        If mnuEdit.Checked = True Then

            Dim IDofField As Integer    'holds the fieldID of an item

            'holds each row of the file (ReadAllLines closes the file automatically when it's done reading)
            Dim splitRow() As String = File.ReadAllLines(FILENAME)
            Dim rowIndex As Integer = 0 'index to access each row in splitRow array

            'for loop iterates the same amount of times as there are items in list box
            For index As Integer = 0 To lstCourses.Items.Count - 1

                'gets the length of fieldID of an item from the list box
                lengthOfID = getIDLength(index)

                '(The reason for the if-statement below is so that when modifying a course, the user isn't
                'able to modify it to a course that already exists).
                'if the list item at the index position is not the selected item, contains the 
                'course name (ex: "CIS 1510-001 2018FA"), and is the same length as the course name + lengthOfID
                '+ 3 ( the +3 is the space and two parentheses that surround the fieldID), then....
                If lstCourses.Items(index) <> lstCourses.SelectedItem AndAlso
                   lstCourses.Items(index).ToString.Contains(courseForList) AndAlso
                   lstCourses.Items(index).ToString.Length = (courseForList.Length + lengthOfID + 3) Then

                    msgBoxError("The course " & courseForList &
                                " could not be updated because such a course already exists!") 'display error msg
                    Return 'stop executing the event handler
                End If
            Next


            IDofField = getItemFieldID() 'get the selected item's (in the list box) fieldID

            'loop until the fieldID matches the fieldID of a row in the splitRow array
            Do Until IDofField = splitRow(rowIndex).Split(","c)(0)
                rowIndex += 1
            Loop

            'replace that row with the information in the fields filled out by the user 
            splitRow(rowIndex) = IDofField & "," & txtSubject.Text.ToUpper & "," & txtCourseNum.Text.PadLeft(4, "0") &
                              "," & txtSection.Text.PadLeft(3, "0") & "," & txtYear.Text &
                              "," & txtFacEmail.Text & "," & getCheckedRad() & "," & isChecked(chkOnline) &
                              "," & isChecked(chkPrereqReq)

            rowIndex = 0 'resets the rowIndex to 0

            'opens the file for overwriting the information inside it, overwrites it with the rows in splitRow array
            'containing the changes the user made to a row, and closes the file.
            outFile = File.CreateText(FILENAME)
            Do Until rowIndex > splitRow.Length - 1
                outFile.WriteLine(splitRow(rowIndex))
                rowIndex += 1
            Loop
            outFile.Close()

            lstCourses.Items.Clear() 'clears the list box of old data
            updateListBox() 'refills the list box with the updated data from the file

            enableAllFields(False) 'disables all fields (because no list item is selected)
            enableButtons(False, False, False) 'disables appropriate buttons because no list item is selected

            gbxGroupBox.Select() 'focuses on the group box so that the focus does not go to the Exit button

            'displays a message box, stating that the changes to the course were successful
            msgBoxSuccess("Changes to the course " & courseForList & " were saved.")


            '@ the code for Save button for "Edit" mode above
            '-----------------------------------------------------------------------------------------------------
            '@ the code for Save button for "New" mode below


            '@ if the user is in New mode
        ElseIf mnuNew.Checked = True Then

            'if the number of courses is more or equal to 1000, an error is displayed and procedure stops
            'executing
            If num_of_courses >= 1000 Then
                msgBoxError("Cannot have more than 1000 courses.")
                Return
            End If

            'this whole for loop basically checks if the same course already exists. If it does,
            'the course is not added and an error message is displayed. This for loop is similar
            'to the for loop above, except that this one does not care what list item is selected.
            For index As Integer = 0 To lstCourses.Items.Count - 1

                lengthOfID = getIDLength(index)

                If lstCourses.Items(index).ToString.Contains(courseForList) AndAlso
                   lstCourses.Items(index).ToString.Length = (courseForList.Length + lengthOfID + 3) Then

                    msgBoxError("The course " & courseForList & " cannot be saved because it already exists!")
                    Return
                End If
            Next

            'gets the largest fieldID currently in the file and adds one to it (to use for the new item)
            fieldIDNumber = getLargestFieldID() + 1

            'opens the file for appending text, formats and appends data entered by the user in the fields to the
            'last line of the file. Then, the file is closed.
            outFile = File.AppendText(FILENAME)
            outFile.WriteLine(fieldIDNumber & "," & txtSubject.Text.ToUpper & "," & txtCourseNum.Text.PadLeft(4, "0") &
                              "," & txtSection.Text.PadLeft(3, "0") & "," & txtYear.Text & "," &
                              txtFacEmail.Text & "," & getCheckedRad() & "," & isChecked(chkOnline) & "," &
                              isChecked(chkPrereqReq))
            outFile.Close()

            lstCourses.Items.Add(courseForList & " (" & fieldIDNumber & ")") 'adds the new item to the list box

            num_of_courses += 1 'increments the number of courses

            resetAllFields()    'resets all fields
            txtSubject.Select() 'focuses on the Subject text box
            btnSave.Enabled = False 'disables the Save button
            btnReloadOrClear.Enabled = False 'disables the Clear button

            'displays a message box, stating that the course was added successfully.
            msgBoxSuccess("The course " & courseForList & " has been added.")

        End If

    End Sub


    '@ Delete button Click event handler
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim outputFile As StreamWriter
        Dim theFieldID As Integer
        Dim splitRowArr() As String = File.ReadAllLines(FILENAME)

        Dim theRowIndex As Integer = 0 'holds the index of an element in splitRowArr
        Dim courseForMsgBox As String = "" 'holds the course (ex: "CIS 1510-001 2018FA") to use in the msg box.

        Dim msgBoxAnswer As MsgBoxResult 'holds a yes or no, depending on user's answer to dialog box

        theFieldID = getItemFieldID() 'gets the fieldID of the selected item in the list box

        'loops until the fieldID above is found in a row of a file 
        Do Until theFieldID = splitRowArr(theRowIndex).Split(",")(0)
            theRowIndex += 1
        Loop

        'format the row's subject, courseNum, section, year, and term to use in msgBox
        courseForMsgBox = splitRowArr(theRowIndex).Split(",")(1) & " " & splitRowArr(theRowIndex).Split(",")(2) & "-" &
                          splitRowArr(theRowIndex).Split(",")(3) & " " & splitRowArr(theRowIndex).Split(",")(4) &
                          splitRowArr(theRowIndex).Split(",")(6).Substring(0, 2)

        msgBoxAnswer = MsgBox("Are you sure you want to delete the course " & courseForMsgBox & "?",
                              MsgBoxStyle.YesNo, "Confirmation") 'display a dialog box

        'if user picks Yes in the dialog box
        If msgBoxAnswer = MsgBoxResult.Yes Then
            'replace the row with "DELETED" string
            splitRowArr(theRowIndex) = "DELETED"

            theRowIndex = 0 'reset row index

            'open the file for overwriting, loop until the end of the file. In each iteration, write every row to the
            'file except for the one marked "DELETED". Then, close the file. In other words, these statements rewrite
            'all of the data in the file except for the row the user wants deleted.
            outputFile = File.CreateText(FILENAME)
            Do Until theRowIndex > splitRowArr.Length - 1
                If splitRowArr(theRowIndex) <> "DELETED" Then
                    outputFile.WriteLine(splitRowArr(theRowIndex))
                End If
                theRowIndex += 1
            Loop
            outputFile.Close()

            lstCourses.Items.Clear() 'clear the items in the list
            updateListBox() 'update the lsit box with the file contents, now without the deleted row

            'resets and disables all fields since no list item is selected and since the deleted item does not exist 
            'to not confuse the user
            resetAllFields()
            enableAllFields(False)

            'disables appropriate buttons
            enableButtons(False, False, False)

            gbxGroupBox.Select() 'focuses on the group box so that the focus does not go to the Exit button

            'displays a message box, stating that the course was deleted successfully.
            msgBoxSuccess("The course " & courseForMsgBox & " was deleted.")
        End If

    End Sub


    '@ Reload or Clear button Click event handler
    Private Sub btnReloadOrClear_Click(sender As Object, e As EventArgs) Handles btnReloadOrClear.Click

        'if the button is the Reload button, then reload all the data
        If btnReloadOrClear.Text = "Re&load" Then
            reloadData()

            'if the button is the Clear button, then reset the fields, focus on Subject textbox,
            'and disable the Save and Clear buttons
        ElseIf btnReloadOrClear.Text = "C&lear" Then
            resetAllFields()
            hideAllRedAsterisks()
            txtSubject.Select()
            btnSave.Enabled = False
            btnReloadOrClear.Enabled = False
        End If

    End Sub


    '@ Exit button Click event handler
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close() 'closes the program
    End Sub


    '@ TextChanged and CheckedChanged event for every field (textboxes, rad buttons, and check boxes).
    Private Sub anyField_TextChanged(sender As Object, e As EventArgs) Handles txtSubject.TextChanged,
                                       txtCourseNum.TextChanged, txtSection.TextChanged, txtYear.TextChanged,
                                       txtFacEmail.TextChanged, radFall.CheckedChanged,
                                       radSpring.CheckedChanged, radSummer.CheckedChanged, chkOnline.CheckedChanged,
                                       chkPrereqReq.CheckedChanged

        'if the user is either in Edit or New mode (if the user is not in View mode)
        If mnuView.Checked = False Then
            'if the Save button is disabled, then enable it
            If btnSave.Enabled = False Then
                btnSave.Enabled = True
            End If

            'if the Reload or Clear button is disabled, then enable it
            If btnReloadOrClear.Enabled = False Then
                btnReloadOrClear.Enabled = True
            End If
        End If

    End Sub


    '@@@@@ five TextChanged event handlers for every text box. They all hide the appropriate red asterisk
    'when text is changed.
    Private Sub txtSubject_TextChanged(sender As Object, e As EventArgs) Handles txtSubject.TextChanged
        hideRedAsterisk(lblSubjectError)
    End Sub
    Private Sub txtCourseNum_TextChanged(sender As Object, e As EventArgs) Handles txtCourseNum.TextChanged
        hideRedAsterisk(lblCourseNumError)
    End Sub
    Private Sub txtSection_TextChanged(sender As Object, e As EventArgs) Handles txtSection.TextChanged
        hideRedAsterisk(lblSectionError)
    End Sub
    Private Sub txtYear_TextChanged(sender As Object, e As EventArgs) Handles txtYear.TextChanged
        hideRedAsterisk(lblYearError)
    End Sub
    Private Sub txtFacEmail_TextChanged(sender As Object, e As EventArgs) Handles txtFacEmail.TextChanged
        hideRedAsterisk(lblFacEmailError)
    End Sub


    '@ SelectedIndexChanged event handler for the list box
    Private Sub lstCourses_SelectedIndexChanged(sender As Object, e As EventArgs) Handles _
                                                                                  lstCourses.SelectedIndexChanged

        Dim itemID As Integer 'holds fieldID

        'if an item is selected in the list and the user is not in the New mode
        If lstCourses.SelectedIndex <> -1 Then
            If mnuNew.Checked = False Then

                itemID = getItemFieldID() 'get fieldID from selected item in list box

                populateFields(itemID) 'fill in the fields of the item selected

                'if, additionally, the user is in the Edit mode (instead of View), enable
                'all of the fields so the user can modify them
                If mnuEdit.Checked = True Then
                    enableAllFields(True)
                End If

                'disable and enable appropriate buttons
                enableButtons(False, True, False)

            End If
        End If

    End Sub


    '@ KeyPress event handler for Subject text box
    Private Sub txtSubject_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSubject.KeyPress
        'Allows any letter or the backspace key
        If e.KeyChar < "A" OrElse e.KeyChar > "Z" Then
            If e.KeyChar < "a" OrElse e.KeyChar > "z" Then
                If e.KeyChar <> ControlChars.Back Then
                    e.Handled = True
                End If
            End If
        End If

    End Sub


    '@ KeyPress event handler for Course #, Section, and Year text boxes
    Private Sub digitsOnly_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCourseNum.KeyPress,
                                                                                      txtSection.KeyPress,
                                                                                      txtYear.KeyPress
        'Allows only digits or the backspace key
        If e.KeyChar < "0" OrElse e.KeyChar > "9" Then
            If e.KeyChar <> ControlChars.Back Then
                e.Handled = True
            End If
        End If

    End Sub


    '@ KeyPress event handler for Faculty Email text box
    Private Sub txtFacEmail_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtFacEmail.KeyPress
        'Allows digits, letters, and the backspace key
        If e.KeyChar < "A" OrElse e.KeyChar > "Z" Then
            If e.KeyChar < "a" OrElse e.KeyChar > "z" Then
                If e.KeyChar < "0" OrElse e.KeyChar > "9" Then
                    If e.KeyChar <> ControlChars.Back Then
                        e.Handled = True
                    End If
                End If
            End If
        End If

        'allows the character _ @ and .
        If e.KeyChar = "_" OrElse e.KeyChar = "@" OrElse e.KeyChar = "." Then
            e.Handled = False
        End If
    End Sub






    '@ All event handler procedures end here
    '@@@ @-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@-@
    '@ Other procedures and functions begin below






    '@ updateListBox procedure
    'Updates the list box with formatted data from the file.
    Private Sub updateListBox()
        Dim inFile As StreamReader
        Dim wholeLine As String 'holds a whole line from a file
        Dim splitString() As String 'holds a line from a file, split by the comma char

        'if the file data.csv exists
        If File.Exists(FILENAME) Then
            inFile = File.OpenText(FILENAME) 'opens the file for reading

            num_of_courses = 0 'resets the number of courses

            'loops until there's no more lines in the file. This loop formats a line in the file and adds the
            'data to the list box.
            Do Until inFile.Peek = -1
                wholeLine = inFile.ReadLine()
                splitString = wholeLine.Split(","c)
                lstCourses.Items.Add(splitString(1) & " " & splitString(2) & "-" & splitString(3) & " " &
                                     splitString(4) & splitString(6).Substring(0, 2) &
                                     " (" & splitString(0) & ")")
                num_of_courses += 1 'number of courses is incremented

            Loop

            inFile.Close()
        End If
    End Sub


    '@ errorExists function that returns a boolean value.
    'This function checks if all the text boxes are filled in correctly. If they are, the function returns 
    'False. If they are not, the function returns True.
    Private Function errorExists() As Boolean

        'if Subject text box is empty
        If txtSubject.Text = "" Then
            displayErrorMessage(txtSubject, lblSubjectError, "You must enter the course's subject.")
            Return True

            'else if Course text box is empty
        ElseIf txtCourseNum.Text = "" OrElse txtCourseNum.Text < 1 Then
            displayErrorMessage(txtCourseNum, lblCourseNumError,
                                "You must enter the course's number from 0001 to 9999.")
            txtCourseNum.SelectAll()
            Return True

            'else if Section text box is empty
        ElseIf txtSection.Text = "" OrElse txtSection.Text < 1 Then
            displayErrorMessage(txtSection, lblSectionError, "You must enter the course's section from 001 to 999.")
            txtSection.SelectAll()
            Return True

            'else if the Year text box doesn't contain a 4 digit number over 1000 (this includes if it's empty)
        ElseIf Not txtYear.Text Like "[1-9]###" Then
            displayErrorMessage(txtYear, lblYearError, "You must enter the year of the course that's 1000 or higher.")
            txtYear.SelectAll() 'select text in text box
            Return True

            'else if the Email texbox is not like *@college.edu or *@county.edu or is not empty
        ElseIf emailLocalPartBad(txtFacEmail.Text) OrElse Not txtFacEmail.Text Like "*@college.edu" AndAlso
            Not txtFacEmail.Text Like "*@county.edu" Then
            If txtFacEmail.Text <> "" Then
                displayErrorMessage(txtFacEmail, lblFacEmailError,
                               "The faculty email must be in one of the following forms: " & ControlChars.NewLine &
                               "Na_me123@college.edu  OR  Na_me123@county.edu")
                txtFacEmail.SelectAll() 'select text in text box
                Return True
            End If
        End If

        Return False 'no errors exist, so return false (errorExists = false)

    End Function


    '@ emailLocalPartBad function that returns a boolean value.
    'Takes one parameter: email as a string.
    'This function checks if the local-part section of the email contains a . or and @ symbol. If it does, it returns
    'True, if it doesn't it returns False.
    Private Function emailLocalPartBad(email As String) As Boolean
        Dim atNum As Integer = 0    'holds the number of @ symbols in the whole email address
        Dim dotNum As Integer = 0   'holds the number of . characters in the whole email address

        'if the length of the email is more than just one character
        If email.Length > 1 Then
            'if the first character is an @ symbol
            If email(0) = "@" Then
                Return True
            End If
        End If

        'loops through the length of the email
        For i As Integer = 0 To email.Length - 1
            'if the character in email is an @ symbol
            If email(i) = "@" Then
                atNum += 1  'increment atNum variable
            End If
            'if the character in email is a dot
            If email(i) = "." Then
                dotNum += 1 'increment dotNum variable
            End If
        Next

        'if the number of @ symbols or dots is more than one
        If atNum > 1 OrElse dotNum > 1 Then
            Return True
        End If

        Return False 'Return False if the procedure doesn't return True (emailLocalPartBad = False)
    End Function


    '@ displayErrorMessage procedure.
    'Takes 3 parameters: textField as a TextBox, lblError as a Label, errMessage as a String
    'This procedure focuses on the correct text box, makes red asterisk visible, and displays an error message.
    Private Sub displayErrorMessage(txtField As TextBox, lblError As Label, errMessage As String)
        lblError.Visible = True
        txtField.Select()
        msgBoxError(errMessage)
    End Sub


    '@ resetAllFields procedure.
    'This procedure resets all the data fields (text boxes, radio buttons, and checknboxes) to how they are when
    'the program is first launched.
    Private Sub resetAllFields()
        txtSubject.Text = ""
        txtCourseNum.Text = ""
        txtSection.Text = ""
        txtYear.Text = ""
        txtFacEmail.Text = ""
        radFall.Checked = True
        chkOnline.Checked = False
        chkPrereqReq.Checked = False
    End Sub


    '@ reloadData procedure.
    'This procedure reloads the data into the text boxes, radio buttons, and checkboxes of a selected item in the
    'list box and also disables any appropriate buttons.
    Private Sub reloadData()
        Dim readFile As StreamReader
        Dim peekVal As Integer 'holds the peek value (if -1, then the file is empty)

        'if the file exists
        If File.Exists(FILENAME) Then
            readFile = File.OpenText(FILENAME) 'open the file for reading
            peekVal = readFile.Peek() 'assign peek value of the file to peekVal
            readFile.Close() 'close the file

            'if the peek value was not -1 (if the file was not empty)
            If peekVal <> -1 Then
                populateFields(getItemFieldID()) 'populate the fields with selected item
                btnSave.Enabled = False 'disable Save button
                btnReloadOrClear.Enabled = False 'disable Reload/Clear button
                gbxGroupBox.Select() 'focuses on the group box so that the focus does not go to the Exit button
            End If
        End If

    End Sub


    '@ getIDLength function that returns an Integer.
    'Takes one parameter: indexOfListItem as an Integer.
    'This function gets the length of the fieldID from the selected list item.
    Private Function getIDLength(indexOfListItem As Integer) As Integer
        Dim indexLeftParenth As Integer     'holds the index of the left parenthese
        Dim indexRightParenth As Integer    'holds the index of the right parenthese

        'the two statements find the index of right and left parenthese from the selected item in the list and
        'assigns them to the appropriate variables.
        indexLeftParenth = lstCourses.Items(indexOfListItem).IndexOf("("c)
        indexRightParenth = lstCourses.Items(indexOfListItem).IndexOf(")"c)

        'returns how many chars is in the fieldID
        Return indexRightParenth - indexLeftParenth - 1
    End Function


    '@ getLargestFieldID function that returns an Integer.
    'This function finds the line in the data.csv file with the largest fieldID and returns that value.
    Private Function getLargestFieldID() As Integer

        Dim infile As StreamReader
        Dim highestNum As Integer = 0   'holds the fieldID
        Dim aLine As String             'holds a line from the file
        Dim splitLine() As String       'holds each element of the line, split by the comma character

        'if the file exists
        If File.Exists(FILENAME) Then
            infile = File.OpenText(FILENAME) 'open file

            'loop until no more lines in the file
            Do Until infile.Peek = -1
                aLine = infile.ReadLine 'read a line and assign it to variable aLine
                splitLine = aLine.Split(","c)    'split the line by commas and assign the elements to splitLine array

                'if the first element is higher than value in highestNum variable
                If splitLine(0) > highestNum Then
                    highestNum = splitLine(0) 'assign the value of first element to highestNum
                End If
            Loop

            infile.Close() 'close file
        End If

        Return highestNum 'return the value in the highestNum variable. This value is the highest fieldID in the file

    End Function


    '@ getItemFieldID function that returns an integer value.
    'This function returns the fieldID of a selected item in the list box.
    Private Function getItemFieldID() As Integer

        Dim currentItem As String = ""  'holds the text of the currently selected item in the list box
        Dim idOfField As Integer        'holds the fieldID
        Dim indexLParenth As Integer    'holds the index of the left parenthese
        Dim indexRParenth As Integer    'hodls the index of the right parenthese

        'assigns the text of the currently selected list item to the currentItem variable
        currentItem = lstCourses.SelectedItem.ToString

        'the two statements below assign the index of left and right parentheses to appropriate variables
        indexLParenth = currentItem.IndexOf("("c)
        indexRParenth = currentItem.IndexOf(")"c)

        'using the indexes of left and right parentheses, the fieldID number is located and assigned to idOfField
        idOfField = currentItem.Substring(indexLParenth + 1, indexRParenth - indexLParenth - 1)

        Return idOfField 'the fieldID number is returned

    End Function


    '@ getCheckedRad function that returns a String value.
    'This function returns either the string FALL, SPRING, or SUMMER based on which radio button is checked.
    Private Function getCheckedRad() As String
        'if Fall radio button is checked
        If radFall.Checked Then
            Return "FALL"

            'if Spring radio button is checked
        ElseIf radSpring.Checked Then
            Return "SPRING"

            'else (if Summer radio button is checked)
        Else
            Return "SUMMER"
        End If

    End Function


    '@ isChecked function that returns an Integer value.
    'Takes onE parameter: chkBox as a CheckBox.
    'If the check box is checked, returns 1, if it's not, returns 0
    Private Function isChecked(chkBox As CheckBox) As Integer

        If chkBox.Checked Then
            Return 1
        Else
            Return 0
        End If

    End Function


    '@ changeMode procedure.
    'Takes three parameters: formText as a String, grpBoxText as a String, mode as a String.
    'This procedure makes several changes to the components in the program, depending on what argument was
    'passed to the mode parameter.
    Private Sub changeMode(formText As String, grpBoxText As String, mode As String)

        Me.Text = formText 'changes the text of the form
        gbxGroupBox.Text = grpBoxText 'changes the text of the group box

        hideAllRedAsterisks() 'hides any red asterisks in case some were visible

        'if mode is "View" then call procedures to alter the user interface
        If mode = "View" Then
            enableMnuCheck(True, False, False)
            enableAllFields(False)
            visibleButtons(False, False, False)
            lstCourses.Visible = True 'make the list box visible

            'if mode is "Edit" then call procedures to alter the user interface
        ElseIf mode = "Edit" Then
            enableMnuCheck(False, True, False)
            enableAllFields(True)
            visibleButtons(True, True, True)
            lstCourses.Visible = True 'make the list box visible

            btnSave.Enabled = False 'disable the save button
            btnReloadOrClear.Text = "Re&load" 'change the btnReloadOrClear text to Re&load

            'if mode is "New" then call procedures to alter the user interface
        ElseIf mode = "New" Then
            enableMnuCheck(False, False, True)
            enableAllFields(True)
            visibleButtons(True, False, True)
            lstCourses.Visible = False

            btnSave.Enabled = False 'disable the save button
            btnReloadOrClear.Text = "C&lear" 'change the btnReloadOrClear text to C&lear
        End If

    End Sub


    '@ enableMnuCheck procedure.
    'Takes three parameters: viewCheck, editCheck, and newCheck as a Booleans.
    'This procedure changes which menu items (either View, Edit, or New) are checked.
    Private Sub enableMnuCheck(viewCheck As Boolean, editCheck As Boolean, newCheck As Boolean)
        mnuView.Checked = viewCheck
        mnuEdit.Checked = editCheck
        mnuNew.Checked = newCheck
    End Sub


    '@ enableAllFields procedure.
    'Takes one parameter: fieldsEnabled as a Boolean.
    'This procedure either disables or enables all of the fields on the form depending on the value of fieldsEnabled.
    Private Sub enableAllFields(fieldsEnabled As Boolean)
        txtSubject.Enabled = fieldsEnabled
        txtCourseNum.Enabled = fieldsEnabled
        txtSection.Enabled = fieldsEnabled
        txtYear.Enabled = fieldsEnabled
        txtFacEmail.Enabled = fieldsEnabled
        radFall.Enabled = fieldsEnabled
        radSpring.Enabled = fieldsEnabled
        radSummer.Enabled = fieldsEnabled
        chkOnline.Enabled = fieldsEnabled
        chkPrereqReq.Enabled = fieldsEnabled
    End Sub


    '@ visibleButtons procedure.
    'Takes three parameters: saveBtn, deleteBtn, and reloadClearBtn as Booleans. 
    'Makes first three buttons on the form visible or invisible, depending on the parameters' Boolean values.
    Private Sub visibleButtons(saveBtn As Boolean, deleteBtn As Boolean, reloadClearBtn As Boolean)
        btnSave.Visible = saveBtn
        btnDelete.Visible = deleteBtn
        btnReloadOrClear.Visible = reloadClearBtn
    End Sub


    '@ enableButtons procedure.
    'Takes three parameters: saveButton, deleteButton, and reloadClearButton as Booleans. 
    'Makes first three buttons on the form enabled or disabled, depending on the parameters' Boolean values.
    Private Sub enableButtons(saveButton As Boolean, deleteButton As Boolean, reloadClearButton As Boolean)
        btnSave.Enabled = saveButton
        btnDelete.Enabled = deleteButton
        btnReloadOrClear.Enabled = reloadClearButton
    End Sub


    '@ populateFields procedure.
    'Takes one parameter: ID as an Integer.
    'This procedure populates the fields with data from the file. The data depends on which item is selected in
    'the list box (the ID parameter is matched with the field ID of the item selected in the list box).
    Private Sub populateFields(ID As Integer)
        Dim inFile As StreamReader
        Dim wholeString As String = ""  'holds a line from the file
        Dim splitString() As String     'hold a row of data from the file that has been split by the comma character

        'if the file exists
        If File.Exists(FILENAME) Then

            'the do-until loop below reads each line of the file and checks if the first element (the fieldID) 
            'matches the value of the ID parameter
            inFile = File.OpenText(FILENAME) 'open the file for reading
            Do
                wholeString = inFile.ReadLine() 'a line form a file is assigned to wholeString
                splitString = wholeString.Split(","c) 'line is split by the comma char
            Loop Until splitString(0) = ID OrElse inFile.Peek = -1
            inFile.Close() 'close file 

            'the five statements below assign an element of splitString to the appropriate text box
            txtSubject.Text = splitString(1)
            txtCourseNum.Text = splitString(2)
            txtSection.Text = splitString(3)
            txtYear.Text = splitString(4)
            txtFacEmail.Text = splitString(5)

            'Depending on whether the element at index 6 of splitString is FALL, SPRING, or SUMMER, the proper
            'radio button is checked.
            If splitString(6) = "FALL" Then
                radFall.Checked = True
            ElseIf splitString(6) = "SPRING" Then
                radSpring.Checked = True
            ElseIf splitString(6) = "SUMMER" Then
                radSummer.Checked = True
            End If

            'If index 7 of splitString is a 1, the Online check box is checked. Otherwise (if 0), it is unchecked.
            If splitString(7) = 1 Then
                chkOnline.Checked = True
            Else
                chkOnline.Checked = False
            End If

            'If index 8 of splitString is a 1, the Prerequisites Req. checkbox is checked. Otherwise (if 0), it
            'is unchecked.
            If splitString(8) = 1 Then
                chkPrereqReq.Checked = True
            Else
                chkPrereqReq.Checked = False
            End If

        End If

    End Sub


    '@ hideAllRedAsterisks procedure.
    'Hides all red asterisk labels (this procedure is called each time modes are switched).
    Private Sub hideAllRedAsterisks()
        hideRedAsterisk(lblSectionError)
        hideRedAsterisk(lblCourseNumError)
        hideRedAsterisk(lblSectionError)
        hideRedAsterisk(lblYearError)
        hideRedAsterisk(lblFacEmailError)
    End Sub


    '@ hideRedAsterisk procedure.
    'Takes one parameter: redAsteriskLabel as a Label.
    Private Sub hideRedAsterisk(redAsteriskLabel As Label)
        'if the red asterisk is visible
        If redAsteriskLabel.Visible = True Then
            redAsteriskLabel.Visible = False 'make it invisible
        End If
    End Sub


    '@ msgBoxSuccess procedure.
    'Takes one parameter: message as a String.
    Private Sub msgBoxSuccess(message As String)
        MsgBox(message, MsgBoxStyle.Information, "Success") 'displays a message box noting success
    End Sub


    '@ msgBoxError procedure.
    'Takes one parameter: errMessage as a String.
    Private Sub msgBoxError(errMessage As String)
        MsgBox(errMessage, MsgBoxStyle.Critical, "Error") 'displays a message box noting an error
    End Sub

End Class
