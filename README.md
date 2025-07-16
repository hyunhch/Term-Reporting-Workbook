Version 1.3.7

	Fixed bug with categories containing spaces not tabulating on MacOS

	Fixed bug importing on MacOS due to the above issue

Version 1.3.6	

	Fixed bug with saving a local copy from the Cover Page on MacOS
	
Version 1.3.5	

	Import function now works on MacOS
	
	Fixed bug with trailing spaces after activity names
	
	Fixed bug when creating a student with a single activity
	
	Fixed bug when deleting a row on an activity that has a single student
	
	
Version 1.3.4	

	Added import function on Cover Page
		
	Only visible cells are checked when pressing "Select all"
	
	The Report Sheet does not need an activity to export to SharePoint. Totals only are acceptable 
	
	Fixed bug where blank credit hours would break tabulation
	
	Issues with whitespace fixed (hopefully all of them)	
	
	Generated version number no longer pulls from the file name, to avoid problems if the workbook is renamed
	"Mathematics" changes to "Math" on the Report Page
	
	Support for MacOS, which cannot use dictionary objects. Procedures will now work on Mac, though they are slower
	The Other Page table now is generated programmatically
	
	Selecting a program is done with a button the Cover Page rather than automatically popping up. This was done to avoid issues when macros are disabled when the file is opened
	Breaking external links done on setup, though that shouldn’t be needed. 
	
Version 1.3.3	

	Fixed bug where all students on the roster would be added to an activity with the "Add Students" button, not just the selected students
	
Version 1.3.2	

	Only visible rows used when making an activity, adding students to an activity, or deleting students
	Fixed bug when deleting all students with an open activity sheet
	
	Fixed bug with alerts showing after reapplying sheet protection
	
	Student roster is no longer case sensitive
	
Version 1.3.1	

	Overhaul of workbook, most of the code rewritten
	
	Workbooks of the three programs merged into a single file that prompts for which program it will be used for
	
	Dramatic increase in efficiency, especially with deleting students
	
	Most tables are now programmatic, makign altering or expanding categories easier
	
	Added first-generation and low-income categories for Transfer Prep and MESA University
	
	Added grades down to 1st for College Prep
	
	Custom fields on the Roster Page will now be included in exported attendance reports
	
	The roster and report will be included when saving from the Cover Page
	
	Tabbing through userforms now follows an intuitive order
	
	Real-time filters added to all userforms that display activities
	
	The Report Page now has a table that can be sorted and filtered
	
	Exported attendence now brings up a save prompt and makes a file name. Detailed attendence always exported along with simple
	
	Name, center, and date are prompted by a userform when opening the file for the first time
	
	Many error messages were removed for button clicks when something hasn't been filled out. Now the buttons simply do nothing
	
	Adjusted the size and position of buttons
	
	The worksheet version now generates automatically by referencing the file name
	
	Student attendance is pulled into an activity sheet after students are added or removed
	
	Reference sheets for low-income and first-generation added on the Cover Page for Transer Prep and MESA University
	
	Parsing the roster will display the number of duplicate students removed, if any
	
	Adding students to an activity and removing missing students only displays the number of students. This was to accommodate very large lists of students
	
	Activities can be retabulated without removing them first
	
	Trying to load an open activity sheet now activates that sheet
	
	Import function temporarially removed
	
	Only recorded activities populate in to the New Activity form. Only recorded activities populate into the Add Students and Loa Activity form
	
	Added "Directory" page, reworked "Narrative" page
