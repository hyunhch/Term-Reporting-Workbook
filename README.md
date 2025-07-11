Version 1.3.4		
	Added import function on Cover Page	
		Select a different workbook with the same structure to copy students and activity information over
		Differences in spelling or added/removed fields can be matched, skipped, or added as a new column
		Applies to demographics fields and activities
		
	Only visible cells are checked when pressing "Select all"	
	
	The Report Sheet does not need an activity to export to SharePoint. Totals only are acceptable 	
	
	Fixed bug where blank credit hours would break tabulation	
	
	Issues with whitespace fixed (hopefully all of them)	
		Leading and trailing spaces for activities
		Leading and trailing spaces for student names, leading to duplicates
	
	Generated version number no longer pulls from the file name, to avoid problems if the workbook is renamed	
	
	"Mathematics" changes to "Math" on the Report Page	
	
	Support for MacOS, which cannot use dictionary objects. Procedures will now work on Mac, though they are slower	
	
	The Other Page table now is generated programmatically	
	
	Selecting a program is done with a button the Cover Page rather than automatically popping up. This was done to avoid issues when macros are disabled when the file is opened	
	
	Breaking external links done on setup, though that shouldn’t be needed. 	
