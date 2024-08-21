' ---Standardize Highlighting With Exceptions Macro v2.0.1---
' Updated on 2024-08-16.
' https://github.com/KSXia/Verbatim-Standardize-Highlighting-With-Exceptions-Macro
' Based on Verbatim 6.0.0's "UniHighlightWithException" function.
Sub StandardizeHighlightingWithExceptions()
	Dim ExceptionColors() As Variant
	
	' ---USER CUSTOMIZATION---
	' <<SET THE HIGHLIGHTING COLORS THAT SHOULD NOT BE STANDARDIZED HERE!>>
	' Add the names of highlighting colors that you want to exempt from standardization to the list in the ExceptionColors array. Make sure that the name of every highlighting color is in quotation marks and that each term is separated by commas.
	' NOTE: This macro does NOT automatically exempt the highlighting color you have set to be exempted in the Verbatim settings. You MUST MANUALLY enter the highlighting colors you would like to exempt into this list.
	'
	' These are the names of the highlighting colors in the each row of the highlighting color selection menu, listed from left to right:
	' First row: Yellow, Bright Green, Turquoise, Pink, Blue
	' Second row: Red, Dark Blue, Teal, Green, Violet
	' Third row: Dark Red, Dark Yellow, Dark Gray, Light Gray, Black
	' MAKE SURE TO USE THIS EXACT CAPITALIZATION AND SPELLING!
	'
	' If you are using gray highlighting, you are likely using the color Light Gray
	'
	' Warning: There needs to be at least one hightlighting color listed in the ExceptionColors array for this macro to work.
	ExceptionColors = Array("Light Gray", "Pink")
	
	' ---INITIAL SETUP---
	Dim r As Range
	Set r = ActiveDocument.Range
	
	Dim GreatestIndex As Integer
	GreatestIndex = UBound(ExceptionColors) - LBound(ExceptionColors)
	
	' ---CONVERT HIGHLIGHTING COLOR NAMES TO VBA INDEXES---
	Dim ExceptionEnums() as Long
	ReDim ExceptionEnums(0 To GreatestIndex) As Long
	For CurrentIndex = 0 to GreatestIndex Step +1
		Select Case ExceptionColors(CurrentIndex)
			Case Is = "None"
				ExceptionEnums(CurrentIndex) = wdNoHighlight
			Case Is = "Black"
				ExceptionEnums(CurrentIndex) = wdBlack
			Case Is = "Blue"
				ExceptionEnums(CurrentIndex) = wdBlue
			Case Is = "Bright Green"
				ExceptionEnums(CurrentIndex) = wdBrightGreen
			Case Is = "Dark Blue"
				ExceptionEnums(CurrentIndex) = wdDarkBlue
			Case Is = "Dark Red"
				ExceptionEnums(CurrentIndex) = wdDarkRed
			Case Is = "Dark Yellow"
				ExceptionEnums(CurrentIndex) = wdDarkYellow
			Case Is = "Light Gray"
				ExceptionEnums(CurrentIndex) = wdGray25
			Case Is = "Dark Gray"
				ExceptionEnums(CurrentIndex) = wdGray50
			Case Is = "Green"
				ExceptionEnums(CurrentIndex) = wdGreen
			Case Is = "Pink"
				ExceptionEnums(CurrentIndex) = wdPink
			Case Is = "Red"
				ExceptionEnums(CurrentIndex) = wdRed
			Case Is = "Teal"
				ExceptionEnums(CurrentIndex) = wdTeal
			Case Is = "Turquoise"
				ExceptionEnums(CurrentIndex) = wdTurquoise
			Case Is = "Violet"
				ExceptionEnums(CurrentIndex) = wdViolet
			Case Is = "White"
				ExceptionEnums(CurrentIndex) = wdWhite
			Case Is = "Yellow"
				ExceptionEnums(CurrentIndex) = wdYellow
			Case Else
				ExceptionEnums(CurrentIndex) = wdNoHighlight
		End Select
	Next CurrentIndex
	
	' ---MORE SETUP---
	' Disable screen updating for faster execution
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
	
	' ---REHIGHLIGHTING---
	With r.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Highlight = True
		.Replacement.Highlight = True
		.Text = ""
		.Replacement.Text = ""
		.Forward = True
		.Wrap = wdFindStop
		.Format = True
		.MatchCase = False
		.MatchWholeWord = False
		.MatchWildcards = False
		.MatchSoundsLike = False
		.MatchAllWordForms = False
		
		Do While .Execute(Forward:=True) = True
			' Check if the color of the current word is one of the exceptions
			Dim IsException As Boolean
			IsException = False
			Dim i
			For i = LBound(ExceptionEnums) To UBound(ExceptionEnums)
				If r.HighlightColorIndex = ExceptionEnums(i) Then
					IsException = True
				End If
			Next I

			If IsException Then
				' If the color of the current word is an exception:
				r.Collapse Direction:=wdCollapseEnd
			Else
				' If the color of the current word is not an exception:
				' Set the highlighting of the current word to the default highlighting color
				r.HighlightColorIndex = Options.DefaultHighlightColorIndex
			End If
		Loop
		
		.ClearFormatting
		.Replacement.ClearFormatting
	End With
	
	' ---FINAL PROCESSES---
	' Re-enable screen updating and alerts
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub
