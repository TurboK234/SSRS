; SSRS - String Search and Replace Script
; Copyright (c) 2018 Henrik Söderström
; This script is published under Creative Commons Attribution-ShareAlike 4.0 International (CC BY-SA 4.0) licence.
; You are free to use the script as you please, even commercially, but you should publish the edited code
; under the same licence and give the original creator appropriate credit. More information about the licence
; can be found at http://creativecommons.org/licenses/by-sa/4.0/ .

; GENERAL SETUP, DON'T EDIT, COMMENTS PROVIDED FOR CLARIFICATION.
SendMode, Input  													; Recommended for new scripts due to its superior speed and reliability.
#NoEnv  															; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance FORCE												; There can be only one instance of this script running.
SetWorkingDir, %A_ScriptDir% 										; Ensures a consistent starting directory (can be overwritten later in this script).
StringCaseSense, On													; Turns on case-sensitivity, which helps to create more specific string searches.
temp_extension = xtt												; This sets the filetype/extension used when seeking (and replacing) inside text files. Leave untouched, unless you really have xtt-files (?)
FileEncoding, UTF-8													; The default encoding is UTF-8. You can change it later to something else, see General Conversion Preferences.
AutoTrim, Off														; Allows using variables (=searches) with leading or trailing spaces. Trimming is done separately.

; GENERAL CONVERSION PREFERENCES (USER CONFIRMATION REQUIRED). DO NOT REMOVE THE 2x DOUBLE QUOTES, FILL THE VALUE BETWEEN THEM, USE "" FOR EMPTY.
dir_files_ssrs := "C:\Sourcedir"												; The complete path (without the last "\") of the files to be renamed
extension_files_ssrs := "ext"											; The extension of the original files. Wildcards are not acceptable.
recursive_ssrs = no												; Search also subdirectories of dir_files_ssrs .
rename_files = yes													; Use the script for renaming files (the main purpose the script was written for)											
search_inside_text_files = no										; Search and replace text inside text files. BE SURE that the extension is set to text files if this is "yes".
textfile_encoding = UTF-8											; If using "search_inside_text_files", make sure that this setting matches the encoding of the files. Preferably this script should also be saved in the same format for the replacements to work. See AHK-guide for options.
days_before_search = 0												; # of days before the file is processed. This option looks at the source file's modification time, and considers only dates and rounds up (i.e. file modified yesterday -> value = 1). Use 0 to edit/rename files regardless of their age. This is a safety measure for automated setups, not to convert files that are currently being written / recorded. As such, "1" is usually a good number.
show_launch_confirmation_ssrs = no									; Show a OK/Cancel dialogue (for 10 sec) before starting the execution.
logging = yes														; Keep log (in folder dir_files_ssrs) of the renamed files. Good for debugging and automated work flow. Note: logging will happen regardless of this option, only the log file (also the pre-existing!) will be deleted afterwards if "no".

; INITIAL SETUP VALIDATION
GoSub, init_zero_ssrs													; This subscript mainly checks if the %dir_files_ssrs% is a valid location, as logging of any kind and the script function requires this. Gives a huge error and exits if not found.

; INITIAL QUERY TO EXECUTE THE SCRIPT OR NOT (CAN BE EASILY DISABLED BY ADDING A SEMICOLON (;) IN FRONT OF THE LINE BELOW)
If (show_launch_confirmation_ssrs = "yes")
{
	GoSub, query_runscript_ssrs											; Confirm the execution of the script with OK/Cancel, if the option is set.
}

; STRINGS TO BE SERARCHED FOR AND REPLACED:
; (THERE CAN BE MORE THAN ONE, PLEASE COPY THE <----...----> MARKED SECTION FOR EACH RULE)
; DO NOT REMOVE THE 2x DOUBLE QUOTES, FILL THE VALUE BETWEEN THEM, USE "" FOR EMPTY.


; THIS RULE IS ONLY AN EXAMPLE (IT IS COMMENTED OUT WITH SEMICOLONS), FEEL FREE TO REMOVE
; ; <----------------------------------------------------------------------
; ; Copy each section starting from the line above
; ; Rule title: *** Remove _OLD -tags from filenames and replace with _ARCHIVED ***		; non-formal title for user reminder
; string_search_ssrs := "_OLD"													; The string in file's name to be searched for.
; string_replace_ssrs := "_ARCHIVED"												; The string that replaces the original string in the file's name. Use "" for empty (the string will be exctracted from the filename).
; string_exclusion_rule_ssrs := "_NEW"											; Optional: User can fill here one string that causes the file to be skipped if found in the file name.
; ; The following line needs to be copy-pasted for each rule.
; Gosub, init_rule
; ;--------------
; ; Copy to the end of this line (and paste below) to add new rules ------>
; ;
; THE EXAMPLE ENDS HERE <-



; -----------------------------------------------------------------------------------------------
; ACTUAL DATA PROCESSING ENGINE(S) BELOW THIS LINE, EDITING THIS DATA CAN EASILY BREAK THE SCRIPT

; GENERAL SETUP VALIDATION
GoSub, init_global_ssrs													; This subscript checks if all of the prerequisites are valid (disk space and global required variables, for example). Failing one of the tests will most likely end the script.

process_main_ssrs:
Loop, Files, %dir_files_ssrs%\*.%extension_files_ssrs%, %is_recursive_ssrs%
{
	If (A_LoopFileName = "ssrs_log.txt")
	{
		FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : SSRS log file matches search, skipping...`n", %dir_files_ssrs%\ssrs_log.txt
		continue		; Skip the log file and continue to the next line.
	}
	
	FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : File " A_LoopFileName " found, starting the search...`n", %dir_files_ssrs%\ssrs_log.txt

	If (days_before_search >= 1)
	{
		StringLeft, current_date, A_Now, 8												; saves current date as new YYYYMMDD variable.
		FileGetTime, mod_time_current_file, %A_LoopFileDir%\%A_LoopFileName%, M		; saves the modification time of the current file
		StringLeft, mod_date_current_file, mod_time_current_file, 8						; saves year, month and date to the new variable in YYYYMMDD form.
		source_mod_age = %current_date%													; to avoid confusion on the next line, copy current_date to a variable that will eventually only show the difference of dates (in days).
		EnvSub, source_mod_age, %mod_date_current_file%, Days
		If (source_mod_age < days_before_search)
		{
			FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : File " A_LoopFileName " (" source_mod_age " day(s) since modified) is newer than days_before_search defines, skipping the file.`n", %dir_files_ssrs%\ssrs_log.txt
			continue
		}
	}
	
	files_found += 1
	filename_current_original = %A_LoopFileName%
	filename_current = %A_LoopFileName%						; The filename is saved in a variable, as the file name needs to be edited within the renaming process.
	StringTrimRight, filename_body_original, A_LoopFileName, %length_extensionplusperiod_files%

	If (search_inside_text_files = "yes")
	{
		file_altered = 0
		FileDelete, %A_LoopFileDir%\*.%temp_extension%		; Make sure that there are no leftover temp-files. And the default .xtt is really obsolete, so no worries.

		GoSub, text_search_and_replace_engine

		If (file_altered > 0)
		{
			FileMove, %A_LoopFileDir%\ssrs_tempfile.%temp_extension%, %A_LoopFileDir%\%A_LoopFileName%, 1
			If (ErrorLevel = 1)
			{
				FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : Editing of the file " A_LoopFileName " failed. Check user permissions??`n", %dir_files_ssrs%\ssrs_log.txt
				file_edit_failed += 1
				FileDelete, %A_LoopFileDir%\*.%temp_extension%		; Make sure that there are no leftover temp-files.
				file_altered = 0
			}
			else
			{
				FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : There were search hits found inside " A_LoopFileName ", and they were replaced.`n", %dir_files_ssrs%\ssrs_log.txt
				FileDelete, %A_LoopFileDir%\*.%temp_extension%		; Make sure that there are no leftover temp-files.
				file_altered = 0
				files_txt_edited += 1
			}
		}
		else
		{
			FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : Nothing to edit inside " A_LoopFileName ", proceeding...`n", %dir_files_ssrs%\ssrs_log.txt
			FileDelete, %A_LoopFileDir%\*.%temp_extension%		; Make sure that there are no leftover temp-files.
			file_altered = 0	
		}
	}
	
	If (rename_files = "yes")
	{
		GoSub, file_rename_engine		; enter the actual search and replace process.
		
		If (file_rename_failed >= 1)
		{
			If (file_renamed = 1)
			{
				filename_body =
				filename_current =
				FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : File " A_LoopFileName " was renamed, but some rules caused an error, check for duplicate filenames.`n", %dir_files_ssrs%\ssrs_log.txt
				filename_altered = 0
				file_rename_failed = 0
				file_renamed = 0
				continue	; Go to the next file.
			}
			If (file_renamed = 0)
			{
				filename_body =
				filename_current =
				FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : File " A_LoopFileName " was not renamed and there was an error, check for duplicate filenames.`n", %dir_files_ssrs%\ssrs_log.txt
				filename_altered = 0
				file_rename_failed = 0
				file_renamed = 0
				continue			; Go to the next file.
			}
		}
		If (file_renamed = 1)
		{
			files_total_renamed += 1
			filename_body =
			filename_current =
			filename_altered = 0
			file_renamed = 0
			file_rename_failed = 0
			continue	; The renaming engine has already produced the log what happened, just advance to the next file.
		}
		If (file_renamed = 0)
		{
			filename_body =
			filename_current =
			filename_altered = 0
			file_renamed = 0
			file_rename_failed = 0
			FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : Nothing to rename with " A_LoopFileName ", proceeding...`n", %dir_files_ssrs%\ssrs_log.txt
			continue	; Advance to the next file. 
		}
	}
}

End_2:
{
	If (rename_files = "yes")
	{
		FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : " files_found " total files found, " files_total_renamed " were renamed.`n", %dir_files_ssrs%\ssrs_log.txt
	}
	If (search_inside_text_files = "yes")
	{
		FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : " files_found " total files found, " files_txt_edited " (text files) were edited.`n", %dir_files_ssrs%\ssrs_log.txt
	}
	
	FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : End of the script. `n", %dir_files_ssrs%\ssrs_log.txt

	If (logging = "no")
	{
		RunWait, %comspec% /c del /Q "%dir_files_ssrs%\ssrs_log.txt"
	}

	ExitApp				; This should be the main command/line that exits the script.
}


; --------------------------------------------------------------------------------
; --------------------------------------------------------------------------------
; Next part consists of different subscripts. They are located here in the end of
; the script to enhance usability/readablility of this script. Subscripts are
; not intended to exit the script, they should either return to main script
; (with "return") or direct to the end of the main script (End_2: subscript logs
; the end and exits the script).

init_zero_ssrs:
{
	If (dir_files_ssrs = "")
	{
		MsgBox, 0, SETUP ERROR,
		(LTrim
			The script fails to run as dir_files_ssrs (the location
			of the searched files) is not set to a  be valid location,
			and even logging will fail. Thus you will only see
			this error and then the script exits.
		)
		ExitApp			; Unconditional exit.
	}
	IfNotExist, %dir_files_ssrs%
	{
		MsgBox, 0, SETUP ERROR,
		(LTrim
			The script fails to run as dir_files_ssrs (the location
			of the searched files) is not set to a  be valid location,
			and even logging will fail. Thus you will only see
			this error and then the script exits.
		)
		ExitApp			; Unconditional exit.
	}
	
	If (textfile_encoding <> "")
	{
		FileEncoding, %textfile_encoding%		; This should be set here, before any search strings are set. UTF-8 is set by default, so "" is ignored.
	}
	return
}

query_runscript_ssrs:
{
	MsgBox, 1, String search and replace is starting,
	(LTrim
		The SSRS script will start automatically
		in 10 seconds.
		
		You can start the script execution now by selecting "OK".
		
		You can cancel the script execution now by selecting "Cancel".
	), 10
	IfMsgBox Cancel
	{
		FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : Script cancelled by user before initialization.`n", %dir_files_ssrs%\ssrs_log.txt
		GoSub, End_2
	}
	return
}

init_global_ssrs:
{
	FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : Initiating script.`n", %dir_files_ssrs%\ssrs_log.txt

	EnvGet, Env_Path, Path
	SetWorkingDir, %dir_files_ssrs%
	StringLen, length_extension_files, extension_files_ssrs
	length_extensionplusperiod_files := (length_extension_files + 1)
	StringLeft, drive_files, dir_files_ssrs, 3
	
	files_found = 0						; Gives a cleaner "0" for the log, if no files were found.
	files_total_renamed = 0				; Gives a cleaner "0" for the log, if no files were found.
	files_txt_edited = 0				; Gives a cleaner "0" for the log, if no files were found.
	
	If (recursive_ssrs = "yes")
	{
		is_recursive_ssrs = R
	}
	else
	{
		is_recursive_ssrs =
	}
	
	; Check the prerequisites.
	
	IfNotExist, %dir_files_ssrs%		; This was already checked, but once more, just to be sure.
	{	
		FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : The directory for the files can't be found, quitting. `n", %dir_files_ssrs%\ssrs_log.txt
		GoSub, End_2
	}
	FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : Prerequisites were met and the script was initiated, proceeding. `n", %dir_files_ssrs%\ssrs_log.txt
	return
}

text_search_and_replace_engine:
; * The actual engine that searches for strings inside text files and replaces them.
{
	Loop, Read, %A_LoopFileDir%\%A_LoopFileName%			; Starts to read the file in hand line by line (one line for each loop)
	{
		text_line_before = %A_LoopReadLine%

		While, (A_Index <= rulecount_ssrs)		; This loop considers each rule for the current line.
		{
			rule_ssrs_current_index = %A_Index%
	
			If (string_exclusion_rule_ssrs_rule_%rule_ssrs_current_index% <> "")
			{
				IfInString, text_line_before, % string_exclusion_rule_ssrs_rule_%rule_ssrs_current_index%
				{
					text_line_after = %text_line_before%			; This ensures that each line gets a "final" form, even if it was skipped here due to an exclusion rule.
					; (the next line is commented out, as it can produce a lot of logging data and the logic works)
					; FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : The exclusion rule for the search with index " rule_ssrs_current_index " was met on the line " text_line_before ", continuing to the next rule...`n", %dir_files_ssrs%\ssrs_log.txt
					continue	; Continue to the next rule for the same line.
				}
			}
			IfInString, text_line_before, % string_search_ssrs_rule_%rule_ssrs_current_index%
			{
				FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : A match with the search index " rule_ssrs_current_index " was met on the line " text_line_before ", replaced, search for other rules continues...`n", %dir_files_ssrs%\ssrs_log.txt
				StringReplace, text_line_after, text_line_before, % string_search_ssrs_rule_%rule_ssrs_current_index%, % string_replace_ssrs_rule_%rule_ssrs_current_index%, 1
				text_line_before = %text_line_after%			; This makes the resulted string the source for the next rule. Otherwise we might end up with conflicting rules (now the rules are considered in order).
				file_altered += 1
				continue		; Continue to the next rule for the same line.
			}
			else
			{
				text_line_after = %text_line_before%			; This ensures that each line gets a "final" (and trimmed) form, even if nothing was replaced after all the rules are considered.
				continue		; Continue to the next rule for the same line.
			}
			 
		}
		
		; Append the (possibly edited) line after checking for the rules in to a temporary text file.
		FileAppend, % text_line_after "`r`n", %A_LoopFileDir%\ssrs_tempfile.%temp_extension%
		continue		; Continue to the next line to be evaluated for each rule.
	}
	return
}

file_rename_engine:
; * The actual engine that searches for strings in file names and replaces them.
{
	filename_body_old = %filename_body_original%
	filename_altered = 0
	file_renamed = 0
	
	While, (A_Index <= rulecount_ssrs)
	{
		rule_ssrs_current_index = %A_Index%	
		If (string_exclusion_rule_ssrs_rule_%rule_ssrs_current_index% <> "")
		{
			IfInString, filename_current, % string_exclusion_rule_ssrs_rule_%rule_ssrs_current_index%
			{
				; (the next line is commented out, as it can produce a lot of logging data and the logic works)
				; FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : The exclusion rule for a the search with index " rule_ssrs_current_index " was met in file " A_LoopFileName ", skipping to the next search rule.`n", %dir_files_ssrs%\ssrs_log.txt
				continue	; Continue to check the same file for other rules.
			}
		}
		IfInString, filename_current, % string_search_ssrs_rule_%rule_ssrs_current_index%
		{
			StringReplace, filename_body_new, filename_body_old, % string_search_ssrs_rule_%rule_ssrs_current_index%, % string_replace_ssrs_rule_%rule_ssrs_current_index%, 1
			AutoTrim, On
			filename_body_new_trim = %filename_body_new%				; This trims the possible leading or trailing spaces in the file name after the replacement.
			AutoTrim, Off
			filename_body_old = %filename_body_new_trim%								; This makes the altered filename body the new source for editing.
			filename_current = %filename_body_new_trim%.%extension_files_ssrs%			; This makes the altered filename the new source for following rules (!). Order of the rules matter.
			If (filename_current <> A_LoopFileName)
			{
				FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : A search match (index " rule_ssrs_current_index ") was met in filename " A_LoopFileName ", replaced, search for other rules continues... `n", %dir_files_ssrs%\ssrs_log.txt
				filename_altered += 1
			}
		}
	}

	If (filename_altered > 0)
	{
		FileMove, %A_LoopFileDir%\%filename_body_original%.%extension_files_ssrs%, %A_LoopFileDir%\%filename_body_new_trim%.%extension_files_ssrs%
		If (ErrorLevel = 1)
		{
			FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : Renaming of file " A_LoopFileName " failed, possibly due to pre-existing file with the new filename? `n", %dir_files_ssrs%\ssrs_log.txt
			file_rename_failed += 1
			return
		}
		else
		{
			FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : File " A_LoopFileName " was renamed to " filename_current ". `n", %dir_files_ssrs%\ssrs_log.txt
			file_renamed = 1
			return
		}
	}
	return
}

; The search initiating below is very central part of the whole script and also very delicate.
; You should not touch the code, unless you are implementing a new features to the script.

init_rule:
{
	If (string_search_ssrs <> "")
	{
		rulecount_ssrs += 1
		rule_tag = rule_%rulecount_ssrs%
		string_search_ssrs_%rule_tag% = %string_search_ssrs%
		string_replace_ssrs_%rule_tag% = %string_replace_ssrs%
		string_exclusion_rule_ssrs_%rule_tag% = %string_exclusion_rule_ssrs%
		; Next, clear the non-tagged variables after tagging
		string_search_ssrs =
		string_replace_ssrs =
		string_exclusion_rule_ssrs =
		ruletag =
		; And return to the next rule / continue with the script.
		return
	}
	else
	{
		FileAppend, % A_DD "/" A_MM "/" A_YYYY " " A_Hour ":" A_Min ":" A_Sec " : Empty search string set, skipping. `n", %dir_files_ssrs%\ssrs_log.txt
		string_search_ssrs =
		string_replace_ssrs =
		string_exclusion_rule_ssrs =
		ruletag =
		; Skip to the next rule / continue with the script.
		return
	}
}
