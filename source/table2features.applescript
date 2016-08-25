--------------------------------------------------------------------------
-- Description  : io functionality for MAC Excel 2016 makro table2features
--------------------------------------------------------------------------

-- Copyright 2016 Matthias Carell
--
--   Licensed under the Apache License, Version 2.0 (the "License");
--   you may not use this file except in compliance with the License.
--   You may obtain a copy of the License at
--
--       http://www.apache.org/licenses/LICENSE-2.0
--
--   Unless required by applicable law or agreed to in writing, software
--   distributed under the License is distributed on an "AS IS" BASIS,
--   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
--   See the License for the specific language governing permissions and
--   limitations under the License.

-----------------------------------------------------------------------------------------
-- description: create a new feature file
-- parameters:		pFeatureFile		- combines full path for the new feature file and content
--													because Excel accepts only one parameter
-- return value: 
-----------------------------------------------------------------------------------------
on writeFeatureToFile(pFeatureFile)
	
	local vFileRef, vFeatureFile, vFeatureFilePath, vFeatureFileContent, vContentLine, vErrDialogResult, vActionOnError
	
	set vFeatureFilePath to first paragraph of pFeatureFile
	set vFeatureFileContent to paragraphs 2 thru (count paragraphs of pFeatureFile) of pFeatureFile
	
	set vFeatureFile to a reference to file vFeatureFilePath
	try
		set vFileRef to (open for access vFeatureFile with write permission)
	on error errMsg number errNum
		set vErrDialogResult to display dialog ("Open for Access, Error Number: " & errNum as string) & return & errMsg buttons {"cancel", "continue"} default button "continue" with icon caution
		if button returned of vErrDialogResult is "cancel" then
			return "cancel"
		end if
	end try
	
	try
		repeat with vContentLine in vFeatureFileContent
			write vContentLine & linefeed to vFileRef as «class utf8»
		end repeat
	on error errMsg number errNum
		set vErrDialogResult to display dialog ("Write, Error Number: " & errNum as string) & return & errMsg buttons {"cancel", "continue"} default button "continue" with icon caution
		if button returned of vErrDialogResult is "cancel" then
			return "cancel"
		end if
	end try
	
	try
		close access vFileRef
	on error errMsg number errNum
		log ("Close, Error Number: " & errNum as string) & return & errMsg
	end try
end writeFeatureToFile


-----------------------------------------------------------------------------------------
-- description: ask user where to expect the .feature files
-- parameters:		pDummy		- it seems that Excel 2016 expect all funtions to have exact one parameter
-- return value: has to be a string
-----------------------------------------------------------------------------------------
on chooseFeatureFolder(pDummy)
	try
		tell application "Finder"
			application activate
			set vPath to (choose folder with prompt "choose feature folder" default location (path to the desktop folder from user domain)) as string
			--return URL of vPath & "#@#@" & displayed name of disk of vPath
			return vPath
		end tell
	on error
		return ""
	end try
end chooseFeatureFolder
