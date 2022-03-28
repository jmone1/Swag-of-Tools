SwagOfToolVer = 29 March 2022
; Author: jmone Thread on Interact: http://yabb.jriver.com/interact/index.php/topic,106802.0.html

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance ignore
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#MaxMem 4095

SetTitleMatchMode, 2 ; As long as it's not 3 it should work, if it's 3 you'd need to alter the title used in the WinGet line below.
DetectHiddenWindows, On
WinGet, BottomMost, IDLast, %A_ScriptFullPath%
If (A_ScriptHwnd != BottomMost)
  ExitApp

;============= Set Vars ===========================
PWKey := UUID()
IniRead, Hashed_UserName, Swag of Tools.ini,MC_Settings,MC_UserName
IniRead, Hashed_Password, Swag of Tools.ini,MC_Settings,MC_Password

Loop, Parse, Hashed_UserName
	MC_UserName .= Chr(((Asc(A_LoopField)-32)^(Asc(SubStr(PWKey,A_Index,1))-32))+32)
Loop, Parse, Hashed_Password
	MC_Password .= Chr(((Asc(A_LoopField)-32)^(Asc(SubStr(PWKey,A_Index,1))-32))+32)

IniRead, MC_Ver, Swag of Tools.ini,MC_Settings,MC_Ver
IniRead, MC_WS, Swag of Tools.ini,MC_Settings,MC_WS
IniRead, MCFieldTarget, Swag of Tools.ini,MediaInfo_Settings,MCFieldTarget,%A_Space%
IniRead, MediaInfoField, Swag of Tools.ini,MediaInfo_Settings,MediaInfoField,%A_Space%
IniRead, MediaInfoSection, Swag of Tools.ini,MediaInfo_Settings,MediaInfoSection, %A_Space%
IniRead, NoMediaInfoFile, Swag of Tools.ini,MediaInfo_Settings,NoMediaInfoFile,%A_Space%
IniRead, SkipMediaInfoFile, Swag of Tools.ini,MediaInfo_Settings,SkipMediaInfoFile,%A_Space%
IniRead, CropScanPoint, Swag of Tools.ini,MediaInfo_Settings,CropScanPoint,%A_Space%
IniRead, CropSettings, Swag of Tools.ini,MediaInfo_Settings,CropSettings,%A_Space%
EOLChar = {EOL}
CRChar = {CR}
LFChar = {LF}
TabChar = {TAB}
FileEncoding, UTF-8

;============= Initialise and Setup Menu===========================
if !FileExist("Swag of Tools.ini")
  {
  MC_Ver = 26
  MC_WS = 127.0.0.1:52199
  MC_Username = username
  MC_Password = password
  msgbox, You must enter details into the "Configuration" Option before using
  }

;============= Main Menu ================================
Gui Add, TreeView, gShowOptions x5 y12 w165 h260
Gui Add, Button,x5 y310 cBlue gButtonClearLog, X
Gui Add, Text,x35 y314, Clear Log of Events
Gui Add, Checkbox, x170 y314 vVerboseLogging, Turn on Verbose Logging of MCWS Calls
Gui Add, Edit, x5 y335 w595 h95 vLog
Gui Add, Button, x5 y280 w167 h23 Default, &Cancel

;=============  Run Time Menu from MC ================================
InputVar = %1%
If InputVar
{
  Gui Add, Text,x15 y15, You are in MC Run Time Mode
  Gui Add, GroupBox, x179 y7 w420 h300, About Swag of Tools - MC Run Time (Launched Directly from MC)
  Gui Add, Text,x200 y50, Version : %SwagOfToolVer%
  Gui Add, Link,, Interact Forum : <a href="Swag of Tools Thread ">http://yabb.jriver.com/interact/index.php/topic,106802.0.html</a>
  Gui Add, Text,, Some features of the Tool Box requires MC Version 22.0.21 or later including`n- Ability to Add the Chapterfy "chapter DB #" Field, and `n- DataFiddler Adding new Items from TXT File 
  MC_Call = http://%MC_WS%/MCWS/v1/Alive
  GoSub, MCWS
  If Status = 200 
    {
	RegExMatch(Result,"(?<=ProgramVersion"">)(.*)(?=</Item)", MC_Version)
	MsgLog = Media Center Version %MC_Version% Found at %MC_WS%`n%MsgLog%
    GuiControl,,Log, %MsgLog%
	}
  Gui Show, w610 h440, Swag of Tools
  MsgLog = The MC Runtime Command Received was : %InputVar%`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  If InputVar = ReadChapterFile
    gosub, ButtonReadChapterFile
  If InputVar = CreateChapterFile
    gosub, ButtonCreateChapterFile
  If InputVar = Burn2BD
    gosub, ButtonBurn2BD
  If InputVar = MediaInfo
    gosub, ButtonMediaInfo
  If InputVar = FilenameUpdater
    gosub, ButtonUpdateFilename
	
  MsgLog = The MC Run Time Command has finished, you can now either press`n - "Cancel" to close the Swag of Tools or`n - "Launch Full Version"`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  Gui Add,Button,x300 y250 gButtonFullVer, Launch Full Version of MC ToolBox
  Return
  ButtonFullVer:
    Reload
}

;=============  Standalone Menu ================================
TV_Add("About")
Chapterfy := TV_Add("Chapterfy",,"Expand")
    TV_Add("Manually Create", Chapterfy)
    TV_Add("ChapterDB Lookup", Chapterfy)
    TV_Add("Read Blu-ray MPLS", Chapterfy)
	TV_Add("chapter.txt File", Chapterfy)
TV_Add("Particlise")
TV_Add("Data Fiddler")
TV_Add("Media Checker")
TV_Add("Burn2BD")
TV_Add("MediaInfo")
TV_Add("Filename Updater")
TV_Add("Configuration")

Gui Add, Tab2, vTab x7 y500 w588 h22 -Wrap +Theme, About|Chapterfy|Manually Create|ChapterDB Lookup|Read Blu-ray MPLS|chapter.txt File|Particlise|Data Fiddler|Media Checker|Burn2BD|MediaInfo|Filename Updater|Configuration
Gui Show, w610 h440, Swag of Tools

; About
  Gui Tab, 1
  Gui Add, GroupBox, x179 y7 w420 h300, About Swag of Tools
  Gui Add, Text,x200 y50, Version : %SwagOfToolVer%
  Gui Add, Link,, Swag of Tools <a href="http://yabb.jriver.com/interact/index.php/topic,106802.0.html">Thread on Interact</a>
  Gui Add, Text,, Some features of the Tool Box requires MC Version 22.0.21 or later including`n- Ability to Add the Chapterfy "chapter DB #" Field, and `n- DataFiddler Adding new Items from TXT File 
  MC_Call = http://%MC_WS%/MCWS/v1/Alive
  GoSub, MCWS
  If Status = 200 
    {
	RegExMatch(Result,"(?<=ProgramVersion"">)(.*)(?=</Item)", MC_Version)
	MsgLog = Media Center Version %MC_Version% Found at %MC_WS%`n%MsgLog%
    GuiControl,,Log, %MsgLog%
	}

; Chapterfy
  Gui Tab, 2
  Gui Add, GroupBox, x179 y7 w420 h300, Chapterfy
  Gui, Add, Text,x189 y30, Chapterfy : Creates particles for Each Chapter of a file.
  Gui, Add, Text,,This is particular useful to create a list of Tracks for Music Videos`njust like you would from a CD
  Gui, Add, Text,,There are several options to Manually and Automatically `n,create and backup the Chapter based Particles
  Gui, Add, Text,,Select the appropriate option and follow the instuctions

; Chapterfy Manually Create
  Gui Tab, 3
  Gui Add, GroupBox, x179 y7 w420 h300, Chapterfy Manually Create
  Gui, Add, Text,x189 y30, Chapterfy (Bulk Create) : Select any ONE Item in MC to`n Manually Create Chapter Based Particles
  Gui, Add, Text,, You will be presented with a Chapter List where you can:`n- Enter Chapter Names`n- Enter Start/End Times
  Gui, Add, Button,x280 y125 gButtonManual, Create these # of Chapter Particles
  Gui, Add, Edit, x250 y125 w25 vManualNum

; Chapterfy ChapterDB Lookup
  Gui Tab, 4
  Gui Add, GroupBox, x179 y7 w420 h300, Chapterfy ChapterDB (Lookup from Plex Archive) 
  Gui, Add, Text,x189 y30, While the ChapterDB site has been decommissioned,there is still a Read Only Archive `nat Plex for chatper information up to that point that can still be accessed `nBrowse the Archive and note the # at the end of the URL for that chapter file`nsuch as "18" in https://chapterdb.plex.tv/browse/18 
  Gui, Add, Link,x189 y90, Click the Link to Browse : <a href="https://chapterdb.plex.tv/browse">ChapterDB</a> 
  Gui, Add, Text,x189 y120, Select any ONE item in MC to`ncreate Chapter Particles based on a specific ChapterDB #
  Gui, Add, Button,x280 y155 gButtonChapterDBNum, Enter the ChapterDB # to use (Any Item in MC)
  Gui, Add, Edit, x225 y155 w50 vAutoNum


; Chapterfy Read Blu-ray MPLS
  Gui Tab, 5
  Gui Add, GroupBox, x179 y7 w420 h300, Chapterfy Read Blu-ray MPLS
  Gui, Add, Text,x189 y30, Chapterfy (MPLS) : Select any ONE Bluray Blu Ray Playlist (MPLS) in MC to`n create Chapter Particles (No Chapter Names)
  Gui, Add, Text,, You may need to refresh and expand the stack of the selected items in MC`n once they have been created
  Gui, Add, Button, x280 y125 gButtonReadMPLS, Create Chapters (MPLS)
  
; Chapterfy chapter.txt Sidecar Files
  Gui Tab, 6
  Gui Add, GroupBox, x179 y7 w420 h300, Chapterfy chapter.txt Sidecar File(s)
  Gui, Add, Text,x189 y30, Create from chapter.txt : Create chapters from a Chapter.TXT (not XML) File
  Gui, Add, Button,x280 y60 gButtonReadChapterFile, Create Chapters
  Gui, Add, Text,x189 y100, Backup to chapter.txt : Select ALL Related chapters in MC
  Gui, Add, Button,x280 y125 gButtonCreateChapterFile, Create a sidecar Chapter.TXT File

; Particlise
  Gui Tab, 7
  Gui Add, GroupBox, x179 y7 w420 h300, Particlise
  Gui, Add, Text,x189 y30, Particles (Bulk Create) : Select any # of Items in MC to Bulk Create Particles
  Gui, Add, Text,, You may need to refresh and expand the stack of the selected items in MC`n once they have been created
  Gui, Add, Button, x280 y125 gButtonParticlise, Create these # of Particles for each Selected Item
  Gui, Add, Edit, x250 y125 w25 vBulkNum

; Data Fiddler
  Gui Tab, 8
  Gui Add, GroupBox, x179 y7 w420 h300, Data Fiddler
  Gui, Add, Text,x189 y30, Data Fiddler : Import / Export / Update the MC Library
  Gui, Add, Button, x220 y50 gButtonReadInfoAll, Export the Entire MC Library (Creates a TXT File)
  Gui, Add, Button,gButtonReadInfo, Export Selected Items Only from the MC Library (Creates a TXT File)
  Gui, Add, Button,gButtonOpenMPL, Convert a MC MPL Playlist File (Creates TXT File)
  Gui, Add, Button,gButtonWriteInfo, Update MC Library (From TXT File)
  Gui, Add, Button,gButtonWriteMPL, Create/Import New Items using a MC Playlist File (From TXT File)

; Media Checker
  Gui Tab, 9
  Gui Add, GroupBox, x179 y7 w420 h300, Media Checker
  Gui, Add, Text,x189 y30, Checks to see if items in a Folder are also in the MC Library
  Gui, Add, Text,, Currently this requires an exact match so it will report as missing BD/DVD Files etc
  Gui, Add, Text,, The Results will be written to a TXT file in the Swag of Tools Directory
  Gui, Add, Button, x300 y125 gButtonMediaChecker, Open Folder to Check

; Burn2BD
  Gui Tab, 10
  Gui Add, GroupBox, x179 y7 w420 h300, Burn2BD
  Gui, Add, Text,x189 y30, Select One Video in MC to Burn to a Blu-ray Structure / Disk and Burn2BD will:
  Gui, Add, Text,,- Transcode the Selected Video to a 1920x1080 MPEG2 Video / DD Audio file
  Gui, Add, Text,,- Run tsMuxeR to create a Blu-Ray ISO
  Gui, Add, Text,,- Run Windows ISO Burn so you can (optionally) create a Blu-ray Disk
  Gui, Add, Button, x300 y150 gButtonBurn2BD, Burn 2 BD

; MediaInfo
  Gui Tab, 11
  Gui Add, GroupBox, x179 y7 w420 h300, MediaInfo and FFMpeg
  Gui, Add, Text,x189 y30, Select One or more items to run MediaInfo / FFMpeg on and it will: `n- Adds the MediaInfo output as an "Extra" for the selected file(s) `n- Saves Atmos / DTS:X Codec information to the Compression Field `n- Optionally add info from MediaInfo to a Media Center Field
  Gui, Add, Edit, x220 y85 w130 h20 vMCFieldTarget, %MCFieldTarget%
  Gui, Add, Text,x360 y85,<--- Enter the Field in MC to Update
  Gui, Add, Edit, x220 y110 w130 h20 vMediaInfoField, %MediaInfoField%
  Gui, Add, Text,x360 y110,<--- Enter MediaInfo Data Field to collect
  Gui, Add, Edit, x220 y135 w130 h20 vMediaInfoSection, %MediaInfoSection%
  Gui, Add, Text,x360 y135,<--- Enter MediaInfo Section to restrict to
  Gui, Add, Checkbox, x200 y160 vNoMediaInfoFile Checked%NoMediaInfoFile%, Don't save MediaInfo output as an "Extra" File
  Gui, Add, Checkbox, x200 y185 vSkipMediaInfoFile Checked%SkipMediaInfoFile%, Use existing MediaInfo "Extra" File if it already exists
  Gui, Add, Link,x189 y210, Crop detect settings <a href="https://yabb.jriver.com/interact/index.php/topic,106802.msg912759.html#msg912759">More Crop Detect Information</a>




  Gui, Add, Edit, x220 y225 w80 h20 vCropScanPoint, %CropScanPoint%
  Gui, Add, Text,x310 y225,<--- Enter Secs into video to measure Crop Factor`n         (leave blank or 0 to disable)
  Gui, Add, Edit, x220 y255 w80 h20 vCropSettings, %CropSettings%
  Gui, Add, Text,x310 y255,<--- Enter Custom Settings or leave blank for default`n         (eg 24:16:2 or 0.1:16:2 etc)
  Gui, Add, Button, x500 y280 gButtonMediaInfo, Get MediaInfo  

; FilenameUpdater
  Gui Tab, 12
  Gui Add, GroupBox, x179 y7 w420 h300, Filename Updater
  Gui, Add, Text,x189 y30, Filename Updater : Select any # of Items in MC to change the Filename field
  Gui, Add, Text,, Main use: To update miscreated Particle Filenames that can not `nbe changed in MC's GUI
  Gui, Add, Button, x280 y125 gButtonUpdateFilename, Update Filename(s)


; Configuration
  Gui Tab, 13
  Gui Add, GroupBox, x179 y7 w420 h300, Setup and Configure
  Gui Add, Text, x189 y30, You may need to "Run as Administrator" for these Configuration changes to be made
  Gui Add, Text, x189 y50, In MC--> Tools--> Option--> Media Nework, Check the; `n- "Use Media Network to share this library and enable DLNA" and `n- "Authentication" 
  Gui, Add, Edit, x220 y100 w130 h20 vMC_Ver, %MC_Ver%
  Gui, Add, Text,x360 y100,<--- Enter the Version of MC (eg 20, 21, 22 etc)
  Gui, Add, Edit, x220 y130 w130 h20 vMC_WS, %MC_WS%
  Gui, Add, Text,x360 y130,<--- Enter MCWS Server Address:Port
  Gui, Add, Edit, x220 y160 w130 h20 vMC_UserName Password, %MC_Username%
  Gui, Add, Text,x360 y160,<--- Enter MCWS Username
  Gui, Add, Edit, x220 y190 w130 h20 vMC_Password Password, %MC_Password%
  Gui, Add, Text,x360 y190,<--- Enter MCWS Password
  Gui, Add, Button, x220 y220 w130 h20 gButtonTestMCWS, Test Configuratioin
  Gui, Add, Button, x220 y250 w130 h20 gButtonSaveMCWS, Save Configuration

ShowOptions:
    TV_GetText(OutputVar, A_EventInfo)
    GuiControl ChooseString, SysTabControl321, % OutputVar
	
Return

;============= Process Choice =============
ButtonClearLog:
  Gui, Submit, NoHide
  guiControlGet, UserInput
  MsgLog = 
  GuiControl,,Log, %MsgLog%
return

ButtonUpdateFilename:
  MC_Call = http://%MC_WS%/MCWS/v1/Files/Current?Action=Serialize
  GoSub, MCWS
  If Status !=200
    return
  MsgLog = Commenced - Please wait....`n%MsgLog%  
  GuiControl,,Log, %MsgLog%
  Loop, parse, Result, `;
	If A_Index > 3
	{
	  MC_Key = %A_LoopField%
	  gosub, MC_GetInfo
      If Status !=200
        return
	  Else
      {
		InputBox, UserInput, Update Filename, Update Filename,,600,100,,,,,%MC_FilenameExt%
		if ErrorLevel
		{
			MsgLog = Operation Cancelled`n%MsgLog%
			GuiControl,,Log, %MsgLog%
			return
		}
		else
		{
			MC_Call = http://%MC_WS%/MCWS/v1/File/SetInfo?File=%MC_Key%&FileType=Key&Field=Filename&Value=%UserInput%
			GoSub, MCWS
			If Status !=200
				return
			Else
			{
				MsgLog = MC %MC_FilenameExt% changed to %UserInput%`n%MsgLog%
				GuiControl,,Log, %MsgLog%
			}
		}
	  }
	}
  MsgLog = Finished`n%MsgLog%  
  GuiControl,,Log, %MsgLog%
return

ButtonReadChapterFile:
  FileSelectFile, SelectedFile, 3, , Open a file
  GuiControl,,Select%A_GuiControl%,%SelectedFile%
  If (SelectedFile = "")
  {
    MsgLog = Operation Cancelled`n%MsgLog%
    GuiControl,,Log, %MsgLog%
    return
  }
  FileCopy, %SelectedFile%, temp.chapters.txt, 1
  gosub, MC_Current
  MsgLog = Using Chapter File %SelectedFile% to Create Chapters for MC File Key %Result%`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  gosub, Process_Chapters
return

ButtonManual:
  guiControlGet, ManualNum
  If ManualNum < 1
    {
	 MsgLog = You must specify 1 or more Chapters to Create`n%MsgLog%
     GuiControl,,Log, %MsgLog%
	 Return
	}
  Loop %ManualNum%
  {
  Chap_N = %A_Index%
  If A_Index < 10
    Chap_N = 0%A_Index%
  FileAppend, CHAPTER%Chap_N%=00:%Chap_N%:00.000`n, temp.chapters.txt
  FileAppend, CHAPTER%Chap_N%NAME=Chapter %Chap_N%`n, temp.chapters.txt
  }
  gosub, MC_Current
  MsgLog = Created Chapter File with %Chap_N% Chapters for MC File Key %Result%`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  gosub, Process_Chapters
return

ButtonParticlise:
  guiControlGet, BulkNum
  If BulkNum < 1
    {
	 MsgLog = You must select at least 1 Particle to Create`n%MsgLog%
     GuiControl,,Log, %MsgLog%
	 Return
	}
  MC_Call = http://%MC_WS%/MCWS/v1/Files/Current?Action=Serialize
  GoSub, MCWS
  If Status !=200
    return
  RegExMatch(result,"[^;]+$", MC_Key)
  If MC_Key < 1
    return
  Loop, parse, Result, `;
    If A_Index > 3
    {
      MC_Call = http://%MC_WS%/MCWS/v1/File/CreateParticle?File=%A_LoopField%&FileType=Key&Count=%BulkNum%
      GoSub, MCWS
      If Status !=200
        return
	  Else
        {
		 MsgLog = Created %BulkNum% Particles for MC File Key %A_LoopField%`n%MsgLog%
         GuiControl,,Log, %MsgLog%
        }
    }
  MsgLog = Finished`n%MsgLog%
  GuiControl,,Log, %MsgLog%
return

ButtonChapterDBNum:
  guiControlGet, AutoNum
  If AutoNum > 0   ; Get by chapterDB # from Plex backup
  {
  	MsgLog = Checking for chapterDB #%AutoNum%, Please Wait
    GuiControl,,Log, %MsgLog%
    urldownloadtofile, https://chapterdb.plex.tv/browse/%AutoNum%.txt, temp.chapters.txt
    gosub, MC_Current
    gosub, Process_Chapters
  }
  Else
  {
    MsgLog = You need to enter a Chapter DB #
    GuiControl,,Log, %MsgLog%
  }
return

ButtonReadMPLS:
  guiControlGet, UserInput
  gosub, MC_Current
  gosub, MC_GetInfo
  If Status !=200
    return
  If MC_FileExt = mpls ;or MC_FileType = MPLS
    {
  	 MsgLog = Checking chapter Info for %MC_FilenameExt%`n%MsgLog%
     GuiControl,,Log, %MsgLog%

;-------Open File and Read as HEX to Var Hexfile-----
gosub, FileToHex

;-------Parse Main Section of MPLS------------------- 
PlaylistSectionA := % hexToDecimal(SubStr(Hexfile,17,8)) ; Location of the Playlist Section
PlaylistMarkSectionA:= % hexToDecimal(SubStr(Hexfile,25,8)) ; Location of the Playlist Mark Section 

;-------Parse Playlist Section ---------------------- 
PlaylistSectionA := PlaylistSectionA*2+1
TempHex := SubStr(HexFile,PlaylistSectionA+12,4)
NumberofPlayItems := % hexToDecimal(TempHex) ; Get number of Playlist Items

;-------Loop PlayItem in Playlist Section------------
PICumulativeDuration = 0 
PITimeOutPrevious = 0
PlayItemA := PlaylistSectionA+20
Loop %NumberofPlayItems% ; Gets the detail for each PlayItem
	{
	TempHex := SubStr(HexFile,PlayItemA,4)
	PlayItemLength := % hexToDecimal(TempHex)

	TempHex := SubStr(HexFile,PlayItemA+28,8)
	Temp := % hexToDecimal(TempHex)
	PITimeIn := Temp/45000 ; PlayItem's Time In value

	TempHex := SubStr(HexFile,PlayItemA+36,8)
	Temp := % hexToDecimal(TempHex)
	PITimeOut := Temp/45000 ; PlayItem's Time Out Value

	PICumulativeDuration%A_Index% := PICumulativeDuration + PITimeIn - PITimeOutPrevious ; Calculate each Play Items Offset
	PICumulativeDuration := PICumulativeDuration%A_Index%

	PlayItemA := PlayItemA+PlayItemLength*2+4
	PITimeOutPrevious = %PITimeOut%
	}

;-------Parse Playlist Mark Section -------------------- 
PlaylistMarkSectionA := PlaylistMarkSectionA*2+1
TempHex := SubStr(HexFile,PlaylistMarkSectionA+8,4)
NumberofPlaylistMarks := % hexToDecimal(TempHex) ; Gets Number of Play List Marks (chapters)

;-------Parse Playlist Mark Entry and Write Chapter File ----------------------- 
Chap_N := 1
PlaylistMarkSectionEntryA := PlaylistMarkSectionA+12
Loop %NumberofPlaylistMarks% ; Gets the details for each chapter and calculates the start time
	{
	TempHex := SubStr(HexFile,PlaylistMarkSectionEntryA+4,4) ; Get PlayItem ID (starts at 0)
	PlayItemID := % hexToDecimal(TempHex) ; Get PlayItem ID (starts at 0)
	PlayItemID := PlayItemID + 1 ; Add 1 to PlayItem ID as it starts from 0 instead of 1 to match the PlayItem ID Array 

	TempHex := SubStr(HexFile,PlaylistMarkSectionEntryA+8,8) ; Time offset associated with this chapter mark.
	Temp := % hexToDecimal(TempHex) ; Time offset associated with this chapter mark.
	TimeOffset := Temp/45000 ; Time offset associated with this chapter mark.

	TempHex := SubStr(HexFile,PlaylistMarkSectionEntryA+2,2) ; Check to see if it is a Entry Mark (1) or Link Point (2) 
	MarkType := % hexToDecimal(TempHex) ; Time offset associated with this chapter mark.
	
	ChapterStartTimeSec%A_Index% := TimeOffset - PICumulativeDuration%PlayItemID% ; ChapterTime for change of PlayItemID
	ChapterStartTime = % SecToChapTime(ChapterStartTimeSec%A_Index%)

	If MarkType != 2 ; ignore Link Points (which are #2) and write the contents to the file
		{
		FileAppend, CHAPTER%Chap_N%=%ChapterStartTime%`n, temp.chapters.txt 
		FileAppend, CHAPTER%Chap_N%NAME=Chapter %Chap_N%`n, temp.chapters.txt
		Chap_N := Chap_N + 1
		}
	PlaylistMarkSectionEntryA := PlaylistMarkSectionEntryA+28
	}

    gosub, Process_Chapters
    }
  Else
    {
  	 MsgLog = %MC_FilenameExt% is not a Bluray Playlist (*.mpls) file!`n%MsgLog%
     GuiControl,,Log, %MsgLog%
	}
return

ButtonCreateChapterFile:
  GoSub, MC_Current
  If MC_Chapters0 < 1
    return
  MsgLog = Creating Chapter File Please Wait`n%MsgLog%
  GuiControl,,Log, %MsgLog%
    
  Loop, %MC_Chapters0%
  {
    MC_Key := MC_Chapters%A_Index%
    GoSub, MC_GetInfo
    ;StringReplace, MC_Name, MC_Name, &amp;, &  This call has been depricated need to test following
	MC_Name := StrReplace(MC_Name,"&amp;","&")
    Chap_N = %A_Index%
    If A_Index < 10
      Chap_N = 0%A_Index%
	  
	ChapFilename=%MC_Filename%.chapters.txt
    If A_Index = 1
      FileDelete, %ChapFilename%
	  
    FileAppend, CHAPTER%Chap_N%=%MC_Start%`n, %ChapFilename%
    FileAppend, CHAPTER%Chap_N%NAME=%MC_Name%`n, %ChapFilename%
  }
  
  MsgLog = Completed`n%MsgLog%
  GuiControl,,Log, %MsgLog%
return

ButtonReadInfoAll:
  MPLFileName = MC_Library_%A_Now%.mpl
  MsgLog = Downloading Library to %MPLFileName%`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  sleep, 100
  URLDownloadToFile, http://%MC_UserName%:%MC_Password%@%MC_WS%/MCWS/v1/Files/Search?Action=mpl&ActiveFile=-1&Zone=-1&ZoneType=ID, %MPLFileName%
  OutputFileName = MC_Output_%A_Now%.txt
  gosub, Process_MPL
return

ButtonReadInfo:
  MPLFileName = MC_Library_%A_Now%.mpl
  MsgLog = Downloading Library to %MPLFileName%`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  sleep, 100
  URLDownloadToFile, http://%MC_UserName%:%MC_Password%@%MC_WS%/MCWS/v1/Files/Current?Action=mpl&ActiveFile=-1&Zone=-1&ZoneType=ID, %MPLFileName%
  OutputFileName = MC_Output_%A_Now%.txt
  gosub, Process_MPL
return

ButtonOpenMPL:
  FileSelectFile, SelectedFile, 3, , Open a file, *.mpl
  GuiControl,,Select%A_GuiControl%,%SelectedFile%
  If (SelectedFile = "")
  {
    MsgLog = Operation Cancelled`n%MsgLog%
    GuiControl,,Log, %MsgLog%
    return
  }
  SplitPath, SelectedFile,,,,OutputFileName
  OutputFileName = %OutputFileName%.txt
  gosub, Process_MPL
return

ButtonWriteInfo:
  FileSelectFile, SelectedFile, 3, , Open a file, *.txt
  GuiControl,,Select%A_GuiControl%,%SelectedFile%
  If (SelectedFile = "")
  {
    MsgLog = Operation Cancelled`n%MsgLog%
    GuiControl,,Log, %MsgLog%
    return
  }
  Fileread,result, %SelectedFile%
  gosub, Update_MC
return

ButtonWriteMPL:
  FileSelectFile, SelectedFile, 3, , Open a file, *.txt
  GuiControl,,Select%A_GuiControl%,%SelectedFile%
  If (SelectedFile = "")
  {
    MsgLog = Operation Cancelled`n%MsgLog%
    GuiControl,,Log, %MsgLog%
    return
  }
  Fileread,result, %SelectedFile%
  SplitPath, SelectedFile,,,,OutputFileName
  OutputFileName = %OutputFileName%_%A_Now%.mpl
  FileDelete, %OutputFileName%
  gosub, Process_TXT

  msgbox, 4,, Do you want to Import this Playlist into MC?
  IfMsgBox Yes
  {
    MC_Call = http://%MC_WS%/MCWS/v1/Library/Import?Path=%A_ScriptDir%\%OutputFileName%
    gosub, MCWS
  }
return

ButtonMediaChecker:
  FileSelectFolder, Folder
  If Folder =
    return
  MissingFileList = MissingFileList%A_now%.txt
  Fileappend, Run Started - %A_Now%, %FileList%
  FileList =
  Loop Files, %Folder%\*.*, R  ; Recurse into subfolders.
  {
    TempData = %A_LoopFileFullPath%
    GuiControl,,Log, Scanning File....`n%A_LoopFileFullPath%`n%MsgLog%
    TempData = % UriEncode(TempData)
    MC_Call = http://%MC_WS%/MCWS/v1/File/GetInfo?File=%TempData%&FileType=Filename
    GoSub, MCWS
    If Status = 500 
      fileappend, %A_LoopFileFullPath%`n`r, %MissingFileList%
  }
  MsgLog = Finished Scanning Files.`nResults are in "%MissingFileList%"`n`n%MsgLog%
  GuiControl,,Log, %MsgLog%
return

ButtonBurn2BD:
  Gui, Submit, NoHide
  guiControlGet, UserInput
  IfNotExist, %A_ScriptDir%\tsMuxeR\tsMuxeR.exe
  {
     MsgBox, 1,, tsMuxeR is missing.  Download Now or Cancel?
	 IfMsgBox OK
	 {
	   MC_Call = https://drive.google.com/uc?export=download&id=0B0VmPcEZTp8NWWUtdkZ1M2h0dFE
       Source = %A_ScriptDir%\tsmuxer.zip
       Dest = %A_ScriptDir%\tsMuxeR
       FileCreateDir, %A_ScriptDir%\tsMuxeR
	   DownloadFile(MC_Call,Source, True, True)
       Unz(Source, Dest)
       FileDelete, %A_ScriptDir%\tsmuxer.zip
	   FileDelete, %A_ScriptDir%\tsMuxeR\MCBurn2BD.meta
       FileAppend,
(
MUXOPT --no-pcr-on-video-pid --new-audio-pes --blu-ray --label="FromMC" --vbr  --auto-chapters=5 --vbv-len=500
V_MPEG-2, "%A_ScriptDir%\temp.ts", track=256
A_AC3, "%A_ScriptDir%\temp.ts",  timeshift=-4ms, track=257
), %A_ScriptDir%\tsMuxeR\MCBurn2BD.meta, UTF-8-RAW
     }
	 Else
	   return
  }

  FileDelete, %A_ScriptDir%\temp.iso
  FileDelete, %A_ScriptDir%\temp.ts
  gosub, MC_Current
  gosub, MC_GetInfo
  gosub, MC_GetToken
  If Status !=200
    return
  MsgLog = Preparing File, Please Wait`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  MC_Call = http://%MC_WS%/MCWS/v1/File/GetFile?File=%MC_Key%&FileType=Key&Conversion=47&Playback=0&Token=%MC_Token%

  TempFile = %A_ScriptDir%\temp.ts
  DownloadFile(MC_Call, TempFile, True, True)

  MsgLog = Making BD Image, Please Wait`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  runwait, %comspec% /c %A_ScriptDir%\tsMuxeR\tsMuxeR.exe %A_ScriptDir%\tsMuxeR\MCBurn2BD.meta %A_ScriptDir%\temp.iso >> log.txt,, Hide
  Fileread, logtxt, log.txt
  Filedelete, log.txt
  MsgLog = Launching Windows ISO Burn to Disk Application`n-----tsMuxerLog-----`n%logtxt%`n----------`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  run, isoburn.exe %A_ScriptDir%\temp.iso

return

ButtonMediaInfo:
  Gui, Submit, NoHide
  guiControlGet, UserInput
  IfNotExist, %A_ScriptDir%\MediaInfo\MediaInfo.exe
  {
     MsgBox, 1,, MediaInfo is missing.  Download Now or Cancel?
	 IfMsgBox OK
	 {
	   MC_Call = https://mediaarea.net/download/binary/mediainfo/20.03/MediaInfo_CLI_20.03_Windows_i386.zip
       Source = %A_ScriptDir%\MediaInfo_CLI.zip
       Dest = %A_ScriptDir%\MediaInfo
       FileCreateDir, %A_ScriptDir%\MediaInfo
	   DownloadFile(MC_Call,Source, True, True)
       Unz(Source, Dest)
       FileDelete, %A_ScriptDir%\MediaInfo_CLI.zip
       FileAppend,
	   MsgLog = Complted - Please check for any errors....`n%MsgLog%  
       GuiControl,,Log, %MsgLog%
     }
	 Else
	   return
  }
  IfNotExist, %A_ScriptDir%\ffmpeg\ffmpeg-N-102631-gbaf5cc5b7a-win64-gpl-shared\bin\ffmpeg.exe
  {
     MsgBox, 1,, FFmpeg is missing.  Download Now or Cancel?
	 IfMsgBox OK
	 {
	   MC_Call = https://github.com/BtbN/FFmpeg-Builds/releases/download/autobuild-2021-05-31-13-09/ffmpeg-N-102631-gbaf5cc5b7a-win64-gpl-shared.zip
       Source = %A_ScriptDir%\ffmpeg-N-102545-g59032494e8-win64-gpl-shared.zip
       Dest = %A_ScriptDir%\ffmpeg
       FileCreateDir, %A_ScriptDir%\ffmpeg
	   DownloadFile(MC_Call,Source, True, True)
       Unz(Source, Dest)
       FileDelete, %A_ScriptDir%\ffmpeg-N-102545-g59032494e8-win64-gpl-shared.zip
       FileAppend,
	   MsgLog = Completed - Please check for any errors....`n%MsgLog%  
       GuiControl,,Log, %MsgLog%
     }
	 Else
	   return
  }


  MC_Call = http://%MC_WS%/MCWS/v1/Files/Current?Action=Serialize
  GoSub, MCWS
  If Status !=200
    return
  MsgLog = Commenced - Please wait....`n%MsgLog%  
  IniWrite, %MCFieldTarget%, Swag of Tools.ini,MediaInfo_Settings,MCFieldTarget
  IniWrite, %MediaInfoField%, Swag of Tools.ini,MediaInfo_Settings,MediaInfoField
  IniWrite, %MediaInfoSection%, Swag of Tools.ini,MediaInfo_Settings,MediaInfoSection
  IniWrite, %NoMediaInfoFile%, Swag of Tools.ini,MediaInfo_Settings,NoMediaInfoFile
  IniWrite, %SkipMediaInfoFile%, Swag of Tools.ini,MediaInfo_Settings,SkipMediaInfoFile
  IniWrite, %CropScanPoint%, Swag of Tools.ini,MediaInfo_Settings,CropScanPoint
  IniWrite, %CropSettings%, Swag of Tools.ini,MediaInfo_Settings,CropSettings
  GuiControl,,Log, %MsgLog%
  Loop, parse, Result, `;
    If A_Index > 3
    {
	  MC_Key = %A_LoopField%
	  gosub, MC_GetInfo
      If Status !=200
        return
	  Else
        {
		 If MC_FileExt = bluray;1
		   RegExMatch(MC_FileNameExt,"(.*)(?=\\BDMV)", MC_FileNameExt)  ; keep this line
		 If MC_FileExt = bluray3d;1
		   RegExMatch(MC_FileNameExt,"(.*)(?=\\BDMV)", MC_FileNameExt)  ; keep this line
		 If MC_FileExt = dvd;1
		   RegExMatch(MC_FileNameExt,"(.*)(?=\\VIDEO_TS)", MC_FileNameExt)
		 If (SkipMediaInfoFile=0 || !FileExist(MC_FileName "_MediaInfo.txt"))
			{
			MsgLog = Commencing MediaInfo Scan (this may take some time) for "%MC_FileNameExt%"`n%MsgLog%
			GuiControl,,Log, %MsgLog%
			MediaInfoOutput = % CmdRet("MediaInfo\MediaInfo.exe" . " """ . MC_FileNameExt . """",, "UTF-8")
			MsgLog = Finished Reading Media Info for "%MC_FileNameExt%"`n%MsgLog%
			GuiControl,,Log, %MsgLog%
			If MC_FileExt = bluray;1
				{
				MPLSFileName = %MC_FileNameExt%\BDMV\PLAYLIST
				MediaInfoOutputExtra = % CmdRet("MediaInfo\MediaInfo.exe" . " """ . MPLSFileName . """",, "UTF-8")
				MediaInfoOutput = %MediaInfoOutput% + %MediaInfoOutputExtra%
				MsgLog = Reading Media Info for all the MPLS in "%MPLSFileName%"`n%MsgLog%
				GuiControl,,Log, %MsgLog%
				}
			If MC_FileExt = dvd;1
				{
				Loop, %MC_FileNameExt%\*.IFO, , 1
					{
					IFOFileName := A_LoopFileFullPath
					MediaInfoOutputExtra = % CmdRet("MediaInfo\MediaInfo.exe" . " """ . IFOFileName . """",, "UTF-8")
					MediaInfoOutput = %MediaInfoOutput% + %MediaInfoOutputExtra%
					MsgLog = Reading Media Infor for IFO file "%IFOFileName%"`n%MsgLog%
					GuiControl,,Log, %MsgLog%
					}
				}
			If (NoMediaInfoFile=0)
				{
				Filedelete, %MC_FileName%_MediaInfo.txt
				Fileappend, %MediaInfoOutput%, %MC_FileName%_MediaInfo.txt
				MsgLog = Created Extra file for "%MC_FileName%_MediaInfo.txt"`n%MsgLog%
				GuiControl,,Log, %MsgLog%
				}
			}
			Else 
				{
				FileRead, MediaInfoOutput, %MC_FileName%_MediaInfo.txt
				MsgLog = Using existing "Extra" MediaInfo file for "%MC_FileName%"`n%MsgLog%
				GuiControl,,Log, %MsgLog%
				}
			If StrLen(MediaInfoOutput) < 10 ;Skip if no MediaInfo Data
				{
				MsgLog = Skipping: Media Info could not analyse "%MC_FileName%_MediaInfo.txt"`n%MsgLog%
				GuiControl,,Log, %MsgLog%
				continue 
				}
		 IfInString, MediaInfoOutput, Atmos
		   IfNotInString, MC_Compression, Atmos
		   {
		    MC_Call = http://%MC_WS%/MCWS/v1/File/SetInfo?File=%MC_Key%&FileType=Key&Field=Compression&Value=%MC_Compression%+Atmos
            GoSub, MCWS
            If Status =200
				MsgLog = Atmos track found and MC Compression Field Updated`n%MsgLog%
		   }
		   Else MsgLog = Atmos track found and already in MC Compression Field`n%MsgLog%
		   
		 IfInString, MediaInfoOutput, : X /
		   IfNotInString, MC_Compression, DTS:X
		   {
		    MC_Call = http://%MC_WS%/MCWS/v1/File/SetInfo?File=%MC_Key%&FileType=Key&Field=Compression&Value=%MC_Compression%+DTS:X
            GoSub, MCWS
            If Status =200
				MsgLog = DTS:X track found and MC Compression Field Updated`n%MsgLog%
		   }
		   Else MsgLog = DTS:X track found and already in MC Compression Field`n%MsgLog%
		   
		 IfInString, MediaInfoOutput, : DTS XLL X
		   IfNotInString, MC_Compression, DTS:X
		   {
		    MC_Call = http://%MC_WS%/MCWS/v1/File/SetInfo?File=%MC_Key%&FileType=Key&Field=Compression&Value=%MC_Compression%+DTS:X
            GoSub, MCWS
            If Status =200
				MsgLog = DTS:X track found and MC Compression Field Updated`n%MsgLog%
		   }
		   Else MsgLog = DTS:X track found and already in MC Compression Field`n%MsgLog%
		   
		 AllMediaInfoLanguage =
		 Needle := "Language "
		 Loop, Parse, MediaInfoOutput, `n, `r
		 {
		   StringGetPos, LangPos, A_LoopField,%Needle% 
		   If (LangPos = 0)
		   {
		     RegExMatch(A_LoopField,"(?=:)(.*)", MediaInfoLanguage)
		     StringTrimLeft,MediaInfoLanguage,MediaInfoLanguage,2
			 IfNotInString,MC_Language,%MediaInfoLanguage%
			   IfNotInString,AllMediaInfoLanguage,%MediaInfoLanguage%
 		         AllMediaInfoLanguage = %AllMediaInfoLanguage%;%MediaInfoLanguage%
		   }
		 }

		 If AllMediaInfoLanguage
		 {
			AllMediaInfoLanguage = %MC_Language%%AllMediaInfoLanguage%
			Needle2 := ";"
			FoundPos := InStr(AllMediaInfoLanguage,Needle2)
			If (FoundPos = 1)
			  AllMediaInfoLanguage := SubStr(AllMediaInfoLanguage,2)

			MC_Call = http://%MC_WS%/MCWS/v1/File/SetInfo?File=%MC_Key%&FileType=Key&Field=Language&Value=%AllMediaInfoLanguage%
			GoSub, MCWS
			If Status =200
				MsgLog = Language track info found and MC Language Field Updated`n%MsgLog%
		 }

		If MCFieldTarget
		{
			MCFieldTargetArray := StrSplit(MCFieldTarget, ";")
			MediaInfoSectionArray := StrSplit(MediaInfoSection, ";")
			MediaInfoFieldArray := StrSplit(MediaInfoField, ";")
			Loop, % MCFieldTargetArray.MaxIndex()
			{
				MCFieldTargetCurrent := MCFieldTargetArray[A_Index]
				MediaInfoSectionCurrent := MediaInfoSectionArray[A_Index]
				MediaInfoFieldCurrent := MediaInfoFieldArray[A_Index]
				MediaInfoSectionFound = 0
				RegExMatch(Result,"(?<=MCFieldTargetCurrent"">)(.*)(?=</Field>)", AllMediaInfoValue)
				Loop, Parse, MediaInfoOutput, `n, `r
				{
					StringLeft, MediaInfoFieldName, A_LoopField, 40
					MediaInfoFieldName = %MediaInfoFieldName%

					If (A_LoopField = "")
						MediaInfoSectionFound = 0

					StringGetPos, SectionPos, A_LoopField,%MediaInfoSectionCurrent%

					If (SectionPos = 0)
						MediaInfoSectionFound = 1

					StringGetPos, FieldPos, A_LoopField,%MediaInfoFieldCurrent%

					If ((MediaInfoFieldName == MediaInfoFieldCurrent) and ((MediaInfoSectionFound = 1) or (MediaInfoSectionCurrent = "")))
					{
						RegExMatch(A_LoopField,"(?=:)(.*)", MediaInfoValue)
						StringTrimLeft, MediaInfoValue, MediaInfoValue, 2
						IfNotInString, AllMediaInfoValue, %MediaInfoValue%
						AllMediaInfoValue = %AllMediaInfoValue%; %MediaInfoValue%
					}
				}

;					If AllMediaInfoValue ; requested by lepa to blank values is no longer there
					{
						StringGetPos, SectionPos, AllMediaInfoValue,;%Space%
						If (SectionPos = 0)
							StringTrimLeft, AllMediaInfoValue, AllMediaInfoValue,2
						MCFieldTargetCurrentClean = % UriEncode(MCFieldTargetCurrent)
						AllMediaInfoValueClean = % UriEncode(AllMediaInfoValue)

						MC_Call = http://%MC_WS%/MCWS/v1/File/SetInfo?File=%MC_Key%&FileType=Key&Field=%MCFieldTargetCurrentClean%&Value=%AllMediaInfoValueClean%
						GoSub, MCWS
						If Status =200
							MsgLog = MC %MCFieldTargetCurrent% updated with %AllMediaInfoValue%`n%MsgLog%
					}
			}
		}
; FFMpeg section to scan for CropScanPoint
		CropScanPointNow := CropScanPoint
		If (CropScanPointNow > 0)
		{
			FileDelete, ffmpeg.txt
			Crop :=

			If (CropScanPointNow > MC_Duration)
			{
				MsgLog = Warning, The Crop Scan Point of %CropScanPointNow% sec is greater than the Duration of the Video (%MC_Duration% sec) `n%MsgLog%
				GuiControl,,Log, %MsgLog%
				CropScanPointNow = 1
			}


			MsgLog = Determining the Crop Factor for "%MC_FileName%" at the %CropScanPointNow% sec mark`n%MsgLog%
			GuiControl,,Log, %MsgLog%

			If MC_FileExt = bluray;1
			{
				LargestFileSizeKB = 0
				Loop, %MC_FileNameExt%\BDMV\STREAM\*.M2TS, , 1
				{
					If (A_LoopFileSizeKB > LargestFileSizeKB)
					{
						M2TSFileName := A_LoopFileFullPath
						LargestFileSizeKB := A_LoopFileSizeKB
					}
				}
				MC_FileNameExt := M2TSFileName
			}

			If MC_FileExt = dvd;1
			{
				LargestFileSizeKB = 0
				Loop, %MC_FileNameExt%\*.VOB, , 1
				{
					If (A_LoopFileSizeKB > LargestFileSizeKB)
					{
						VOBFileName := A_LoopFileFullPath
						LargestFileSizeKB := A_LoopFileSizeKB
					}
				}
				MC_FileNameExt := VOBFileName
			}
			runwait, %comspec% /c ffmpeg\ffmpeg-N-102631-gbaf5cc5b7a-win64-gpl-shared\bin\ffmpeg.exe -ss %CropScanPointNow% -i "%MC_FileNameExt%" -t 10 -vf cropdetect=%CropSettings% -f null dummy > ffmpeg.txt 2>&1,, Hide



			loop, Read, ffmpeg.txt
			{
				if InStr(A_LoopReadLine, "crop=")
				{
					FoundPos := InStr(A_LoopReadLine, "crop=")+5
					Crop := SubStr(A_LoopReadLine,FoundPos)
				}
			}

			MC_Call = http://%MC_WS%/MCWS/v1/File/SetInfo?File=%MC_Key%&FileType=Key&Field=Crop&Value=%Crop%
			GoSub, MCWS
			If Status =200
				MsgLog = The Crop Factor for "%MC_FileNameExt%" is %Crop%`n%MsgLog%
		}
; End FFMpeg Section
		}
		GuiControl,,Log, %MsgLog%   
	}

  MsgLog = Finished`n%MsgLog%  
  GuiControl,,Log, %MsgLog%
return

;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;++++++++++++++ MC FUNCTIONS and SUBS +++++++++++++++++++++++
;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;------------ Chapterise from chapter.TXT File --------------------

Process_Chapters:
MsgBox, On the next screen you can review and edit the Chapter File if needed.  When finished just Close Notepad (saving any changes) and you will then be asked if you want to Chapterfy based on this file.`n`n%AddMsg%
RunWait, %comspec% /c notepad.exe temp.chapters.txt,, Hide
MsgBox, 4,, Would you like to now Chapterfy:`n`n %MC_Name%
IfMsgBox No
  {
  MsgLog = Operation Cancelled`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  return
  }

Loop, read, temp.chapters.txt
    last_line++

Last_Chap:=last_line-2
File_Line:=1
Loop
{
  If File_Line > %last_line%
    Break

  ;------------- Create Chapter Particles ----------------

  MC_Call = http://%MC_WS%/MCWS/v1/File/CreateParticle?File=%MC_Key%&FileType=Key&Count=1
  GoSub, MCWS
  If Status !=200
    return
  RegExMatch(Result,"(?<=Keys"">)(.*)(?=</Item>)", MCParticle_Key)
  AllParticle_Keys = %AllParticle_Keys%,%MCParticle_Key%
  ;------------ Update Chapter Particles ----------------

  FileReadLine, Chapter_Start, temp.chapters.txt, File_Line
  FileReadLine, Chapter_Name, temp.chapters.txt, File_Line+1
  FileReadLine, Chapter_Finish, temp.chapters.txt, File_Line+2

  If File_Line > %last_line%
    Break

  StringTrimLeft, MC_Name, Chapter_Name, 14
  MC_Name = % UriEncode(MC_Name)

  StringTrimLeft, Chapter_Start, Chapter_Start, 10
  StringTrimLeft, Chapter_Finish, Chapter_Finish, 10
  MC_PlaybackRange = %Chapter_Start%-%Chapter_Finish%
  If File_Line > %Last_Chap%
     MC_PlaybackRange = %Chapter_Start%-

  MC_Call = http://%MC_WS%/MCWS/v1/File/SetInfo?File=%MCParticle_Key%&FileType=Key&Field=Name,Playback Range, Track #&Value=%MC_Name%,%MC_PlaybackRange%,%A_Index%&List=CSV

  GoSub, MCWS
  If Status !=200
    return

  File_Line:=File_Line+2
}
MsgLog = Created the following Chapter Particles %AllParticle_Keys%`n%MsgLog%
GuiControl,,Log, %MsgLog%
Return

;------------ Convert TXT to MPL File --------------------
Process_TXT:

AllToWrite =
MC_Field_Num = 0
Loop, parse, result, `n, `r
{
  numlines++
  If A_Index = 1
    StringSplit, MC_Field_Headers, A_Loopfield, %A_Tab%
}

FileAppend,<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>`n, %OutputFileName%
FileAppend,<MPL Version="2.0" Title="MC Fiddler Import %A_Now%" PathSeparator="\">`n, %OutputFileName%

Loop, parse, result, `n, `r
{
  L_Progress := A_Index/numlines*100

  If A_Index = 1 ; Skip first header line
    Continue
  If !A_Loopfield
    Continue
  If (L_Progress<>L_Progress_Last)
  {
    GuiControl,,Log, Writing %OutputFileName% - %L_Progress%`%`n%MsgLog%
	L_Progress_Last := L_Progress
  }
  StringSplit, MC_Field_Data, A_Loopfield, %A_Tab%

  AllToWrite = %AllToWrite%<Item>`r`n

  Loop, %MC_Field_Data0%
  {
    tempHeader := MC_Field_Headers%a_index%
    tempData := MC_Field_Data%a_index% 
    AllToWrite = %AllToWrite%<Field Name="%tempHeader%">%tempData%</Field>`r`n
  }
  AllToWrite = %AllToWrite%</Item>`r`n

  If !(Mod(A_Index,10000)) ; wite files in chuncks so to not run out of memory but keep file writes low
  {
    StringReplace, AllToWrite, AllToWrite, %EOLChar%, `r`n, All
    StringReplace, AllToWrite, AllToWrite, %CRChar%, `r, All
    StringReplace, AllToWrite, AllToWrite, %LFChar%, `n, All
    StringReplace, AllToWrite, AllToWrite, %TabChar%, %A_Tab%, All
    FileAppend, %AllToWrite%, %OutputFileName%
    AllToWrite =
  }
}
MsgLog = Preparing MPL File, Please Wait`n%MsgLog%
GuiControl,,Log, %MsgLog%
StringReplace, AllToWrite, AllToWrite, %EOLChar%, `r`n, All
StringReplace, AllToWrite, AllToWrite, %CRChar%, `r, All
StringReplace, AllToWrite, AllToWrite, %LFChar%, `n, All
StringReplace, AllToWrite, AllToWrite, %TabChar%, %A_Tab%, All

FileAppend, %AllToWrite%</MPL>`n, %OutputFileName%

MsgLog = Created %OutputFileName%`n%MsgLog%
GuiControl,,Log, %MsgLog%
Return

;------------ Convert MPL to TXT File --------------------
Process_MPL:
MsgLog = Preparing MPL File, Please Wait`n%MsgLog%
GuiControl,,Log, %MsgLog%
; ----------------------------------------------------------------------------------------------------------------------
XML_Obj := ComObjCreate("Msxml2.DOMDocument.6.0")
XML_Obj.setProperty("SelectionLanguage", "XPath")
XML_Obj.async := False
XML_Path = %MPLFileName%
XML_Obj.load(XML_Path)
XML_Doc := XML_Obj.documentElement
; ----------------------------------------------------------------------------------------------------------------------
NamesArr := {}
For Node In XML_Doc.selectNodes("Item/Field/@Name")
   NamesArr[Node.Value] := ""
NamesCount := NamesArr.Count()
; ----------------------------------------------------------------------------------------------------------------------
ResultStr := ""
For Name In NamesArr
   ResultStr .= Name . (A_Index < NamesCount ? "`t" : "")
For Item In XML_Doc.selectNodes("Item") {
   TextStr := ""
   For Name In NamesArr {
      Txt := Item.selectSingleNode("Field[@Name='" . Name . "']").text
      TextStr .= Txt . (A_Index < NamesCount ? "`t" : "")
   }
   GuiControl,,Log, Processing Item %A_Index% `n%MsgLog%
   
   ; --------- Swap out Special Chars that break the TXT File ----------
   StringReplace, TextStr, TextStr, `n, {LF}, All
   ResultStr .= "`n" . TextStr
}

FileAppend, %ResultStr%, %OutputFileName%, UTF-16

MsgLog = Created %OutputFileName%`n%MsgLog%
GuiControl,,Log, %MsgLog%
Return

;------------ Update MC from TXT File --------------------
Update_MC:

MC_Field_Data := ""
AllToWrite =
MC_Items_Updated = 0
MC_Field_Num = 0
Loop, parse, result, `n, `r
{
  numlines++
  If A_Index = 1
    StringSplit, MC_Field_Headers, A_Loopfield, %A_Tab%
}

Loop, parse, result, `n, `r
{
  If A_Index = 1 ; Skip first header line
    Continue

  GuiControl,,Log, Processing Row %A_Index% of the TXT File `n%MsgLog%

  StringSplit, MC_Field_Data, A_Loopfield, %A_Tab%

  FieldsToWrite = http://%MC_WS%/MCWS/v1/File/SetInfo?File=&FileType=Key&Field=ZNoField
  DataToWrite = &Value=zNoData

  Loop, %MC_Field_Data0%
  {
    tempField := MC_Field_Headers%a_index%
    tempData := MC_Field_Data%a_index% 
    ; find key
    if tempField = Key
    {
      MC_Key = %tempData%
      Continue
    }
    StringReplace, AllToWrite, AllToWrite, %EOLChar%, `r`n, All
    StringReplace, tempData, tempData, %CRChar%, `r, All
    StringReplace, tempData, tempData, %LFChar%, `n, All
    StringReplace, tempData, tempData, %TabChar%, %A_Tab%, All
    StringReplace, tempData, tempData, ", "", All
    tempData = % UriEncode(tempData)
    FieldsToWrite = %FieldsToWrite%,%tempField%
    DataToWrite = %DataToWrite%,%tempData%
  }
; New Create Call to go in here  
  If MC_Key = new
  {
	MC_Call = http://%MC_WS%/MCWS/v1/Library/CreateFile
    GoSub, MCWS
    If Status !=200
      return
	RegExMatch(Result,"(?<=Key"">)(.*)(?=</Item)", MC_Key)
  }
  
  StringReplace, FieldsToWrite, FieldsToWrite, File=, File=%MC_Key%
  StringReplace, FieldsToWrite, FieldsToWrite, =`,, =
  StringReplace, DataToWrite, DataToWrite, =`,, =
  If MC_Key
  {
    MC_Call = %FieldsToWrite%%DataToWrite%&List=CSV

	StringReplace, MC_Call, MC_Call, `%2c, "`%2c", All ; Escape out Comma
    StringReplace, MC_Call, MC_Call, `%0D, "`%0D", All ; Escape out CR
    StringReplace, MC_Call, MC_Call, `%0A, "`%0A", All ; Escape out LF
    StringReplace, MC_Call, MC_Call, `%0D""`%0A, `%0D`%0A, All ; Remove extra Escape out CR/LF (EOL Char)
    StringReplace, MC_Call, MC_Call, `%09, "`%09", All ; Escape out Tab	
	MC_Items_Updated := MC_Items_Updated + 1
    GoSub, MCWS
  }
  MC_Key =
}
MsgLog = Updated %MC_Items_Updated% items in MC from the TXT File`n%MsgLog%
If MC_Items_Updated = 0
	MsgLog = No Items were updated - Check you have a Col Heading called "Key"`n%MsgLog%
GuiControl,,Log, %MsgLog%
Return
;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;+++++++++ MC Std FUNCTIONS and SUBS +++++++++++++++++++++
;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;------------ MC Get Current Item Key(s)-----------------
MC_Current:
  MC_Call = http://%MC_WS%/MCWS/v1/Files/Current?Action=Serialize
  GoSub, MCWS
  If Status !=200
    return
  RegExMatch(result,"[^;]+$", MC_Key)
  If MC_Key < 1
    {
    MsgLog = Nothing is Selected in MC!`n%MsgLog%
    GuiControl,,Log, %MsgLog%
    return
    }
  RegExMatch(Result,"(?<=-1;)(.*)", Result)
  StringSplit, MC_Chapters, Result, ';
Return

;------------ MC Get Details with an Key-----------------
MC_GetInfo:
  MC_Call = http://%MC_WS%/MCWS/v1/File/GetInfo?File=%MC_Key%&Fields=Name,Playback Range,Filename,File Type,Compression,Language,Duration
  GoSub, MCWS

  RegExMatch(Result,"(?<=Name"">)(.*)(?=</Field>)", MC_Name)
  RegExMatch(Result,"(?<=Range"">)(.*)(?=-)", MC_Start)
  RegExMatch(Result,"(?<=Filename"">)(.*)(?=[.])", MC_Filename)
  RegExMatch(Result,"(?<=Filename"">)(.*)(?=</Field>)", MC_FilenameExt)
  RegExMatch(Result,"(?<=Type"">)(.*)(?=</Field>)", MC_FileType)
  RegExMatch(Result,"(?<=Compression"">)(.*)(?=</Field>)", MC_Compression)
  RegExMatch(Result,"(?<=Language"">)(.*)(?=</Field>)", MC_Language)
  RegExMatch(Result,"(?<=Duration"">)(.*)(?=</Field>)", MC_Duration)
  StringTrimLeft, MC_FileExt, MC_FilenameExt, StrLen(MC_Filename)+1
Return

;------------ MC Get Authentication Token-----------------
MC_GetToken:
  MC_Call = http://%MC_WS%/MCWS/v1/Authenticate
  GoSub, MCWS
  RegExMatch(Result,"(?<=Token"">)(.*)(?=</Item>)", MC_Token)
Return

;---------- CheckMC and User Settings ----------------
ButtonTestMCWS:
  Gui, Submit, NoHide
  guiControlGet, UserInput
  MsgLog = Checking for Media Center at %MC_WS%`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  MC_Call = http://%MC_WS%/MCWS/v1/Authenticate
  GoSub, MCWS
  If Status = 200 
    {
	MsgLog = Success, Media Center Found at %MC_WS%!`n%MsgLog%
    GuiControl,,Log, %MsgLog%
	}
Return

;---------- Make INI and Registry Changes ----------------
ButtonSaveMCWS:
  Gui, Submit, NoHide
  guiControlGet, UserInput
  MsgLog = Writing to Swag of Tools.ini`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  IniWrite, %MC_Ver%, Swag of Tools.ini,MC_Settings,MC_Ver
  IniWrite, %MC_WS%, Swag of Tools.ini,MC_Settings,MC_WS
  
  Hashed_UserName := ""
  Hashed_Password := ""
  Loop, Parse, MC_UserName
	Hashed_UserName .= Chr(((Asc(A_LoopField)-32)^(Asc(SubStr(PWKey,A_Index,1))-32))+32)
  Loop, Parse, MC_Password
	Hashed_Password .= Chr(((Asc(A_LoopField)-32)^(Asc(SubStr(PWKey,A_Index,1))-32))+32)
  IniWrite, %Hashed_UserName%, Swag of Tools.ini,MC_Settings,MC_UserName
  IniWrite, %Hashed_Password%, Swag of Tools.ini,MC_Settings,MC_Password
  
  RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\J. River\Media Center %MC_Ver%\Properties\External Tools,Swag of Tools - Main Menu,%A_ScriptDir%\Swag of Tools.exe||0
  RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\J. River\Media Center %MC_Ver%\Properties\External Tools,Swag of Tools - Chapterfy (Auto),%A_ScriptDir%\Swag of Tools.exe|Auto|0
  RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\J. River\Media Center %MC_Ver%\Properties\External Tools,Swag of Tools - Chapterfy (ReadChapterFile),%A_ScriptDir%\Swag of Tools.exe|ReadChapterFile|0
  RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\J. River\Media Center %MC_Ver%\Properties\External Tools,Swag of Tools - Chapterfy (CreateChapterFile),%A_ScriptDir%\Swag of Tools.exe|CreateChapterFile|0
  RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\J. River\Media Center %MC_Ver%\Properties\External Tools,Swag of Tools - Burn2BD,%A_ScriptDir%\Swag of Tools.exe|Burn2BD|0 
  RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\J. River\Media Center %MC_Ver%\Properties\External Tools,Swag of Tools - Filename Updater,%A_ScriptDir%\Swag of Tools.exe|FilenameUpdater|0 
  RegWrite, REG_SZ, HKEY_CURRENT_USER, SOFTWARE\J. River\Media Center %MC_Ver%\Properties\External Tools,Swag of Tools - MediaInfo,%A_ScriptDir%\Swag of Tools.exe|MediaInfo|0 
  MsgLog = Writing Registry Keys to HKEY_CURRENT_USER, SOFTWARE\J. River\Media Center %MC_Ver%`n%MsgLog%

  MC_Call = http://%MC_WS%/MCWS/v1/Library/Fields
  GoSub, MCWS

  MC_Fields := Result

  MsgLog = Check if "Crop" Field already exists in MC Library`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  IfInString, MC_Fields, Crop
    MsgLog = "Crop" Field already exists in MC Library`n%MsgLog%
  Else
	{  
    MC_Call = http://%MC_WS%/MCWS/v1/Library/CreateField?Name=Crop&Type=string
    MsgLog = Adding "Crop" Field to MC Library`n%MsgLog%
    GoSub, MCWS
    }

  MsgLog = Check if "Aspect Ratio (Crop)" Field already exists in MC Library`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  IfInString, MC_Fields, Aspect Ratio (Crop)
    MsgLog = "Aspect Ratio (Crop)" Field already exists in MC Library`n%MsgLog%
  Else
	{  
    MC_Call = http://%MC_WS%/MCWS/v1/Library/CreateField?Name=Aspect Ratio (Crop)&Expression=save(ifelse(not(isempty([crop])), Math(ListItem([Crop],0,:)//ListItem([Crop],1,:)),not(isempty([Aspect Ratio])), Math(Replace([Aspect Ratio],:,//)),1,Math(Replace([Dimensions],x,//))),v_ARCalculated)IfCase(Replace([v_ARCalculated],/,,.), 3, 0.30, unknown, 0.9, portrait, 1.17, 1.00, 1.5, 1.33, 1.72, 1.66, 1.82, 1.78, 1.93, 1.85, 2.1, 2.00, 2.28, 2.20, 2.37, 2.35, 2.47, 2.39, 2.6, 2.55, 2.71, 2.65, 2.99, 2.76, 20, wide)
    MsgLog = Adding "Aspect Ratio (Crop)" Field to MC Library`n%MsgLog%
    GoSub, MCWS
    }

  MsgLog = Check if "FileKey" Field already exists in MC Library`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  IfInString, MC_Fields, FileKey
    MsgLog = "FileKey" Field already exists in MC Library`n%MsgLog%
  Else
	{  
    MC_Call = http://%MC_WS%/MCWS/v1/Library/CreateField?Name=FileKey&Expression=FileKey()
    MsgLog = Adding "FileKey" Field to MC Library`n%MsgLog%
    GoSub, MCWS
    }

    GuiControl,,Log, %MsgLog%
	Return

;-------------MCWS---------------------------
MCWS:
  WinHTTP := ComObjCreate("WinHTTP.WinHttpRequest.5.1")
  ComObjError(false)
  WinHTTP.Open("GET", MC_Call)
  WinHTTP.SetCredentials(MC_UserName,MC_Password,0)
  WinHTTP.SetRequestHeader("Content-type", "application/x-www-form-urlencoded; charset=UTF-8")
  Body = ""
  WinHTTP.Send(Body)
  Result := WinHTTP.ResponseText
  Status := WinHTTP.Status

  Result := BinArr_ToString(WinHTTP.ResponseBody, "UTF-8")
  Result := StrReplace(Result,"&amp;","&")

  IfInString, Result, "No changes."
	{
	MsgLog = --- MC already has the same information for the above results ---`n%MsgLog%
	Status = 200
	}

  Gui, Submit, NoHide
	If (VerboseLogging = 1)
	{
	MsgLog = MCWS Status - %Status%  MCWS Call and Result (see details below)`n%MC_Call%`n%Result%`n%MsgLog%
	GuiControl,,Log, %MsgLog%
	}
  If Status = 200 
    Return
  else If !Status
    MsgLog = Error - MediaCenter not Found at %MC_WS% `n- Check MC Server is running and the Configuration Options are Correct`n%MsgLog%
  else If Status = 401
    MsgLog = Error 401 Access Denied`n- Check Username / Password is correct in the Configuration Options`n%MsgLog%
  else MsgLog = Error - %Status%  Processed Failed (see details below)`n%MC_Call%`n%Result%`n%MsgLog%
  GuiControl,,Log, %MsgLog%
Return

;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;+++++++++ 3rd Party FUNCTIONS and SUBS +++++++++++++++++++++
;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

;--------------ChapterDB API Call-------------------------
CDB_API:
  WinHTTP := ComObjCreate("WinHTTP.WinHttpRequest.5.1")
  ComObjError(false)
  WinHTTP.Open("GET",CDB_Call)
  WinHTTP.SetRequestHeader("ApiKey","IVB32FJYYTFQDP5UQ1A7")
  Body = ""
  WinHTTP.Send(Body)
  Result := WinHTTP.ResponseText
  Status := WinHTTP.Status
  RegExMatch(Result,"(?<=chapterSetId>)(.*)(?=</chapterSetId>)", CDB_ID)
Return

;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;+++++++++ GENERAL FUNCTIONS and SUBS +++++++++++++++++++++++
;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

; ---------------UTF-8 Encoding-------------------------
; Thanks to mikeyww at - https://www.autohotkey.com/boards/viewtopic.php?t=87449&p=384631
BinArr_ToString(BinArr, Encoding := "UTF-8") {
 ; https://gist.github.com/tmplinshi/a97d9a99b9aa5a65fd20
 ; https://www.autohotkey.com/boards/viewtopic.php?p=100984#p100984
 oADO := ComObjCreate("ADODB.Stream"), oADO.Type := 1, oADO.Mode := 3 ; adTypeBinary, adModeReadWrite
 oADO.Open
 oADO.Write(BinArr)
 oADO.Position := 0, oADO.Type := 2, oADO.Charset := Encoding ; adTypeText
 Return oADO.ReadText, oADO.Close
}

; ----------Unzip -------------------------------
Unz(sZip, sUnz)
{
	fso := ComObjCreate("Scripting.FileSystemObject")
	If !fso.FolderExists(sUnz)
		 fso.CreateFolder(sUnz)
	ComObjCreate("Shell.Application").Namespace(sUnz).CopyHere(sZip "\*", 4|16)
}

; ----------URL Encode/Decode -------------------------------
; Thanks to users at - https://autohotkey.com/board/topic/75390-ahk-l-unicode-uri-encode-url-encode-function/
; modified from jackieku's code (http://www.autohotkey.com/forum/post-310959.html#310959)
UriEncode(Uri, Enc = "UTF-8")
{
	StrPutVar(Uri, Var, Enc)
	f := A_FormatInteger
	SetFormat, IntegerFast, H
	Loop
	{
		Code := NumGet(Var, A_Index - 1, "UChar")
		If (!Code)
			Break
		If (Code >= 0x30 && Code <= 0x39 ; 0-9
			|| Code >= 0x41 && Code <= 0x5A ; A-Z
			|| Code >= 0x61 && Code <= 0x7A) ; a-z
			Res .= Chr(Code)
		Else
			Res .= "%" . SubStr(Code + 0x100, -1)
	}
	SetFormat, IntegerFast, %f%
	Return, Res
}

UriDecode(Uri, Enc = "UTF-8")
{
	Pos := 1
	Loop
	{
		Pos := RegExMatch(Uri, "i)(?:%[\da-f]{2})+", Code, Pos++)
		If (Pos = 0)
			Break
		VarSetCapacity(Var, StrLen(Code) // 3, 0)
		StringTrimLeft, Code, Code, 1
		Loop, Parse, Code, `%
			NumPut("0x" . A_LoopField, Var, A_Index - 1, "UChar")
		StringReplace, Uri, Uri, `%%Code%, % StrGet(&Var, Enc), All
	}
	Return, Uri
}

StrPutVar(Str, ByRef Var, Enc = "")
{
	Len := StrPut(Str, Enc) * (Enc = "UTF-16" || Enc = "CP1200" ? 2 : 1)
	VarSetCapacity(Var, Len, 0)
	Return, StrPut(Str, &Var, Enc)
}

; ----------------------------------------------------------------------------------------------------------------------
; Name .........: FileToHex
; Description ..: Get the hexadecimal code of a file. Use this with the Windows "Send To" feature.
; AHK Version ..: AHK_L x32/64 Unicode
; Author .......: Cyruz - http://ciroprincipe.info
; License ......: WTFPL - http://www.wtfpl.net/txt/copying/
; Changelog ....: Dic. 31, 2013 - v0.1 - AHK_L version.
; ..............: Jan. 04, 2014 - v0.2 - Little adjustment to CryptBinToHex. Use it as default.
; ----------------------------------------------------------------------------------------------------------------------
FileToHex:
;  N_SPLIT := 112 ; Adjust this to split every X chars. If = 0 split will not occur.
  FileRead, cBuf, *C %MC_FilenameExt%
  ; Transform its content in Hexadecimal. CryptBinToHex is the fastest function of the two.
  ; It will be used if present in the system, otherwise we fall back to ToHex.
  (FileExist(A_WinDir "\System32\Crypt32.dll")) ? CryptBinToHex(sHex, cBuf) : ToHex(sHex, cBuf)
  ; Insert a newline every X char.
  If ( N_SPLIT ) {
      sSplitHex := RegExReplace(sHex, "iS)(.{" N_SPLIT "})", "$1`n")
      StringTrimRight, sSplitHex, sSplitHex, 1 ; Remove last `n.
      }
  HexFile := (N_SPLIT) ? sSplitHex : sHex
  Return 

  ; Thanks to Laszlo: http://www.autohotkey.com/forum/viewtopic.php?p=131700#131700.
  ToHex(ByRef sHex, ByRef cBuf, nSz:=-1) {
    nBz := VarSetCapacity(cBuf)
    adr := &cBuf
    f := A_FormatInteger
    SetFormat, Integer, Hex
    Loop % nSz < 0 ? nBz : nSz
        sHex .= *adr++
    SetFormat, Integer, %f%
    sHex := RegExReplace(sHex, "S)x(?=.0x|.$)|0x(?=..0x|..$)")
  }
 
  ; Thanks to nnnik: http://ahkscript.org/boards/viewtopic.php?f=6&t=1242#p8376.
  CryptBinToHex(ByRef sHex, ByRef cBuf) {
    szBuf := VarSetCapacity(cBuf)
    DllCall( "Crypt32.dll\CryptBinaryToString", Ptr,&cBuf, UInt,szBuf, UInt,4, Ptr,0, UIntP,szHex )
    VarSetCapacity(cHex, szHex*2, 0)
    DllCall( "Crypt32.dll\CryptBinaryToString", Ptr,&cBuf, UInt,szBuf, UInt,4, Ptr,&cHex, UIntP,szHex )
    sHex := RegExReplace(StrGet(&cHex, szHex, "UTF-16"), "S)\s")
  }

; -----------hexToDecimal-----------------------------------------------------------------------------------------------------------
; Thanks to users at - http://www.autohotkey.com/board/topic/95502-getting-md5-hash-works-in-ansi-but-fails-with-unicode/#entry601707
hexToDecimal(str){
    static _0:=0,_1:=1,_2:=2,_3:=3,_4:=4,_5:=5,_6:=6,_7:=7,_8:=8,_9:=9,_a:=10,_b:=11,_c:=12,_d:=13,_e:=14,_f:=15
    str:=ltrim(str,"0x `t`n`r"),   len := StrLen(str),  ret:=0
    Loop,Parse,str
      ret += _%A_LoopField%*(16**(len-A_Index))
    return ret
}

;-------------------MD5 Hash Calculator------------------
; Thanks to users at - http://www.autohotkey.com/board/topic/95502-getting-md5-hash-works-in-ansi-but-fails-with-unicode/#entry601707
MD5(ByRef MD5Filepath)
{
  FileGetSize, MD5Size, %MD5Filepath%
  If !MD5Size
    {
    msgbox, Error - The File is Empty or Does Not Exist`n%MD5Filepath%
    return
    }
  FileRead MD5Data, *c %MD5Filepath%
  VarSetCapacity(MD5_CTX, 104)
  DllCall("advapi32\MD5Init",  "Ptr", &MD5_CTX)
  DllCall("advapi32\MD5Update","Ptr", &MD5_CTX, "Str", MD5Data, "UInt", MD5Size)
  DllCall("advapi32\MD5Final", "Ptr", &MD5_CTX)
  VarSetCapacity(Hash, Size := 64 * (A_IsUnicode ? 2 : 1))
  DllCall("Crypt32.dll\CryptBinaryToString"
        , "Ptr", &MD5_CTX+88
        , "UInt", 16
        , "UInt", 4
        , "Str", Hash
        , "UIntP", Size
        , "CDECL UInt")
  StringUpper Hash, Hash
  return  RegExReplace(Hash,"[^A-F0-9]")
}

;------------------ Change Sec to Chapter Time ---------
SecToChapTime(decsec) 
{
  hrs := floor(decsec/60/60)
  if hrs < 10
    hrs = 0%hrs%
  min := floor(decsec/60 - hrs*60)
  if min < 10
    min = 0%min%
  sec := decsec - hrs*60*60 - min*60
  if sec < 10
    sec = 0%sec%
  StringLeft, sec, sec, 6
  Return Hrs ":" Min ":" Sec
}

; -----------Hex2Text(Hex)-----------------------------------------------------------------------------------------------------------
; Thanks to users at - http://www.autohotkey.com/board/topic/76561-ascii-to-hex-to-ascii-again-unicode/
Hex2Text(Hex) {
	startpos:=1
	Loop % StrLen(Hex)/2
		{
		n .= Chr( "0x" . SubStr(Hex, StartPos+2 , 2) . SubStr(Hex, StartPos , 2) )
		startpos +=4
		}
	Return n
	}

; -----------Pipe "runwait, %comspec%..." output to a Variable?----------------------------------------------------------------
; Thanks to users including teadrinker at - https://www.autohotkey.com/boards/viewtopic.php?f=76&t=96199
CmdRet(sCmd, callBackFunc := "", encoding := "")
{
   static flags := [HANDLE_FLAG_INHERIT := 0x1, CREATE_NO_WINDOW := 0x8000000], STARTF_USESTDHANDLES := 0x100
        
   (encoding = "" && encoding := "cp" . DllCall("GetOEMCP", "UInt"))
   DllCall("CreatePipe", "PtrP", hPipeRead, "PtrP", hPipeWrite, "Ptr", 0, "UInt", 0)
   DllCall("SetHandleInformation", "Ptr", hPipeWrite, "UInt", flags[1], "UInt", flags[1])
   
   VarSetCapacity(STARTUPINFO , siSize :=    A_PtrSize*9 + 4*8, 0)
   NumPut(siSize              , STARTUPINFO)
   NumPut(STARTF_USESTDHANDLES, STARTUPINFO, A_PtrSize*4 + 4*7)
   NumPut(hPipeWrite          , STARTUPINFO, siSize - A_PtrSize*2)
   NumPut(hPipeWrite          , STARTUPINFO, siSize - A_PtrSize)
   
   VarSetCapacity(PROCESS_INFORMATION, A_PtrSize*2 + 4*2, 0)
   if !DllCall("CreateProcess", "Ptr", 0, "Str", sCmd, "Ptr", 0, "Ptr", 0, "UInt", true, "UInt", flags[2]
                              , "Ptr", 0, "Ptr", 0, "Ptr", &STARTUPINFO, "Ptr", &PROCESS_INFORMATION)
   {
      DllCall("CloseHandle", "Ptr", hPipeRead)
      DllCall("CloseHandle", "Ptr", hPipeWrite)
      throw "CreateProcess is failed"
   }
   DllCall("CloseHandle", "Ptr", hPipeWrite)
   VarSetCapacity(sTemp, 4096), nSize := 0
   while DllCall("ReadFile", "Ptr", hPipeRead, "Ptr", &sTemp, "UInt", 4096, "UIntP", nSize, "UInt", 0) {
      sOutput .= stdOut := StrGet(&sTemp, nSize, encoding)
      ( callBackFunc && %callBackFunc%(stdOut) )
   }
   DllCall("CloseHandle", "Ptr", NumGet(PROCESS_INFORMATION))
   DllCall("CloseHandle", "Ptr", NumGet(PROCESS_INFORMATION, A_PtrSize))
   DllCall("CloseHandle", "Ptr", hPipeRead)
   Return sOutput
}


;============ Download Function =============
; All thanks to use Bruttosozialprodukt for this fucntion at at http://www.autohotkey.com/board/topic/101007-super-simple-download-with-progress-bar/
DownloadFile(UrlToFile, SaveFileAs, Overwrite := True, UseProgressBar := True) {
    ;Check if the file already exists and if we must not overwrite it
      If (!Overwrite && FileExist(SaveFileAs))
          Return
    ;Check if the user wants a progressbar
      If (UseProgressBar) {
          ;Initialize the WinHttpRequest Object
            WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
          ;Download the headers
            WebRequest.Open("HEAD", UrlToFile)
            WebRequest.Send()
          ;Store the header which holds the file size in a variable:
            FinalSize := WebRequest.GetResponseHeader("Content-Length")
          ;Create the progressbar and the timer
		  	MsgLog = Downloading File : %UrlToFile%`n%MsgLog%
			GuiControl,,Log, %MsgLog%
;            Progress, H80, , Downloading %UrlToFile%
            SetTimer, __UpdateProgressBar, 100
      }
    ;Download the file
      UrlDownloadToFile, %UrlToFile%, %SaveFileAs%
    ;Remove the timer and the progressbar because the download has finished
      If (UseProgressBar) {
;          Progress, Off
          SetTimer, __UpdateProgressBar, Off
      }
    Return
    
    ;The label that updates the progressbar
      __UpdateProgressBar:
          ;Get the current filesize and tick
            CurrentSize := FileOpen(SaveFileAs, "r").Length ;FileGetSize wouldn't return reliable results
            CurrentSizeTick := A_TickCount
          ;Calculate the downloadspeed
            Speed := Round((CurrentSize/1024-LastSize/1024)/((CurrentSizeTick-LastSizeTick)/1000)) . " Kb/s"
          ;Save the current filesize and tick for the next time
            LastSizeTick := CurrentSizeTick
            LastSize := FileOpen(SaveFileAs, "r").Length
          ;Calculate percent done
            PercentDone := Round(CurrentSize/FinalSize*100)
          ;Update the ProgressBar
			GuiControl,,Log, Downloading File at %Speed% / %PercentDone%`%`n%MsgLog%
;            Progress, %PercentDone%, Please Wait - This may take awhile!, Downloading at...  (%Speed%), Preparing %MC_Name%
      Return
}
; =============== UUID Function =====================
; Thanks to FanaticGuru from https://www.autohotkey.com/boards/viewtopic.php?t=71638&p=310199
; Function UUID
; 	returns UUID member of the System Information structure in the SMBIOS information
;	this should be unique to a particular computer
UUID()
{
	For obj in ComObjGet("winmgmts:{impersonationLevel=impersonate}!\\" . A_ComputerName . "\root\cimv2").ExecQuery("Select * From Win32_ComputerSystemProduct")
		return obj.UUID
}
;=================== APP END ACTIONS ==========================
ButtonCancel:
GuiClose:
Script_End:
Gui, Hide
FileDelete, temp.chapters.txt
Progress, Off
exitapp
