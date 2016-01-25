Attribute VB_Name = "MWinAPIShell32Folder"
' ==========================================================================
' Module      : MWinAPIShell32Folder
' Type        : Module
' Description : Support for GetFolderPath and GetKnownFolder
' --------------------------------------------------------------------------
' Procedures  : GUIDToCSIDLEquivalent           String
'               GUIDToKnownFolderID             String
'               ShellGetFolderPath              String
'               ShellGetKnownFolderByGUID       String
'               ShellGetKnownFolderPath         String
' --------------------------------------------------------------------------
' Dependencies: MWinAPIShell32
' --------------------------------------------------------------------------
' Comments    : As of Windows Vista, SHGetFolderPath is
'               merely a wrapper for SHGetKnownFolderPath.
'               The CSIDL value is translated to its
'               associated KNOWNFOLDERID and then
'               SHGetKnownFolderPath is called.
'               New applications should use the known folder
'               system rather than the older CSIDL system,
'               which is supported only for backward compatibility.
'               Both are fully implemented in this module.
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

' KNOWNFOLDER IDs are described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd378457(v=vs.85).aspx
' -----------------------------------
Public Const FOLDERID_AccountPictures           As String = "{008CA0B1-55B4-4C56-B8A8-4DE4B299D3BE}"
Public Const FOLDERID_AddNewPrograms            As String = "{DE61D971-5EBC-4F02-A3A9-6C82895E5C04}"
Public Const FOLDERID_AdminTools                As String = "{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"
Public Const FOLDERID_ApplicationShortcuts      As String = "{A3918781-E5F2-4890-B3D9-A7E54332328C}"
Public Const FOLDERID_AppsFolder                As String = "{1E87508D-89C2-42F0-8A7E-645A0F50CA58}"
Public Const FOLDERID_AppUpdates                As String = "{A305CE99-F527-492B-8B1A-7E76FA98D6E4}"
Public Const FOLDERID_CameraRoll                As String = "{AB5FB87B-7CE2-4F83-915D-550846C9537B}"
Public Const FOLDERID_CDBurning                 As String = "{9E52AB10-F80D-49DF-ACB8-4330F5687855}"
Public Const FOLDERID_ChangeRemovePrograms      As String = "{DF7266AC-9274-4867-8D55-3BD661DE872D}"
Public Const FOLDERID_CommonAdminTools          As String = "{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"
Public Const FOLDERID_CommonOEMLinks            As String = "{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"
Public Const FOLDERID_CommonPrograms            As String = "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"
Public Const FOLDERID_CommonStartMenu           As String = "{A4115719-D62E-491D-AA7C-E74B8BE3B067}"
Public Const FOLDERID_CommonStartup             As String = "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"
Public Const FOLDERID_CommonTemplates           As String = "{B94237E7-57AC-4347-9151-B08C6C32D1F7}"
Public Const FOLDERID_ComputerFolder            As String = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"
Public Const FOLDERID_ConflictFolder            As String = "{4BFEFB45-347D-4006-A5BE-AC0CB0567192}"
Public Const FOLDERID_ConnectionsFolder         As String = "{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}"
Public Const FOLDERID_Contacts                  As String = "{56784854-C6CB-462B-8169-88E350ACB882}"
Public Const FOLDERID_ControlPanelFolder        As String = "{82A74AEB-AEB4-465C-A014-D097EE346D63}"
Public Const FOLDERID_Cookies                   As String = "{2B0F765D-C0E9-4171-908E-08A611B84FF6}"
Public Const FOLDERID_Desktop                   As String = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
Public Const FOLDERID_DeviceMetadataStore       As String = "{5CE4A5E9-E4EB-479D-B89F-130C02886155}"
Public Const FOLDERID_Documents                 As String = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
Public Const FOLDERID_DocumentsLibrary          As String = "{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}"
Public Const FOLDERID_Downloads                 As String = "{374DE290-123F-4565-9164-39C4925E467B}"
Public Const FOLDERID_Favorites                 As String = "{1777F761-68AD-4D8A-87BD-30B759FA33DD}"
Public Const FOLDERID_Fonts                     As String = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
Public Const FOLDERID_Games                     As String = "{CAC52C1A-B53D-4EDC-92D7-6B2E8AC19434}"
Public Const FOLDERID_GameTasks                 As String = "{054FAE61-4DD8-4787-80B6-090220C4B700}"
Public Const FOLDERID_History                   As String = "{D9DC8A3B-B784-432E-A781-5A1130A75963}"
Public Const FOLDERID_HomeGroup                 As String = "{52528A6B-B9E3-4ADD-B60D-588C2DBA842D}"
Public Const FOLDERID_HomeGroupCurrentUser      As String = "{9B74B6A3-0DFD-4f11-9E78-5F7800F2E772}"
Public Const FOLDERID_ImplicitAppShortcuts      As String = "{BCB5256F-79F6-4CEE-B725-DC34E402FD46}"
Public Const FOLDERID_InternetCache             As String = "{352481E8-33BE-4251-BA85-6007CAEDCF9D}"
Public Const FOLDERID_InternetFolder            As String = "{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}"
Public Const FOLDERID_Libraries                 As String = "{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}"
Public Const FOLDERID_Links                     As String = "{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}"
Public Const FOLDERID_LocalAppData              As String = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
Public Const FOLDERID_LocalAppDataLow           As String = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
Public Const FOLDERID_LocalizedResourcesDir     As String = "{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"
Public Const FOLDERID_Music                     As String = "{4BD8D571-6D19-48D3-BE97-422220080E43}"
Public Const FOLDERID_MusicLibrary              As String = "{2112AB0A-C86A-4FFE-A368-0DE96E47012E}"
Public Const FOLDERID_NetHood                   As String = "{C5ABBF53-E17F-4121-8900-86626FC2C973}"
Public Const FOLDERID_NetworkFolder             As String = "{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}"
Public Const FOLDERID_OriginalImages            As String = "{2C36C0AA-5812-4B87-BFD0-4CD0DFB19B39}"
Public Const FOLDERID_PhotoAlbums               As String = "{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"
Public Const FOLDERID_PicturesLibrary           As String = "{A990AE9F-A03B-4E80-94BC-9912D7504104}"
Public Const FOLDERID_Pictures                  As String = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
Public Const FOLDERID_Playlists                 As String = "{DE92C1C7-837F-4F69-A3BB-86E631204A23}"
Public Const FOLDERID_PrintersFolder            As String = "{76FC4E2D-D6AD-4519-A663-37BD56068185}"
Public Const FOLDERID_PrintHood                 As String = "{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"
Public Const FOLDERID_Profile                   As String = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
Public Const FOLDERID_ProgramData               As String = "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"
Public Const FOLDERID_ProgramFiles              As String = "{905E63B6-C1BF-494E-B29C-65B732D3D21A}"
Public Const FOLDERID_ProgramFilesX64           As String = "{6D809377-6AF0-444B-8957-A3773F02200E}"
Public Const FOLDERID_ProgramFilesX86           As String = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
Public Const FOLDERID_ProgramFilesCommon        As String = "{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"
Public Const FOLDERID_ProgramFilesCommonX64     As String = "{6365D5A7-0F0D-45E5-87F6-0DA56B6A4F7D}"
Public Const FOLDERID_ProgramFilesCommonX86     As String = "{DE974D24-D9C6-4D3E-BF91-F4455120B917}"
Public Const FOLDERID_Programs                  As String = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"
Public Const FOLDERID_Public                    As String = "{DFDF76A2-C82A-4D63-906A-5644AC457385}"
Public Const FOLDERID_PublicDesktop             As String = "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"
Public Const FOLDERID_PublicDocuments           As String = "{ED4824AF-DCE4-45A8-81E2-FC7965083634}"
Public Const FOLDERID_PublicDownloads           As String = "{3D644C9B-1FB8-4F30-9B45-F670235F79C0}"
Public Const FOLDERID_PublicGameTasks           As String = "{DEBF2536-E1A8-4C59-B6A2-414586476AEA}"
Public Const FOLDERID_PublicLibraries           As String = "{48DAF80B-E6CF-4F4E-B800-0E69D84EE384}"
Public Const FOLDERID_PublicMusic               As String = "{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"
Public Const FOLDERID_PublicPictures            As String = "{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"
Public Const FOLDERID_PublicRingtones           As String = "{E555AB60-153B-4D17-9F04-A5FE99FC15EC}"
Public Const FOLDERID_PublicUserTiles           As String = "{0482AF6C-08F1-4C34-8C90-E17EC98B1E17}"
Public Const FOLDERID_PublicVideos              As String = "{2400183A-6185-49FB-A2D8-4A392A602BA3}"
Public Const FOLDERID_QuickLaunch               As String = "{52A4F021-7B75-48A9-9F6B-4B87A210BC8F}"
Public Const FOLDERID_Recent                    As String = "{AE50C081-EBD2-438A-8655-8A092E34987A}"
'Public Const FOLDERID_RecordedTV                As String = "{BD85E001-112E-431E-983B-7B15AC09FFF1}"   ' Undefined as of Windows 7
Public Const FOLDERID_RecordedTVLibrary         As String = "{1A6FDBA2-F42D-4358-A798-B74D745926C5}"
Public Const FOLDERID_RecycleBinFolder          As String = "{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}"
Public Const FOLDERID_ResourceDir               As String = "{8AD10C31-2ADB-4296-A8F7-E4701232C972}"
Public Const FOLDERID_Ringtones                 As String = "{C870044B-F49E-4126-A9C3-B52A1FF411E8}"
Public Const FOLDERID_RoamingAppData            As String = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
Public Const FOLDERID_RoamedTileImages          As String = "{AAA8D5A5-F1D6-4259-BAA8-78E7EF60835E}"
Public Const FOLDERID_RoamingTiles              As String = "{00BCFC5A-ED94-4e48-96A1-3F6217F21990}"
Public Const FOLDERID_SampleMusic               As String = "{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"
Public Const FOLDERID_SamplePictures            As String = "{C4900540-2379-4C75-844B-64E6FAF8716B}"
Public Const FOLDERID_SamplePlaylists           As String = "{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"
Public Const FOLDERID_SampleVideos              As String = "{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"
Public Const FOLDERID_SavedGames                As String = "{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}"
Public Const FOLDERID_SavedSearches             As String = "{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}"
Public Const FOLDERID_Screenshots               As String = "{B7BEDE81-DF94-4682-A7D8-57A52620B86F}"
Public Const FOLDERID_SEARCH_CSC                As String = "{EE32E446-31CA-4ABA-814F-A5EBD2FD6D5E}"
Public Const FOLDERID_SearchHistory             As String = "{0D4C3DB6-03A3-462F-A0E6-08924C41B5D4}"
Public Const FOLDERID_SearchHome                As String = "{190337D1-B8CA-4121-A639-6D472D16972A}"
Public Const FOLDERID_SEARCH_MAPI               As String = "{98EC0E18-2098-4D44-8644-66979315A281}"
Public Const FOLDERID_SearchTemplates           As String = "{7E636BFE-DFA9-4D5E-B456-D7B39851D8A9}"
Public Const FOLDERID_SendTo                    As String = "{8983036C-27C0-404B-8F08-102D10DCFD74}"
Public Const FOLDERID_SidebarDefaultParts       As String = "{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"
Public Const FOLDERID_SidebarParts              As String = "{A75D362E-50FC-4FB7-AC2C-A8BEAA314493}"
Public Const FOLDERID_SkyDrive                  As String = "{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}"
Public Const FOLDERID_SkyDriveCameraRoll        As String = "{767E6811-49CB-4273-87C2-20F355E1085B}"
Public Const FOLDERID_SkyDriveDocuments         As String = "{24D89E24-2F19-4534-9DDE-6A6671FBB8FE}"
Public Const FOLDERID_SkyDrivePictures          As String = "{339719B5-8C47-4894-94C2-D8F77ADD44A6}"
Public Const FOLDERID_StartMenu                 As String = "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"
Public Const FOLDERID_Startup                   As String = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"
Public Const FOLDERID_SyncManagerFolder         As String = "{43668BF8-C14E-49B2-97C9-747784D784B7}"
Public Const FOLDERID_SyncResultsFolder         As String = "{289A9A43-BE44-4057-A41B-587A76D7E7F9}"
Public Const FOLDERID_SyncSetupFolder           As String = "{0F214138-B1D3-4A90-BBA9-27CBC0C5389A}"
Public Const FOLDERID_System                    As String = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
Public Const FOLDERID_SystemX86                 As String = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"
Public Const FOLDERID_Templates                 As String = "{A63293E8-664E-48DB-A079-DF759E0509F7}"
'Public Const FOLDERID_TreeProperties            As String = "{5B3749AD-B49F-49C1-83EB-15370FBD4882}"   ' Unsupported as of Windows 7
Public Const FOLDERID_UserPinned                As String = "{9E3995AB-1F9C-4F13-B827-48B24B6C7174}"
Public Const FOLDERID_UserProfiles              As String = "{0762D272-C50A-4BB0-A382-697DCD729B80}"
Public Const FOLDERID_UserProgramFiles          As String = "{5CD7AEE2-2219-4A67-B85D-6C9CE15660CB}"
Public Const FOLDERID_UserProgramFilesCommon    As String = "{BCBD3057-CA5C-4622-B42D-BC56DB0AE516}"
Public Const FOLDERID_UsersFiles                As String = "{F3CE0F7C-4901-4ACC-8648-D5D44B04EF8F}"
Public Const FOLDERID_UsersLibraries            As String = "{A302545D-DEFF-464b-ABE8-61C8648D939B}"
Public Const FOLDERID_Videos                    As String = "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"
Public Const FOLDERID_VideosLibrary             As String = "{491E922F-5643-4AF4-A7EB-4E7A138D8174}"
Public Const FOLDERID_Windows                   As String = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"

' ----------------
' Module Level
' ----------------

Private Const msMODULE                          As String = "MWinAPIShell32Folder"

' WinError.h
' -----------------------------------
Private Const NOERROR                           As Long = &H0           ' The CLSID was obtained successfully.
Private Const CO_E_CLASSSTRING                  As Long = &H800401F3    ' The class string was improperly formatted.
Private Const REGDB_E_CLASSNOTREG               As Long = &H80040154    ' The CLSID corresponding to the class string was not found in the registry.
Private Const REGDB_E_READREGDB                 As Long = &H80040150    ' The registry could not be opened for reading.

' CSIDL (constant special item ID list) values are described on MSDN at
' http://msdn.microsoft.com/en-us/library/bb762494(v=vs.85).aspx
' -----------------------------------
Private Const CSIDL_DESKTOP                     As Long = &H0   '  0
Private Const CSIDL_INTERNET                    As Long = &H1   '  1
Private Const CSIDL_PROGRAMS                    As Long = &H2   '  2
Private Const CSIDL_CONTROLS                    As Long = &H3   '  3
Private Const CSIDL_PRINTERS                    As Long = &H4   '  4
Private Const CSIDL_MYDOCUMENTS                 As Long = &H5   '  5
Private Const CSIDL_PERSONAL                    As Long = CSIDL_MYDOCUMENTS
Private Const CSIDL_FAVORITES                   As Long = &H6   '  6
Private Const CSIDL_STARTUP                     As Long = &H7   '  7
Private Const CSIDL_RECENT                      As Long = &H8   '  8
Private Const CSIDL_SENDTO                      As Long = &H9   '  9
Private Const CSIDL_BITBUCKET                   As Long = &HA   ' 10
Private Const CSIDL_STARTMENU                   As Long = &HB   ' 11
Private Const CSIDL_MYMUSIC                     As Long = &HD   ' 13
Private Const CSIDL_MYVIDEO                     As Long = &HE   ' 14
Private Const CSIDL_DESKTOPDIRECTORY            As Long = &H10  ' 16
Private Const CSIDL_DRIVES                      As Long = &H11  ' 17
Private Const CSIDL_NETWORK                     As Long = &H12  ' 18
Private Const CSIDL_NETHOOD                     As Long = &H13  ' 19
Private Const CSIDL_FONTS                       As Long = &H14  ' 20
Private Const CSIDL_TEMPLATES                   As Long = &H15  ' 21
Private Const CSIDL_COMMON_STARTMENU            As Long = &H16  ' 22
Private Const CSIDL_COMMON_PROGRAMS             As Long = &H17  ' 23
Private Const CSIDL_COMMON_STARTUP              As Long = &H18  ' 24
Private Const CSIDL_COMMON_DESKTOPDIRECTORY     As Long = &H19  ' 25
Private Const CSIDL_APPDATA                     As Long = &H1A  ' 26
Private Const CSIDL_PRINTHOOD                   As Long = &H1B  ' 27
Private Const CSIDL_LOCAL_APPDATA               As Long = &H1C  ' 28
Private Const CSIDL_ALTSTARTUP                  As Long = &H1D  ' 29
Private Const CSIDL_COMMON_ALTSTARTUP           As Long = &H1E  ' 30
Private Const CSIDL_COMMON_FAVORITES            As Long = &H1F  ' 31
Private Const CSIDL_INTERNET_CACHE              As Long = &H20  ' 32
Private Const CSIDL_COOKIES                     As Long = &H21  ' 33
Private Const CSIDL_HISTORY                     As Long = &H22  ' 34
Private Const CSIDL_COMMON_APPDATA              As Long = &H23  ' 35
Private Const CSIDL_WINDOWS                     As Long = &H24  ' 36
Private Const CSIDL_SYSTEM                      As Long = &H25  ' 37
Private Const CSIDL_PROGRAM_FILES               As Long = &H26  ' 38
Private Const CSIDL_MYPICTURES                  As Long = &H27  ' 39
Private Const CSIDL_PROFILE                     As Long = &H28  ' 40
Private Const CSIDL_SYSTEMX86                   As Long = &H29  ' 41
Private Const CSIDL_PROGRAM_FILESX86            As Long = &H2A  ' 42
Private Const CSIDL_PROGRAM_FILES_COMMON        As Long = &H2B  ' 43
Private Const CSIDL_PROGRAM_FILES_COMMONX86     As Long = &H2C  ' 44
Private Const CSIDL_COMMON_TEMPLATES            As Long = &H2D  ' 45
Private Const CSIDL_COMMON_DOCUMENTS            As Long = &H2E  ' 46
Private Const CSIDL_COMMON_ADMINTOOLS           As Long = &H2F  ' 47
Private Const CSIDL_ADMINTOOLS                  As Long = &H30  ' 48
Private Const CSIDL_CONNECTIONS                 As Long = &H31  ' 49
Private Const CSIDL_COMMON_MUSIC                As Long = &H35  ' 53
Private Const CSIDL_COMMON_PICTURES             As Long = &H36  ' 54
Private Const CSIDL_COMMON_VIDEO                As Long = &H37  ' 55
Private Const CSIDL_RESOURCES                   As Long = &H38  ' 56
Private Const CSIDL_RESOURCES_LOCALIZED         As Long = &H39  ' 57
Private Const CSIDL_COMMON_OEM_LINKS            As Long = &H3A  ' 58
Private Const CSIDL_CDBURN_AREA                 As Long = &H3B  ' 59
Private Const CSIDL_COMPUTERSNEARME             As Long = &H3D  ' 61

' CSIDL Equivalents identify which CSIDL IDs correspond to KNOWNFOLDER IDs
' --------------------------------------------------------------------------
Private Const CSIDL_EQ_AccountPictures          As String = "None, value introduced in Windows 8"
Private Const CSIDL_EQ_AddNewPrograms           As String = "None"
Private Const CSIDL_EQ_AdminTools               As String = "CSIDL_ADMINTOOLS"
Private Const CSIDL_EQ_ApplicationShortcuts     As String = "None, value introduced in Windows 8"
Private Const CSIDL_EQ_AppsFolder               As String = "None, value introduced in Windows 8"
Private Const CSIDL_EQ_AppUpdates               As String = "None"
Private Const CSIDL_EQ_CameraRoll               As String = "None, value introduced in Windows 8.1"
Private Const CSIDL_EQ_CDBurning                As String = "CSIDL_CDBURN_AREA"
Private Const CSIDL_EQ_ChangeRemovePrograms     As String = "None"
Private Const CSIDL_EQ_CommonAdminTools         As String = "CSIDL_COMMON_ADMINTOOLS"
Private Const CSIDL_EQ_CommonOEMLinks           As String = "CSIDL_COMMON_OEM_LINKS"
Private Const CSIDL_EQ_CommonPrograms           As String = "CSIDL_COMMON_PROGRAMS"
Private Const CSIDL_EQ_CommonStartMenu          As String = "CSIDL_COMMON_STARTMENU"
Private Const CSIDL_EQ_CommonStartup            As String = "CSIDL_COMMON_STARTUP, CSIDL_COMMON_ALTSTARTUP"
Private Const CSIDL_EQ_CommonTemplates          As String = "CSIDL_COMMON_TEMPLATES"
Private Const CSIDL_EQ_ComputerFolder           As String = "CSIDL_DRIVES"
Private Const CSIDL_EQ_ConflictFolder           As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_ConnectionsFolder        As String = "CSIDL_CONNECTIONS"
Private Const CSIDL_EQ_Contacts                 As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_ControlPanelFolder       As String = "CSIDL_CONTROLS"
Private Const CSIDL_EQ_Cookies                  As String = "CSIDL_COOKIES"
Private Const CSIDL_EQ_Desktop                  As String = "CSIDL_DESKTOP, CSIDL_DESKTOPDIRECTORY"
Private Const CSIDL_EQ_DeviceMetadataStore      As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_Documents                As String = "CSIDL_MYDOCUMENTS, CSIDL_PERSONAL"
Private Const CSIDL_EQ_DocumentsLibrary         As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_Downloads                As String = "None"
Private Const CSIDL_EQ_Favorites                As String = "CSIDL_FAVORITES, CSIDL_COMMON_FAVORITES"
Private Const CSIDL_EQ_Fonts                    As String = "CSIDL_FONTS"
Private Const CSIDL_EQ_Games                    As String = "None"
Private Const CSIDL_EQ_GameTasks                As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_History                  As String = "CSIDL_HISTORY"
Private Const CSIDL_EQ_HomeGroup                As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_HomeGroupCurrentUser     As String = "None, value introduced in Windows 8"
Private Const CSIDL_EQ_ImplicitAppShortcuts     As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_InternetCache            As String = "CSIDL_INTERNET_CACHE"
Private Const CSIDL_EQ_InternetFolder           As String = "CSIDL_INTERNET"
Private Const CSIDL_EQ_Libraries                As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_Links                    As String = "None"
Private Const CSIDL_EQ_LocalAppData             As String = "None"
Private Const CSIDL_EQ_LocalAppDataLow          As String = "None"
Private Const CSIDL_EQ_LocalizedResourcesDir    As String = "CSIDL_RESOURCES_LOCALIZED"
Private Const CSIDL_EQ_Music                    As String = "CSIDL_MYMUSIC"
Private Const CSIDL_EQ_MusicLibrary             As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_NetHood                  As String = "CSIDL_NETHOOD"
Private Const CSIDL_EQ_NetworkFolder            As String = "CSIDL_NETWORK, CSIDL_COMPUTERSNEARME"
Private Const CSIDL_EQ_OriginalImages           As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_PhotoAlbums              As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_PicturesLibrary          As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_Pictures                 As String = "None"
Private Const CSIDL_EQ_Playlists                As String = "None"
Private Const CSIDL_EQ_PrintersFolder           As String = "CSIDL_PRINTERS"
Private Const CSIDL_EQ_PrintHood                As String = "CSIDL_PRINTHOOD"
Private Const CSIDL_EQ_Profile                  As String = "CSIDL_PROFILE"
Private Const CSIDL_EQ_ProgramData              As String = "CSIDL_COMMON_APPDATA"
Private Const CSIDL_EQ_ProgramFiles             As String = "CSIDL_PROGRAM_FILES"
Private Const CSIDL_EQ_ProgramFilesX64          As String = "None"
Private Const CSIDL_EQ_ProgramFilesX86          As String = "CSIDL_PROGRAM_FILESX86"
Private Const CSIDL_EQ_ProgramFilesCommon       As String = "CSIDL_PROGRAM_FILES_COMMON"
Private Const CSIDL_EQ_ProgramFilesCommonX64    As String = "None"
Private Const CSIDL_EQ_ProgramFilesCommonX86    As String = "CSIDL_PROGRAM_FILES_COMMONX86"
Private Const CSIDL_EQ_Programs                 As String = "None"
Private Const CSIDL_EQ_Public                   As String = "None"
Private Const CSIDL_EQ_PublicDesktop            As String = "CSIDL_COMMON_DESKTOPDIRECTORY"
Private Const CSIDL_EQ_PublicDocuments          As String = "CSIDL_COMMON_DOCUMENTS"
Private Const CSIDL_EQ_PublicDownloads          As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_PublicGameTasks          As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_PublicLibraries          As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_PublicMusic              As String = "CSIDL_COMMON_MUSIC"
Private Const CSIDL_EQ_PublicPictures           As String = "CSIDL_COMMON_PICTURES"
Private Const CSIDL_EQ_PublicRingtones          As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_PublicUserTiles          As String = "None, value introduced in Windows 8"
Private Const CSIDL_EQ_PublicVideos             As String = "CSIDL_COMMON_VIDEO"
Private Const CSIDL_EQ_Private                  As String = "None"
Private Const CSIDL_EQ_QuickLaunch              As String = "None"
Private Const CSIDL_EQ_Recent                   As String = "CSIDL_RECENT"
'Private Const CSIDL_EQ_RecordedTV               As String = "None"
Private Const CSIDL_EQ_RecordedTVLibrary        As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_RecycleBinFolder         As String = "CSIDL_BITBUCKET"
Private Const CSIDL_EQ_ResourceDir              As String = "CSIDL_RESOURCES"
Private Const CSIDL_EQ_Ringtones                As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_RoamingAppData           As String = "CSIDL_APPDATA"
Private Const CSIDL_EQ_RoamedTileImages         As String = "None, value introduced in Windows 8"
Private Const CSIDL_EQ_RoamingTiles             As String = "None, value introduced in Windows 8"
Private Const CSIDL_EQ_SampleMusic              As String = "None"
Private Const CSIDL_EQ_SamplePictures           As String = "None"
Private Const CSIDL_EQ_SamplePlaylists          As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_SampleVideos             As String = "None"
Private Const CSIDL_EQ_SavedGames               As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_SavedSearches            As String = "None"
Private Const CSIDL_EQ_Screenshots              As String = "None, value introduced in Windows 8"
Private Const CSIDL_EQ_SEARCH_CSC               As String = "None"
Private Const CSIDL_EQ_SearchHistory            As String = "None, value introduced in Windows 8.1"
Private Const CSIDL_EQ_SearchHome               As String = "None"
Private Const CSIDL_EQ_SEARCH_MAPI              As String = "None"
Private Const CSIDL_EQ_SearchTemplates          As String = "None, value introduced in Windows 8.1"
Private Const CSIDL_EQ_SendTo                   As String = "CSIDL_SENDTO"
Private Const CSIDL_EQ_SidebarDefaultParts      As String = "None, new for Windows 7"
Private Const CSIDL_EQ_SidebarParts             As String = "None, new for Windows 7"
Private Const CSIDL_EQ_SkyDrive                 As String = "None, value introduced in Windows 8.1"
Private Const CSIDL_EQ_SkyDriveCameraRoll       As String = "None, value introduced in Windows 8.1"
Private Const CSIDL_EQ_SkyDriveDocuments        As String = "None, value introduced in Windows 8.1"
Private Const CSIDL_EQ_SkyDrivePictures         As String = "None, value introduced in Windows 8.1"
Private Const CSIDL_EQ_StartMenu                As String = "CSIDL_STARTMENU"
Private Const CSIDL_EQ_Startup                  As String = "CSIDL_STARTUP, CSIDL_ALTSTARTUP"
Private Const CSIDL_EQ_SyncManagerFolder        As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_SyncResultsFolder        As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_SyncSetupFolder          As String = "None, value introduced in Windows Vista"
Private Const CSIDL_EQ_System                   As String = "CSIDL_SYSTEM"
Private Const CSIDL_EQ_SystemX86                As String = "CSIDL_SYSTEMX86"
Private Const CSIDL_EQ_Templates                As String = "CSIDL_TEMPLATES"
'Private Const CSIDL_EQ_TreeProperties           As String = "None"
Private Const CSIDL_EQ_UserPinned               As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_UserProfiles             As String = "None, new for Windows Vista"
Private Const CSIDL_EQ_UserProgramFiles         As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_UserProgramFilesCommon   As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_UsersFiles               As String = "None"
Private Const CSIDL_EQ_UsersLibraries           As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_Videos                   As String = "CSIDL_MYVIDEO"
Private Const CSIDL_EQ_VideosLibrary            As String = "None, value introduced in Windows 7"
Private Const CSIDL_EQ_Windows                  As String = "CSIDL_WINDOWS"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuKnownFolder
    kfMyDesktop
    kfMyDocuments
    kfMyPictures
    kfMyTemplates
    kfMyVideos
    kfProgramFilesCommon
    kfPublicDesktop
    kfPublicDocuments
    kfPublicTemplates
    kfPublicPictures
    kfPublicVideos
End Enum

Public Enum enuShellFolder
    sfProgramData = CSIDL_COMMON_APPDATA
    sfAppDataLocal = CSIDL_LOCAL_APPDATA
    sfAppDataRoaming = CSIDL_APPDATA
    sfCommonDesktop = CSIDL_COMMON_DESKTOPDIRECTORY
    sfCommonDocuments = CSIDL_COMMON_DOCUMENTS
    sfCommonFavorites = CSIDL_COMMON_FAVORITES
    sfCommonMusic = CSIDL_COMMON_MUSIC
    sfCommonPictures = CSIDL_COMMON_PICTURES
    sfCommonTemplates = CSIDL_COMMON_TEMPLATES
    sfCommonVideo = CSIDL_COMMON_VIDEO
    sfMyDesktop = CSIDL_DESKTOP
    sfMyDocuments = CSIDL_MYDOCUMENTS
    sfMyFavorites = CSIDL_FAVORITES
    sfMyMusic = CSIDL_MYMUSIC
    sfMyPictures = CSIDL_MYPICTURES
    sfMyTemplates = CSIDL_TEMPLATES
    sfMyVideo = CSIDL_MYVIDEO
    sfProfile = CSIDL_PROFILE
    sfProgramFiles = CSIDL_PROGRAM_FILES
    sfProgramFilesCommon = CSIDL_PROGRAM_FILES_COMMON
    sfWindows = CSIDL_WINDOWS
End Enum

' ----------------
' Module Level
' ----------------

Private Enum enuGetFolderPathFlag
    SHGFP_TYPE_CURRENT = 0
    SHGFP_TYPE_DEFAULT = 1
End Enum

Private Enum enuKnownFolderFlag
    KF_FLAG_SIMPLE_IDLIST = &H100
    KF_FLAG_NOT_PARENT_RELATIVE = &H200
    KF_FLAG_DEFAULT_PATH = &H400
    KF_FLAG_INIT = &H800
    KF_FLAG_NO_ALIAS = &H1000
    KF_FLAG_DONT_UNEXPAND = &H2000
    KF_FLAG_DONT_VERIFY = &H4000
    KF_FLAG_CREATE = &H8000
    KF_FLAG_NO_APPCONTAINER_REDIRECTION = &H10000
    KF_FLAG_ALIAS_ONLY = &H80000000
End Enum

' -----------------------------------
' External Function declarations
' -----------------------------------

' The CLSIDFromString function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms680589(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function CLSIDFromString _
            Lib "ole32" (ByVal lpsz As LongPtr, _
                         ByRef pclsid As Any) _
            As LongPtr
#Else
    Private Declare _
            Function CLSIDFromString _
            Lib "ole32" (ByVal lpsz As Long, _
                         ByRef pclsid As Any) _
            As Long
#End If

' The CopyMemory function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/aa366535(VS.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Sub CopyMemory _
            Lib "Kernel32" _
            Alias "RtlMoveMemory" (pDest As Any, _
                                   pSource As Any, _
                                   ByVal dwLength As Long)
#Else
    Private Declare _
            Sub CopyMemory _
            Lib "Kernel32" _
            Alias "RtlMoveMemory" (pDest As Any, _
                                   pSource As Any, _
                                   ByVal dwLength As Long)
#End If

' The CoTaskMemFree function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms680722(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Sub CoTaskMemFree _
            Lib "ole32" (ByVal pv As LongPtr)
#Else
    Private Declare _
            Sub CoTaskMemFree _
            Lib "ole32" (ByVal pv As Long)
#End If

' The lstrlen function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms647492(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function lstrlen _
            Lib "Kernel32" _
            Alias "lstrlenW" (ByVal lpString As LongPtr) _
            As Long
#Else
    Private Declare _
            Function lstrlen _
            Lib "Kernel32" _
            Alias "lstrlenW" (ByVal lpString As Long) _
            As Long
#End If

' The SHGetFolderPath function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/bb762181(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SHGetFolderPath _
            Lib "shfolder" _
            Alias "SHGetFolderPathA" (ByVal hWndOwner As LongPtr, _
                                      ByVal nFolder As enuShellFolder, _
                                      ByVal hToken As LongPtr, _
                                      ByVal dwFlags As enuGetFolderPathFlag, _
                                      ByVal pszPath As String) _
            As LongPtr
#Else
    Private Declare _
            Function SHGetFolderPath _
            Lib "shfolder" _
            Alias "SHGetFolderPathA" (ByVal hWndOwner As Long, _
                                      ByVal nFolder As enuShellFolder, _
                                      ByVal hToken As Long, _
                                      ByVal dwFlags As enuGetFolderPathFlag, _
                                      ByVal pszPath As String) _
            As Long
#End If

' The SHGetKnownFolderPath function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/bb762188(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function SHGetKnownFolderPath _
            Lib "shell32" (ByRef rfid As Any, _
                           ByVal dwFlags As Long, _
                           ByVal hToken As LongPtr, _
                           ByRef ppszPath As LongPtr) _
            As LongPtr
#Else
    Private Declare _
            Function SHGetKnownFolderPath _
            Lib "shell32" (ByRef rfid As Any, _
                           ByVal dwFlags As Long, _
                           ByVal hToken As Long, _
                           ByRef ppszPath As Long) _
            As Long
#End If

Private Function GetKnownFolderGUID(ByVal FolderId As enuKnownFolder) _
        As String
' ==========================================================================
' Description : Convert an enumeration to a GUID string
'
' Parameters  : FolderId    The enumerated value to convert
'
' Returns     : String      The corresponding GUID
' ==========================================================================

    Const sPROC As String = "GetKnownFolderGUID"

    Dim sRtn    As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    
    Select Case FolderId
    Case kfMyDesktop
        sRtn = FOLDERID_Desktop
    Case kfMyDocuments
        sRtn = FOLDERID_Documents
    Case kfMyPictures
        sRtn = FOLDERID_Pictures
    Case kfMyTemplates
        sRtn = FOLDERID_Templates
    Case kfMyVideos
        sRtn = FOLDERID_Videos

    Case kfProgramFilesCommon
        sRtn = FOLDERID_ProgramFilesCommon

    Case kfPublicDesktop
        sRtn = FOLDERID_PublicDesktop
    Case kfPublicDocuments
        sRtn = FOLDERID_PublicDocuments
    Case kfPublicPictures
        sRtn = FOLDERID_PublicPictures
    Case kfPublicTemplates
        sRtn = FOLDERID_CommonTemplates
    Case kfPublicVideos
        sRtn = FOLDERID_PublicVideos

    Case Else
        sRtn = FOLDERID_Desktop
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetKnownFolderGUID = sRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
      Stop
      Resume
    Else
      Resume PROC_EXIT
    End If

End Function

#If VBA7 Then
Private Function GetStringFromPointer(ByVal lpsz As LongPtr) As String
#Else
Private Function GetStringFromPointer(ByVal lpsz As Long) As String
#End If
' ==========================================================================
' Description : Return the string located at a memory pointer
'
' Parameters  : lpsz        Pointer to a null-terminated string
'
' Returns     : String      The string at the pointer location
' ==========================================================================

    Const sPROC     As String = "GetStringFromPointer"

    Dim bytRtn()    As Byte
    Dim lLen        As Long
    Dim sRtn        As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (lpsz <> 0) Then
        ' Get the size of the unicode string
        ' ----------------------------------
        lLen = lstrlen(lpsz) * 2

        If (lLen <> 0) Then
            ' Create an array of bytes
            ' then copy it to the variable
            ' ----------------------------
            ReDim bytRtn(0 To (lLen - 1)) As Byte
            Call CopyMemory(bytRtn(0), ByVal lpsz, lLen)
            sRtn = bytRtn
        End If
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetStringFromPointer = sRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function GUIDToCSIDLEquivalent(ByVal GUID As String) As String
' ==========================================================================
' Description : Return the CSIDL equivalent for a known folder GUID
'
' Parameters  : GUID        The known folder GUID to evaluate
'
' Returns     : String
'
' Comments    : The return value is descriptive in nature
' ==========================================================================

    Const sPROC As String = "GUIDToCSIDLEquivalent"

    Dim sRtn    As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    
    Select Case GUID
    Case FOLDERID_AccountPictures
        sRtn = CSIDL_EQ_AccountPictures

    Case FOLDERID_AddNewPrograms
        sRtn = CSIDL_EQ_AddNewPrograms

    Case FOLDERID_AdminTools
        sRtn = CSIDL_EQ_AdminTools

    Case FOLDERID_ApplicationShortcuts
        sRtn = CSIDL_EQ_ApplicationShortcuts

    Case FOLDERID_AppsFolder
        sRtn = CSIDL_EQ_AppsFolder

    Case FOLDERID_AppUpdates
        sRtn = CSIDL_EQ_AppUpdates

    Case FOLDERID_CameraRoll
        sRtn = CSIDL_EQ_CameraRoll

    Case FOLDERID_CDBurning
        sRtn = CSIDL_EQ_CDBurning

    Case FOLDERID_ChangeRemovePrograms
        sRtn = CSIDL_EQ_ChangeRemovePrograms

    Case FOLDERID_CommonAdminTools
        sRtn = CSIDL_EQ_CommonAdminTools

    Case FOLDERID_CommonOEMLinks
        sRtn = CSIDL_EQ_CommonOEMLinks

    Case FOLDERID_CommonPrograms
        sRtn = CSIDL_EQ_CommonPrograms

    Case FOLDERID_CommonStartMenu
        sRtn = CSIDL_EQ_CommonStartMenu

    Case FOLDERID_CommonStartup
        sRtn = CSIDL_EQ_CommonStartup

    Case FOLDERID_CommonTemplates
        sRtn = CSIDL_EQ_CommonTemplates

    Case FOLDERID_ComputerFolder
        sRtn = CSIDL_EQ_ComputerFolder

    Case FOLDERID_ConflictFolder
        sRtn = CSIDL_EQ_ConflictFolder

    Case FOLDERID_ConnectionsFolder
        sRtn = CSIDL_EQ_ConnectionsFolder

    Case FOLDERID_Contacts
        sRtn = CSIDL_EQ_Contacts

    Case FOLDERID_ControlPanelFolder
        sRtn = CSIDL_EQ_ControlPanelFolder

    Case FOLDERID_Cookies
        sRtn = CSIDL_EQ_Cookies

    Case FOLDERID_Desktop
        sRtn = CSIDL_EQ_Desktop

    Case FOLDERID_DeviceMetadataStore
        sRtn = CSIDL_EQ_DeviceMetadataStore

    Case FOLDERID_Documents
        sRtn = CSIDL_EQ_Documents

    Case FOLDERID_DocumentsLibrary
        sRtn = CSIDL_EQ_DocumentsLibrary

    Case FOLDERID_Downloads
        sRtn = CSIDL_EQ_Downloads

    Case FOLDERID_Favorites
        sRtn = CSIDL_EQ_Favorites

    Case FOLDERID_Fonts
        sRtn = CSIDL_EQ_Fonts

    Case FOLDERID_Games
        sRtn = CSIDL_EQ_Games

    Case FOLDERID_GameTasks
        sRtn = CSIDL_EQ_GameTasks

    Case FOLDERID_History
        sRtn = CSIDL_EQ_History

    Case FOLDERID_HomeGroup
        sRtn = CSIDL_EQ_HomeGroup

    Case FOLDERID_HomeGroupCurrentUser
        sRtn = CSIDL_EQ_HomeGroupCurrentUser

    Case FOLDERID_ImplicitAppShortcuts
        sRtn = CSIDL_EQ_ImplicitAppShortcuts

    Case FOLDERID_InternetCache
        sRtn = CSIDL_EQ_InternetCache

    Case FOLDERID_InternetFolder
        sRtn = CSIDL_EQ_InternetFolder

    Case FOLDERID_Libraries
        sRtn = CSIDL_EQ_Libraries

    Case FOLDERID_Links
        sRtn = CSIDL_EQ_Links

    Case FOLDERID_LocalAppData
        sRtn = CSIDL_EQ_LocalAppData

    Case FOLDERID_LocalAppDataLow
        sRtn = CSIDL_EQ_LocalAppDataLow

    Case FOLDERID_LocalizedResourcesDir
        sRtn = CSIDL_EQ_LocalizedResourcesDir

    Case FOLDERID_Music
        sRtn = CSIDL_EQ_Music

    Case FOLDERID_MusicLibrary
        sRtn = CSIDL_EQ_MusicLibrary

    Case FOLDERID_NetHood
        sRtn = CSIDL_EQ_NetHood

    Case FOLDERID_NetworkFolder
        sRtn = CSIDL_EQ_NetworkFolder

    Case FOLDERID_OriginalImages
        sRtn = CSIDL_EQ_OriginalImages

    Case FOLDERID_PhotoAlbums
        sRtn = CSIDL_EQ_PhotoAlbums

    Case FOLDERID_PicturesLibrary
        sRtn = CSIDL_EQ_PicturesLibrary

    Case FOLDERID_Pictures
        sRtn = CSIDL_EQ_Pictures

    Case FOLDERID_Playlists
        sRtn = CSIDL_EQ_Playlists

    Case FOLDERID_PrintersFolder
        sRtn = CSIDL_EQ_PrintersFolder

    Case FOLDERID_PrintHood
        sRtn = CSIDL_EQ_PrintHood

    Case FOLDERID_Profile
        sRtn = CSIDL_EQ_Profile

    Case FOLDERID_ProgramData
        sRtn = CSIDL_EQ_ProgramData

    Case FOLDERID_ProgramFiles
        sRtn = CSIDL_EQ_ProgramFiles

    Case FOLDERID_ProgramFilesCommon
        sRtn = CSIDL_EQ_ProgramFilesCommon

    Case FOLDERID_ProgramFilesCommonX64
        sRtn = CSIDL_EQ_ProgramFilesCommonX64

    Case FOLDERID_ProgramFilesCommonX86
        sRtn = CSIDL_EQ_ProgramFilesCommonX86

    Case FOLDERID_ProgramFilesX64
        sRtn = CSIDL_EQ_ProgramFilesX64

    Case FOLDERID_ProgramFilesX86
        sRtn = CSIDL_EQ_ProgramFilesX86

    Case FOLDERID_Programs
        sRtn = CSIDL_EQ_Programs

    Case FOLDERID_Public
        sRtn = CSIDL_EQ_Public

    Case FOLDERID_PublicDesktop
        sRtn = CSIDL_EQ_PublicDesktop

    Case FOLDERID_PublicDocuments
        sRtn = CSIDL_EQ_PublicDocuments

    Case FOLDERID_PublicDownloads
        sRtn = CSIDL_EQ_PublicDownloads

    Case FOLDERID_PublicGameTasks
        sRtn = CSIDL_EQ_PublicGameTasks

    Case FOLDERID_PublicLibraries
        sRtn = CSIDL_EQ_PublicLibraries

    Case FOLDERID_PublicMusic
        sRtn = CSIDL_EQ_PublicMusic

    Case FOLDERID_PublicPictures
        sRtn = CSIDL_EQ_PublicPictures

    Case FOLDERID_PublicRingtones
        sRtn = CSIDL_EQ_PublicRingtones

    Case FOLDERID_PublicUserTiles
        sRtn = CSIDL_EQ_PublicUserTiles

    Case FOLDERID_PublicVideos
        sRtn = CSIDL_EQ_PublicVideos

    Case FOLDERID_QuickLaunch
        sRtn = CSIDL_EQ_QuickLaunch

    Case FOLDERID_Recent
        sRtn = CSIDL_EQ_Recent

'    Case FOLDERID_RecordedTV
'        sRtn = CSIDL_EQ_RecordedTV

    Case FOLDERID_RecordedTVLibrary
        sRtn = CSIDL_EQ_RecordedTVLibrary

    Case FOLDERID_RecycleBinFolder
        sRtn = CSIDL_EQ_RecycleBinFolder

    Case FOLDERID_ResourceDir
        sRtn = CSIDL_EQ_ResourceDir

    Case FOLDERID_Ringtones
        sRtn = CSIDL_EQ_Ringtones

    Case FOLDERID_RoamingAppData
        sRtn = CSIDL_EQ_RoamingAppData

    Case FOLDERID_RoamedTileImages
        sRtn = CSIDL_EQ_RoamedTileImages

    Case FOLDERID_RoamingTiles
        sRtn = CSIDL_EQ_RoamingTiles

    Case FOLDERID_SampleMusic
        sRtn = CSIDL_EQ_SampleMusic

    Case FOLDERID_SamplePictures
        sRtn = CSIDL_EQ_SamplePictures

    Case FOLDERID_SamplePlaylists
        sRtn = CSIDL_EQ_SamplePlaylists

    Case FOLDERID_SampleVideos
        sRtn = CSIDL_EQ_SampleVideos

    Case FOLDERID_SavedGames
        sRtn = CSIDL_EQ_SavedGames

    Case FOLDERID_SavedSearches
        sRtn = CSIDL_EQ_SavedSearches

    Case FOLDERID_Screenshots
        sRtn = CSIDL_EQ_Screenshots

    Case FOLDERID_SEARCH_CSC
        sRtn = CSIDL_EQ_SEARCH_CSC

    Case FOLDERID_SearchHistory
        sRtn = CSIDL_EQ_SearchHistory

    Case FOLDERID_SearchHome
        sRtn = CSIDL_EQ_SearchHome

    Case FOLDERID_SEARCH_MAPI
        sRtn = CSIDL_EQ_SEARCH_MAPI

    Case FOLDERID_SearchTemplates
        sRtn = CSIDL_EQ_SearchTemplates

    Case FOLDERID_SendTo
        sRtn = CSIDL_EQ_SendTo

    Case FOLDERID_SidebarDefaultParts
        sRtn = CSIDL_EQ_SidebarDefaultParts

    Case FOLDERID_SidebarParts
        sRtn = CSIDL_EQ_SidebarParts

    Case FOLDERID_SkyDrive
        sRtn = CSIDL_EQ_SkyDrive

    Case FOLDERID_SkyDriveCameraRoll
        sRtn = CSIDL_EQ_SkyDriveCameraRoll

    Case FOLDERID_SkyDriveDocuments
        sRtn = CSIDL_EQ_SkyDriveDocuments

    Case FOLDERID_SkyDrivePictures
        sRtn = CSIDL_EQ_SkyDrivePictures

    Case FOLDERID_StartMenu
        sRtn = CSIDL_EQ_StartMenu

    Case FOLDERID_Startup
        sRtn = CSIDL_EQ_Startup

    Case FOLDERID_SyncManagerFolder
        sRtn = CSIDL_EQ_SyncManagerFolder

    Case FOLDERID_SyncResultsFolder
        sRtn = CSIDL_EQ_SyncResultsFolder

    Case FOLDERID_SyncSetupFolder
        sRtn = CSIDL_EQ_SyncSetupFolder

    Case FOLDERID_System
        sRtn = CSIDL_EQ_System

    Case FOLDERID_SystemX86
        sRtn = CSIDL_EQ_SystemX86

    Case FOLDERID_Templates
        sRtn = CSIDL_EQ_Templates

'    Case FOLDERID_TreeProperties
'        sRtn = CSIDL_EQ_TreeProperties

    Case FOLDERID_UserPinned
        sRtn = CSIDL_EQ_UserPinned

    Case FOLDERID_UserProfiles
        sRtn = CSIDL_EQ_UserProfiles

    Case FOLDERID_UserProgramFiles
        sRtn = CSIDL_EQ_UserProgramFiles

    Case FOLDERID_UserProgramFilesCommon
        sRtn = CSIDL_EQ_UserProgramFilesCommon

    Case FOLDERID_UsersFiles
        sRtn = CSIDL_EQ_UsersFiles

    Case FOLDERID_UsersLibraries
        sRtn = CSIDL_EQ_UsersLibraries

    Case FOLDERID_Videos
        sRtn = CSIDL_EQ_Videos

    Case FOLDERID_VideosLibrary
        sRtn = CSIDL_EQ_VideosLibrary

    Case FOLDERID_Windows
        sRtn = CSIDL_EQ_Windows

    Case Else
        sRtn = vbNullString

    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GUIDToCSIDLEquivalent = sRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
      Stop
      Resume
    Else
      Resume PROC_EXIT
    End If

End Function

Public Function GUIDToKnownFolderID(ByVal GUID As String) As String
' ==========================================================================
' Description : Return the known folder ID (name) based on the GUID
'
' Parameters  : GUID        The known folder GUID to evaluate
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "GUIDToKnownFolderID"

    Dim sRtn    As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    
    Select Case GUID
    Case FOLDERID_AccountPictures
        sRtn = "FOLDERID_AccountPictures"

    Case FOLDERID_AddNewPrograms
        sRtn = "FOLDERID_AddNewPrograms"

    Case FOLDERID_AdminTools
        sRtn = "FOLDERID_AdminTools"

    Case FOLDERID_ApplicationShortcuts
        sRtn = "FOLDERID_ApplicationShortcuts"

    Case FOLDERID_AppsFolder
        sRtn = "FOLDERID_AppsFolder"

    Case FOLDERID_AppUpdates
        sRtn = "FOLDERID_AppUpdates"

    Case FOLDERID_CameraRoll
        sRtn = "FOLDERID_CameraRoll"

    Case FOLDERID_CDBurning
        sRtn = "FOLDERID_CDBurning"

    Case FOLDERID_ChangeRemovePrograms
        sRtn = "FOLDERID_ChangeRemovePrograms"

    Case FOLDERID_CommonAdminTools
        sRtn = "FOLDERID_CommonAdminTools"

    Case FOLDERID_CommonOEMLinks
        sRtn = "FOLDERID_CommonOEMLinks"

    Case FOLDERID_CommonPrograms
        sRtn = "FOLDERID_CommonPrograms"

    Case FOLDERID_CommonStartMenu
        sRtn = "FOLDERID_CommonStartMenu"

    Case FOLDERID_CommonStartup
        sRtn = "FOLDERID_CommonStartup"

    Case FOLDERID_CommonTemplates
        sRtn = "FOLDERID_CommonTemplates"

    Case FOLDERID_ComputerFolder
        sRtn = "FOLDERID_ComputerFolder"

    Case FOLDERID_ConflictFolder
        sRtn = "FOLDERID_ConflictFolder"

    Case FOLDERID_ConnectionsFolder
        sRtn = "FOLDERID_ConnectionsFolder"

    Case FOLDERID_Contacts
        sRtn = "FOLDERID_Contacts"

    Case FOLDERID_ControlPanelFolder
        sRtn = "FOLDERID_ControlPanelFolder"

    Case FOLDERID_Cookies
        sRtn = "FOLDERID_Cookies"

    Case FOLDERID_Desktop
        sRtn = "FOLDERID_Desktop"

    Case FOLDERID_DeviceMetadataStore
        sRtn = "FOLDERID_DeviceMetadataStore"

    Case FOLDERID_Documents
        sRtn = "FOLDERID_Documents"

    Case FOLDERID_DocumentsLibrary
        sRtn = "FOLDERID_DocumentsLibrary"

    Case FOLDERID_Downloads
        sRtn = "FOLDERID_Downloads"

    Case FOLDERID_Favorites
        sRtn = "FOLDERID_Favorites"

    Case FOLDERID_Fonts
        sRtn = "FOLDERID_Fonts"

    Case FOLDERID_Games
        sRtn = "FOLDERID_Games"

    Case FOLDERID_GameTasks
        sRtn = "FOLDERID_GameTasks"

    Case FOLDERID_History
        sRtn = "FOLDERID_History"

    Case FOLDERID_HomeGroup
        sRtn = "FOLDERID_HomeGroup"

    Case FOLDERID_HomeGroupCurrentUser
        sRtn = "FOLDERID_HomeGroupCurrentUser"

    Case FOLDERID_ImplicitAppShortcuts
        sRtn = "FOLDERID_ImplicitAppShortcuts"

    Case FOLDERID_InternetCache
        sRtn = "FOLDERID_InternetCache"

    Case FOLDERID_InternetFolder
        sRtn = "FOLDERID_InternetFolder"

    Case FOLDERID_Libraries
        sRtn = "FOLDERID_Libraries"

    Case FOLDERID_Links
        sRtn = "FOLDERID_Links"

    Case FOLDERID_LocalAppData
        sRtn = "FOLDERID_LocalAppData"

    Case FOLDERID_LocalAppDataLow
        sRtn = "FOLDERID_LocalAppDataLow"

    Case FOLDERID_LocalizedResourcesDir
        sRtn = "FOLDERID_LocalizedResourcesDir"

    Case FOLDERID_Music
        sRtn = "FOLDERID_Music"

    Case FOLDERID_MusicLibrary
        sRtn = "FOLDERID_MusicLibrary"

    Case FOLDERID_NetHood
        sRtn = "FOLDERID_NetHood"

    Case FOLDERID_NetworkFolder
        sRtn = "FOLDERID_NetworkFolder"

    Case FOLDERID_OriginalImages
        sRtn = "FOLDERID_OriginalImages"

    Case FOLDERID_PhotoAlbums
        sRtn = "FOLDERID_PhotoAlbums"

    Case FOLDERID_PicturesLibrary
        sRtn = "FOLDERID_PicturesLibrary"

    Case FOLDERID_Pictures
        sRtn = "FOLDERID_Pictures"

    Case FOLDERID_Playlists
        sRtn = "FOLDERID_Playlists"

    Case FOLDERID_PrintersFolder
        sRtn = "FOLDERID_PrintersFolder"

    Case FOLDERID_PrintHood
        sRtn = "FOLDERID_PrintHood"

    Case FOLDERID_Profile
        sRtn = "FOLDERID_Profile"

    Case FOLDERID_ProgramData
        sRtn = "FOLDERID_ProgramData"

    Case FOLDERID_ProgramFiles
        sRtn = "FOLDERID_ProgramFiles"

    Case FOLDERID_ProgramFilesCommon
        sRtn = "FOLDERID_ProgramFilesCommon"

    Case FOLDERID_ProgramFilesCommonX64
        sRtn = "FOLDERID_ProgramFilesCommonX64"

    Case FOLDERID_ProgramFilesCommonX86
        sRtn = "FOLDERID_ProgramFilesCommonX86"

    Case FOLDERID_ProgramFilesX64
        sRtn = "FOLDERID_ProgramFilesX64"

    Case FOLDERID_ProgramFilesX86
        sRtn = "FOLDERID_ProgramFilesX86"

    Case FOLDERID_Programs
        sRtn = "FOLDERID_Programs"

    Case FOLDERID_Public
        sRtn = "FOLDERID_Public"

    Case FOLDERID_PublicDesktop
        sRtn = "FOLDERID_PublicDesktop"

    Case FOLDERID_PublicDocuments
        sRtn = "FOLDERID_PublicDocuments"

    Case FOLDERID_PublicDownloads
        sRtn = "FOLDERID_PublicDownloads"

    Case FOLDERID_PublicGameTasks
        sRtn = "FOLDERID_PublicGameTasks"

    Case FOLDERID_PublicLibraries
        sRtn = "FOLDERID_PublicLibraries"

    Case FOLDERID_PublicMusic
        sRtn = "FOLDERID_PublicMusic"

    Case FOLDERID_PublicPictures
        sRtn = "FOLDERID_PublicPictures"

    Case FOLDERID_PublicRingtones
        sRtn = "FOLDERID_PublicRingtones"

    Case FOLDERID_PublicUserTiles
        sRtn = "FOLDERID_PublicUserTiles"

    Case FOLDERID_PublicVideos
        sRtn = "FOLDERID_PublicVideos"

    Case FOLDERID_QuickLaunch
        sRtn = "FOLDERID_QuickLaunch"

    Case FOLDERID_Recent
        sRtn = "FOLDERID_Recent"

'    Case FOLDERID_RecordedTV
'        sRtn = "FOLDERID_RecordedTV"

    Case FOLDERID_RecordedTVLibrary
        sRtn = "FOLDERID_RecordedTVLibrary"

    Case FOLDERID_RecycleBinFolder
        sRtn = "FOLDERID_RecycleBinFolder"

    Case FOLDERID_ResourceDir
        sRtn = "FOLDERID_ResourceDir"

    Case FOLDERID_Ringtones
        sRtn = "FOLDERID_Ringtones"

    Case FOLDERID_RoamingAppData
        sRtn = "FOLDERID_RoamingAppData"

    Case FOLDERID_RoamedTileImages
        sRtn = "FOLDERID_RoamedTileImages"

    Case FOLDERID_RoamingTiles
        sRtn = "FOLDERID_RoamingTiles"

    Case FOLDERID_SampleMusic
        sRtn = "FOLDERID_SampleMusic"

    Case FOLDERID_SamplePictures
        sRtn = "FOLDERID_SamplePictures"

    Case FOLDERID_SamplePlaylists
        sRtn = "FOLDERID_SamplePlaylists"

    Case FOLDERID_SampleVideos
        sRtn = "FOLDERID_SampleVideos"

    Case FOLDERID_SavedGames
        sRtn = "FOLDERID_SavedGames"

    Case FOLDERID_SavedSearches
        sRtn = "FOLDERID_SavedSearches"

    Case FOLDERID_Screenshots
        sRtn = "FOLDERID_Screenshots"

    Case FOLDERID_SEARCH_CSC
        sRtn = "FOLDERID_SEARCH_CSC"

    Case FOLDERID_SearchHistory
        sRtn = "FOLDERID_SearchHistory"

    Case FOLDERID_SearchHome
        sRtn = "FOLDERID_SearchHome"

    Case FOLDERID_SEARCH_MAPI
        sRtn = "FOLDERID_SEARCH_MAPI"

    Case FOLDERID_SearchTemplates
        sRtn = "FOLDERID_SearchTemplates"

    Case FOLDERID_SendTo
        sRtn = "FOLDERID_SendTo"

    Case FOLDERID_SidebarDefaultParts
        sRtn = "FOLDERID_SidebarDefaultParts"

    Case FOLDERID_SidebarParts
        sRtn = "FOLDERID_SidebarParts"

    Case FOLDERID_SkyDrive
        sRtn = "FOLDERID_SkyDrive"

    Case FOLDERID_SkyDriveCameraRoll
        sRtn = "FOLDERID_SkyDriveCameraRoll"

    Case FOLDERID_SkyDriveDocuments
        sRtn = "FOLDERID_SkyDriveDocuments"

    Case FOLDERID_SkyDrivePictures
        sRtn = "FOLDERID_SkyDrivePictures"

    Case FOLDERID_StartMenu
        sRtn = "FOLDERID_StartMenu"

    Case FOLDERID_Startup
        sRtn = "FOLDERID_Startup"

    Case FOLDERID_SyncManagerFolder
        sRtn = "FOLDERID_SyncManagerFolder"

    Case FOLDERID_SyncResultsFolder
        sRtn = "FOLDERID_SyncResultsFolder"

    Case FOLDERID_SyncSetupFolder
        sRtn = "FOLDERID_SyncSetupFolder"

    Case FOLDERID_System
        sRtn = "FOLDERID_System"

    Case FOLDERID_SystemX86
        sRtn = "FOLDERID_SystemX86"

    Case FOLDERID_Templates
        sRtn = "FOLDERID_Templates"

'    Case FOLDERID_TreeProperties
'        sRtn = "FOLDERID_TreeProperties"

    Case FOLDERID_UserPinned
        sRtn = "FOLDERID_UserPinned"

    Case FOLDERID_UserProfiles
        sRtn = "FOLDERID_UserProfiles"

    Case FOLDERID_UserProgramFiles
        sRtn = "FOLDERID_UserProgramFiles"

    Case FOLDERID_UserProgramFilesCommon
        sRtn = "FOLDERID_UserProgramFilesCommon"

    Case FOLDERID_UsersFiles
        sRtn = "FOLDERID_UsersFiles"

    Case FOLDERID_UsersLibraries
        sRtn = "FOLDERID_UsersLibraries"

    Case FOLDERID_Videos
        sRtn = "FOLDERID_Videos"

    Case FOLDERID_VideosLibrary
        sRtn = "FOLDERID_VideosLibrary"

    Case FOLDERID_Windows
        sRtn = "FOLDERID_Windows"

    Case Else
        sRtn = vbNullString

    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GUIDToKnownFolderID = sRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
      Stop
      Resume
    Else
      Resume PROC_EXIT
    End If

End Function

Public Function ShellGetFolderPath(ByVal Folder As enuShellFolder) As String
' ==========================================================================
' Description : Locate a folder from the Windows Shell.
'
' Parameters  : Folder    The shell folder to locate
'
' Returns     : String
' ==========================================================================

    Const sPROC     As String = "ShellGetFolderPath"

    #If VBA7 Then
        Dim lRtn    As LongPtr
    #Else
        Dim lRtn    As Long
    #End If

    Dim sRtn        As String
    Dim sPath       As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Folder)

    ' ----------------------------------------------------------------------

    sPath = String$(MAX_PATH, 0)
    lRtn = SHGetFolderPath(0&, Folder, 0&, SHGFP_TYPE_CURRENT, sPath)

    Select Case lRtn
    Case S_OK
        sRtn = TrimToNull(sPath) & "\"
    Case S_FALSE
        Stop
    Case E_INVALIDARG
        Stop
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ShellGetFolderPath = sRtn

    Call Trace(tlMaximum, msMODULE, sPROC, sRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function ShellGetKnownFolderByGUID(ByVal GUID As String) As String
' ==========================================================================
' Description : Return the known folder using a GUID
'
' Parameters  : GUID        The GUID for the known folder
'
' Returns     : String
' ==========================================================================

    Const sPROC         As String = "ShellGetKnownFolderByGUID"

    Dim sRtn            As String

    #If VBA7 Then
        Dim lhResult    As LongPtr
        Dim lpsz        As LongPtr
    #Else
        Dim lhResult    As Long
        Dim lpsz        As Long
    #End If

    Dim udtGUID         As TGUID


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    If (CLSIDFromString(StrPtr(GUID), udtGUID) = NOERROR) Then
        ' Get a pointer to the Unicode
        ' string identified by the GUID
        ' -----------------------------
        lhResult = SHGetKnownFolderPath(udtGUID, 0, 0, lpsz)

        If (lhResult = S_OK) Then
            ' Copy the string from memory
            ' and release the pointer
            ' ---------------------------
            sRtn = GetStringFromPointer(lpsz) & "\"
            Call CoTaskMemFree(lpsz)
        End If
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ShellGetKnownFolderByGUID = sRtn
    
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
      Stop
      Resume
    Else
      Resume PROC_EXIT
    End If

End Function

Public Function ShellGetKnownFolderPath(ByVal KnownFolder _
                                           As enuKnownFolder) As String
' ==========================================================================
' Description : Locate a known folder.
'
' Parameters  : FolderId    An enumerated value that identifies the GUID
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "ShellGetKnownFolderPath"

    Dim sGUID   As String
    Dim sRtn    As String

    
    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Convert the KnownFolder ID to a GUID
    ' ------------------------------------
    sGUID = GetKnownFolderGUID(KnownFolder)
    sRtn = ShellGetKnownFolderByGUID(sGUID)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ShellGetKnownFolderPath = sRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function
