Attribute VB_Name = "MMSOutlookConstants"
' ==========================================================================
' Module      : MMSOutlookConstants
' Type        : Module
' Description :
' --------------------------------------------------------------------------
' Dependencies: XXX
' --------------------------------------------------------------------------
' Comments    : THIS MODULE IS IN PROGRESS
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE As String = "MMSOutlookConstants"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum OlActionCopyLike
    olReply = 0
    olReplyAll = 1
    olForward = 2
    olReplyFolder = 3
    olRespond = 4
End Enum

Public Enum OlActionReplyStyle
    olOmitOriginalText = 0
    olEmbedOriginalItem = 1
    olIncludeOriginalText = 2
    olIndentOriginalText = 3
    olLinkOriginalItem = 4
    olUserPreference = 5
    olReplyTickOriginalText = 1000
End Enum

Public Enum OlActionResponseStyle
    olOpen = 0
    olSend = 1
    olPrompt = 2
End Enum

Public Enum OlActionShowOn
    olDontShow = 0
    olMenu = 1
    olMenuAndToolbar = 2
End Enum

Public Enum OlAttachmentType
    olByValue = 1
    olByReference = 4
    olEmbeddeditem = 5
    olOLE = 6
End Enum

Public Enum OlBodyFormat
    olFormatUnspecified = 0
    olFormatPlain = 1
    olFormatHTML = 2
    olFormatRichText = 3
End Enum

Public Enum OlBusyStatus
    olFree = 0
    olTentative = 1
    olBusy = 2
    olOutOfOffice = 3
End Enum

Public Enum OlConnectionMode
    olOffline = 100
    olLowBandwidth = 200
    olOnline = 300
End Enum

Public Enum OlDaysOfWeek
    olSunday = 1
    olMonday = 2
    olTuesday = 4
    olWednesday = 8
    olThursday = 16
    olFriday = 32
    olSaturday = 64
End Enum

Public Enum OlDefaultFolders
    olFolderDeletedItems = 3
    olFolderOutbox = 4
    olFolderSentMail = 5
    olFolderInbox = 6
    olFolderCalendar = 9
    olFolderContacts = 10
    olFolderJournal = 11
    olFolderNotes = 12
    olFolderTasks = 13
    olFolderDrafts = 16
    olPublicFoldersAllPublicFolders = 18
    olFolderConflicts = 19
    olFolderSyncIssues = 20
    olFolderLocalFailures = 21
    olFolderServerFailures = 22
    olFolderJunk = 23
End Enum

Public Enum OlDisplayType
    olUser = 0
    olDistList = 1
    olForum = 2
    olAgent = 3
    olOrganization = 4
    olPrivateDistList = 5
    olRemoteUser = 6
End Enum

Public Enum OlDownloadState
    olHeaderOnly = 0
    olFullItem = 1
End Enum

Public Enum OlEditorType
    olEditorText = 1
    olEditorHTML = 2
    olEditorRTF = 3
    olEditorWord = 4
End Enum

Public Enum OlExchangeConnectionMode
    olNoExchange = 0
    olOffline = 100
    olCachedOffline = 200
    olDisconnected = 300
    olCachedDisconnected = 400
    olCachedConnectedHeaders = 500
    olCachedConnectedDrizzle = 600
    olCachedConnectedFull = 700
    olOnline = 800
End Enum

Public Enum OlFlagIcon
    olNoFlagIcon
    olPurpleFlagIcon
    olOrangeFlagIcon
    olGreenFlagIcon
    olYellowFlagIcon
    olBlueFlagIcon
    olRedFlagIcon
End Enum

Public Enum OlFlagStatus
    olNoFlag = 0
    olFlagComplete = 1
    olFlagMarked = 2
End Enum

Public Enum OlFolderDisplayMode
    olFolderDisplayNormal = 0
    olFolderDisplayFolderOnly = 1
    olFolderDisplayNoNavigation = 2
End Enum

Public Enum OlFormRegistry
    olDefaultRegistry = 0
    olPersonalRegistry = 2
    olFolderRegistry = 3
    olOrganizationRegistry = 4
End Enum

Public Enum OlGender
    olUnspecified = 0
    olFemale = 1
    olMale = 2
End Enum

Public Enum OlImportance
    olImportanceLow = 0
    olImportanceNormal = 1
    olImportanceHigh = 2
End Enum

Public Enum OlInspectorClose
    olSave = 0
    olDiscard = 1
    olPromptForSave = 2
End Enum

Public Enum OlItemType
    olMailItem = 0
    olAppointmentItem = 1
    olContactItem = 2
    olTaskItem = 3
    olJournalItem = 4
    olNoteItem = 5
    olPostItem = 6
    olDistributionListItem = 7
End Enum

Public Enum OlJournalRecipientType
    olAssociatedContact = 1
End Enum


Public Enum OlMailingAddress
    olNone = 0
    olHome = 1
    olBusiness = 2
    olOther = 3
End Enum

Public Enum OlMailRecipientType
    olOriginator = 0
    olTo = 1
    olCC = 2
    olBCC = 3
End Enum

Public Enum OlMeetingRecipientType
    olOrganizer = 0
    olRequired = 1
    olOptional = 2
    olResource = 3
End Enum

Public Enum OlMeetingResponse
    olMeetingTentative = 2
    olMeetingAccepted = 3
    olMeetingDeclined = 4
End Enum

Public Enum OlMeetingStatus
    olNonMeeting = 0
    olMeeting = 1
    olMeetingReceived = 3
    olMeetingCanceled = 5
End Enum

Public Enum OlNetMeetingType
    olNetMeeting = 0
    olNetShow = 1
    olExchangeConferencing = 2
End Enum

Public Enum OlNoteColor
    olBlue = 0
    olGreen = 1
    olPink = 2
    olYellow = 3
    olWhite = 4
End Enum

Public Enum OlObjectClass
    olApplication = 0
    olNamespace = 1
    olFolder = 2
    olRecipient = 4
    olAttachment = 5
    olAddressList = 7
    olAddressEntry = 8
    olFolders = 15
    olItems = 16
    olRecipients = 17
    olAttachments = 18
    olAddressLists = 20
    olAddressEntries = 21
    olAppointment = 26
    olRecurrencePattern = 28
    olExceptions = 29
    olException = 30
    olAction = 32
    olActions = 33
    olExplorer = 34
    olInspector = 35
    olPages = 36
    olFormDescription = 37
    olUserProperties = 38
    olUserProperty = 39
    olContact = 40
    olDocument = 41
    olJournal = 42
    olMail = 43
    olNote = 44
    olPost = 45
    olReport = 46
    olRemote = 47
    olTask = 48
    olTaskRequest = 49
    olTaskRequestUpdate = 50
    olTaskRequestAccept = 51
    olTaskRequestDecline = 52
    olMeetingRequest = 53
    olMeetingCancellation = 54
    olMeetingResponseNegative = 55
    olMeetingResponsePositive = 56
    olMeetingResponseTentative = 57
    olExplorers = 60
    olInspectors = 61
    olPanes = 62
    olOutlookBarPane = 63
    olOutlookBarStorage = 64
    olOutlookBarGroups = 65
    olOutlookBarGroup = 66
    olOutlookBarShortcuts = 67
    olOutlookBarShortcut = 68
    olDistributionList = 69
    olPropertyPageSite = 70
    olPropertyPages = 71
    olSyncObject = 72
    olSyncObjects = 73
    olSelection = 74
    olLink = 75
    olLinks = 76
    olSearch = 77
    olResults = 78
    olViews = 79
    olView = 80
    olItemProperties = 98
    olItemProperty = 99
    olReminders = 100
    olReminder = 101
    olConflict = 117
    olConflicts = 118
End Enum

Public Enum OlOfficeDocItemsType
    olExcelWorkSheetItem = 8
    olWordDocumentItem = 9
    olPowerPointShowItem = 10
End Enum

Public Enum OlOutlookBarViewType
    olLargeIcon = 0
    olSmallIcon = 1
End Enum


Public Enum OlPane
    olOutlookBar = 1
    olFolderList = 2
    olPreview = 3
    olNavigationPane = 4
End Enum

Public Enum OlPermission
    olUnrestricted = 0
    olDoNotForward = 1
    olPermissionTemplate = 2
End Enum

Public Enum OlPermissionService
    olUnknown = 0
    olWindows = 1
    olPassport = 2
End Enum


Public Enum OlRecurrenceState
    olApptNotRecurring = 0
    olApptMaster = 1
    olApptOccurrence = 2
    olApptException = 3
End Enum

Public Enum OlRecurrenceType
    olRecursDaily = 0
    olRecursWeekly = 1
    olRecursMonthly = 2
    olRecursMonthNth = 3
    olRecursYearly = 5
    olRecursYearNth = 6
End Enum

Public Enum OlRemoteStatus
    olRemoteStatusNone = 0
    olUnMarked = 1
    olMarkedForDownload = 2
    olMarkedForCopy = 3
    olMarkedForDelete = 4
End Enum

Public Enum OlResponseStatus
    olResponseNone = 0
    olResponseOrganized = 1
    olResponseTentative = 2
    olResponseAccepted = 3
    olResponseDeclined = 4
    olResponseNotResponded = 5
End Enum

Public Enum OlSaveAsType
    olTXT = 0
    olRTF = 1
    olTemplate = 2
    olMSG = 3
    olDoc = 4
    olHTML = 5
    olVCard = 6
    olVCal = 7
    olICal = 8
    olMSGUnicode = 9
End Enum

Public Enum OlSensitivity
    olNormal = 0
    olPersonal = 1
    olPrivate = 2
    olConfidential = 3
End Enum

Public Enum OlShowItemCount
    olNoItemCount = 0
    olShowUnreadItemCount = 1
    olShowTotalItemCount = 2
End Enum

Public Enum OlSortOrder
    olSortNone = 0
    olAscending = 1
    olDescending = 2
End Enum

Public Enum OlStoreType
    olStoreDefault = 1
    olStoreUnicode = 2
    olStoreANSI = 3
End Enum

Public Enum OlSyncState
    olSyncStopped = 0
    olSyncStarted = 1
End Enum

Public Enum OlTaskDelegationState
    olTaskNotDelegated = 0
    olTaskDelegationUnknown = 1
    olTaskDelegationAccepted = 2
    olTaskDelegationDeclined = 3
End Enum

Public Enum OlTaskOwnership
    olNewTask = 0
    olDelegatedTask = 1
    olOwnTask = 2
End Enum

Public Enum OlTaskRecipientType
    olUpdate = 2
    olFinalStatus = 3
End Enum

Public Enum OlTaskResponse
    olTaskSimple = 0
    olTaskAssign = 1
    olTaskAccept = 2
    olTaskDecline = 3
End Enum

Public Enum OlTaskStatus
    olTaskNotStarted = 0
    olTaskInProgress = 1
    olTaskComplete = 2
    olTaskWaiting = 3
    olTaskDeferred = 4
End Enum

Public Enum OlTrackingStatus
    olTrackingNone = 0
    olTrackingDelivered = 1
    olTrackingNotDelivered = 2
    olTrackingNotRead = 3
    olTrackingRecallFailure = 4
    olTrackingRecallSuccess = 5
    olTrackingRead = 6
    olTrackingReplied = 7
End Enum

Public Enum OlUserPropertyType
    olOutlookInternal = 0
    olText = 1
    olNumber = 3
    olDateTime = 5
    olYesNo = 6
    olDuration = 7
    olKeywords = 11
    olPercent = 12
    olCurrency = 14
    olFormula = 18
    olCombination = 19
End Enum

Public Enum OlViewSaveOption
    olViewSaveOptionThisFolderEveryone = 0
    olViewSaveOptionThisFolderOnlyMe = 1
    olViewSaveOptionAllFoldersOfType = 2
End Enum

Public Enum OlViewType
    olTableView = 0
    olCardView = 1
    olCalendarView = 2
    olIconView = 3
    olTimelineView = 4
End Enum

Public Enum OlWindowState
    olMaximized = 0
    olMinimized = 1
    olNormalWindow = 2
End Enum
