Attribute VB_Name = "modPublic"
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Public Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Any) As Long
Public Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hModule As Long, ByVal lpName As Any, ByVal lpType As Any) As Long
Public Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As Long, ByVal lpType As Any, ByVal lpName As Any, ByVal wLanguage As Long) As Long
Public Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Public Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
' Predefined Resource Types
Public Const RT_CURSOR = 1&
Public Const RT_BITMAP = 2&
Public Const RT_ICON = 3&
Public Const RT_MENU = 4&
Public Const RT_DIALOG = 5&
Public Const RT_STRING = 6&
Public Const RT_FONTDIR = 7&
Public Const RT_FONT = 8&
Public Const RT_ACCELERATOR = 9&
Public Const RT_RCDATA = 10&
Public Const RT_HTML = 23&
'API language values
'--------------------------------------------------------------------------------
'0x0000 Language Neutral
'0x007f The language for the invariant locale (LOCALE_INVARIANT). See MAKELCID.
'0x0400 Process or User Default Language
'0x0800 System Default Language
'0x0436 Afrikaans
'0x041c Albanian
'0x0401 Arabic (Saudi Arabia)
'0x0801 Arabic (Iraq)
'0x0c01 Arabic (Egypt)
'0x1001 Arabic (Libya)
'0x1401 Arabic (Algeria)
'0x1801 Arabic (Morocco)
'0x1c01 Arabic (Tunisia)
'0x2001 Arabic (Oman)
'0x2401 Arabic (Yemen)
'0x2801 Arabic (Syria)
'0x2c01 Arabic (Jordan)
'0x3001 Arabic (Lebanon)
'0x3401 Arabic (Kuwait)
'0x3801 Arabic (U.A.E.)
'0x3c01 Arabic (Bahrain)
'0x4001 Arabic (Qatar)
'0x042b Windows 2000 or later: Armenian. This is Unicode only.
'0x042c Azeri (Latin)
'0x082c Azeri (Cyrillic)
'0x042d Basque
'0x0423 Belarussian
'0x0402 Bulgarian
'0x0455 Burmese
'0x0403 Catalan
'0x0404 Chinese (Taiwan)
'0x0804 Chinese (PRC)
'0x0c04 Chinese (Hong Kong SAR, PRC)
'0x1004 Chinese (Singapore)
'0x1404 Windows 98/Me, Windows 2000 or later: Chinese (Macau SAR)
'0x041a Croatian
'0x0405 Czech
'0x0406 Danish
'0x0465 Whistler: Divehi. This is Unicode only.
'0x0413 Dutch (Netherlands)
'0x0813 Dutch (Belgium)
'0x0409 English (United States)
'0x0809 English (United Kingdom)
'0x0c09 English (Australian)
'0x1009 English (Canadian)
'0x1409 English (New Zealand)
'0x1809 English (Ireland)
'0x1c09 English (South Africa)
'0x2009 English (Jamaica)
'0x2409 English (Caribbean)
'0x2809 English (Belize)
'0x2c09 English (Trinidad)
'0x3009 Windows 98/Me, Windows 2000 or later: English (Zimbabwe)
'0x3409 Windows 98/Me, Windows 2000 or later: English (Philippines)
'0x0425 Estonian
'0x0438 Faeroese
'0x0429 Farsi
'0x040b Finnish
'0x040c French (Standard)
'0x080c French (Belgian)
'0x0c0c French (Canadian)
'0x100c French (Switzerland)
'0x140c French (Luxembourg)
'0x180c Windows 98/Me, Windows 2000 or later: French (Monaco)
'0x0456 Whistler: Galician
'0x0437 Windows 2000 and later: Georgian. This is Unicode only.
'0x0407 German (Standard)
'0x0807 German (Switzerland)
'0x0c07 German (Austria)
'0x1007 German (Luxembourg)
'0x1407 German (Liechtenstein)
'0x0408 Greek
'0x0447 Whistler: Gujarati. This is Unicode only.
'0x040d Hebrew
'0x0439 Windows 2000 and later: Hindi. This is Unicode only.
'0x040e Hungarian
'0x040f Icelandic
'0x0421 Indonesian
'0x0410 Italian (Standard)
'0x0810 Italian (Switzerland)
'0x0411 Japanese
'0x044b Whistler: Kannada. This is Unicode only.
'0x0860 Kashmiri
'0x043f Kazakh
'0x0457 Windows 2000 and later: Konkani. This is Unicode only.
'0x0412 Korean
'0x0812 Windows 95, Windows NT 4.0 only: Korean (Johab)
'0x0440 Whistler: Kyrgyz.
'0x0426 Latvian
'0x0427 Lithuanian
'0x0827 Windows 98 only: Lithuanian (Classic)
'0x042f FYRO Macedonian
'0x043e Malay (Malaysian)
'0x083e Malay (Brunei Darussalam)
'0x0458 Manipuri
'0x044e Windows 2000 and later: Marathi. This is Unicode only.
'0x0450 Whistler: Mongolian
'0x0414 Norwegian (Bokmal)
'0x0814 Norwegian (Nynorsk)
'0x0415 Polish
'0x0416 Portuguese (Brazil)
'0x0816 Portuguese (Portugal)
'0x0446 Whistler: Punjabi. This is Unicode only.
'0x0418 Romanian
'0x0419 Russian
'0x044f Windows 2000 and later: Sanskrit. This is Unicode only.
'0x0c1a Serbian (Cyrillic)
'0x081a Serbian (Latin)
'0x0459 Sindhi
'0x041b Slovak
'0x0424 Slovenian
'0x040a Spanish (Traditional Sort)
'0x080a Spanish (Mexican)
'0x0c0a Spanish (Modern Sort)
'0x100a Spanish (Guatemala)
'0x140a Spanish (Costa Rica)
'0x180a Spanish (Panama)
'0x1c0a Spanish (Dominican Republic)
'0x200a Spanish (Venezuela)
'0x240a Spanish (Colombia)
'0x280a Spanish (Peru)
'0x2c0a Spanish (Argentina)
'0x300a Spanish (Ecuador)
'0x340a Spanish (Chile)
'0x380a Spanish (Uruguay)
'0x3c0a Spanish (Paraguay)
'0x400a Spanish (Bolivia)
'0x440a Spanish (El Salvador)
'0x480a Spanish (Honduras)
'0x4c0a Spanish (Nicaragua)
'0x500a Spanish (Puerto Rico)
'0x0430 Sutu
'0x0441 Swahili (Kenya)
'0x041d Swedish
'0x081d Swedish (Finland)
'0x045a Whistler: Syriac. This is Unicode only.
'0x0449 Windows 2000 and later: Tamil. This is Unicode only.
'0x0444 Tatar (Tatarstan)
'0x044a Whistler: Telugu. This is Unicode only.
'0x041e Thai
'0x041f Turkish
'0x0422 Ukrainian
'0x0420 Windows 98/Me, Windows 2000 or later: Urdu (Pakistan)
'0x0820 Urdu (India)
'0x0443 Uzbek (Latin)
'0x0843 Uzbek (Cyrillic)
'0x042a Windows 98/Me, Windows NT 4.0 and later: Vietnamese

'created by Florian Santi 25/9/2001
'return the resource string matching with the ID and the language ID (see above)
Public Function LoadResString(Library As String, ResourceID As Long, LanguageID As Long) As String
    On Error GoTo Err_LoadResString
    
    Dim hRes As Long
    Dim hInst As Long
    Dim lngResID As Long
    Dim lngLanguage As Long
    Dim hResInfo As Long
    Dim lpData As Long
    Dim lngArraySize As Long
    Dim bytData() As Byte
    Dim lngPos As Long, lngID As Long, lngLength As Long
    Dim strBuffer As String
    
    LoadResString = ""                                                  'set the default value
    lngResID = Int(ResourceID / 16 + 1)                                 'the string resources are stored by block of 16
    lngLanguage = LanguageID                                            'set the local language ID
    hInst = LoadLibrary(Library)                                        'get the instance of the EXE or the DLL
    hRes = FindResourceEx(hInst, ByVal CLng(RT_STRING), _
        ByVal lngResID, ByVal lngLanguage)                              'search for the resource
    If hRes > 0 Then                                                    'if the resource is found
        hResInfo = LoadResource(hInst, hRes)                            'load the resource in memory
        lpData = LockResource(hResInfo)                                 'lock the resource in memory
        lngArraySize = SizeofResource(hInst, hRes)                      'get the size in byte of the resource
        If lngArraySize > 0 Then                                        'if the size is greater than 0
            ReDim bytData(lngArraySize - 1)                             'redim an array of byte with the size -1 (null terminating string)
            CopyMemory bytData(0), ByVal lpData, lngArraySize           'copy from global memory to the local array of byte
            For lngID = ((lngResID - 1) * 16) To (lngResID * 16 - 1)    'for the lock of 16 resource string
                CopyMemory lngLength, bytData(lngPos), 2                'get the size of the string
                If lngLength Then                                       'if the length is greater than 0
                    strBuffer = String(lngLength, 0)                    'prepare the buffer with null string
                    CopyMemory ByVal StrPtr(strBuffer), _
                        bytData(lngPos + 2), lngLength * 2              'copy the string to the buffer
                    If lngID = ResourceID Then                          'if the current resource string ID is matching the argument
                        LoadResString = strBuffer                       'return the value of the buffer
                        Exit For
                    End If                                              'endif the current resource...
                    lngPos = lngPos + lngLength * 2 + 2                 'goto the next position in the array of byte of the resource ID
                Else                                                    'if the length is = 0, no resource string at this position
                    lngPos = lngPos + 2                                 'goto next position in the array of byte of the resource ID
                End If                                                  'end if the length...
            Next lngID                                                  'next resource ID
        End If                                                          'end if the size of the resource is greater than 0
        FreeResource hResInfo                                           'free the resource
    End If                                                              'end if the resource id found
    FreeLibrary ByVal hInst                                             'free the library
    
Exit_LoadResString:
    Exit Function
Err_LoadResString:
    MsgBox CStr(Err.Number) + ":" + Err.Description, vbOKOnly + vbCritical, App.Title
    Resume Exit_LoadResString
End Function

