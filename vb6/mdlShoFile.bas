Attribute VB_Name = "mdlShoFile"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                  '
' MarkInfo - A Player Database management program for Spellbinder Games' Kickabout '
' Copyright (C) 2000-2001  Tim Dodge                                               '
'                                                                                  '
' This program is free software; you can redistribute it and/or                    '
' modify it under the terms of the GNU General Public License                      '
' as published by the Free Software Foundation; either version 2                   '
' of the License, or (at your option) any later version.                           '
'                                                                                  '
' This program is distributed in the hope that it will be useful,                  '
' but WITHOUT ANY WARRANTY; without even the implied warranty of                   '
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                    '
' GNU General Public License for more details.                                     '
'                                                                                  '
' You should have received a copy of the GNU General Public License                '
' along with this program; if not, write to the Free Software                      '
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.      '
'                                                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Constants
Private Const cm_TEMPFOLDER_s As String = "\markinfo"
Private Const cm_ZIP_s As String = "zip"
Private Const cm_TXT_s As String = "txt"
Private Const cm_FORMATSTART_s As String = "<$"
Private Const cm_FORMATEND_s As String = ">"
Private Const cm_INTSQUAD_s As String = "INTERNATIONAL SQUAD"
Private Const cm_REPSQUAD_s As String = "REPRESENTATIVE SQUAD"
Private Const cm_AUCTION_s As String = "AUCTION RESULTS"
Private Const cm_DEALS_s As String = "PRIVATE DEALS"
Private Const cm_SPACE_s As String = " "
Private Const cm_DASH_s As String = "-"
Private Const cm_SKILL_s As String = "S"
Private Const cm_POWER_s As String = "P"
Private Const cm_POWERSKILL_s As String = "P/S"
Private Const cm_KEEPER_s As String = "GK"
Private Const cm_OPENBRACKET_s As String = "("
Private Const cm_CLOSEBRACKET_s As String = ")"
Private Const cm_NOCAT_s As String = "---"
Private Const cm_QUESTION_s As String = "?"
Private Const cm_STAR_s As String = "STAR"
Private Const cm_SBY_s As String = "SBY"
Private Const cm_SWAPPED_s As String = "SWAPPED"
Private Const cm_SOLD_s As String = "SOLD"
Private Const cm_AND_s As String = "&"
Private Const cm_FOR_s As String = "FOR"
Private Const cm_OPENSQUAREB_s As String = "["
Private Const cm_CURLYPS_s As String = "{P/S}"
Private Const cm_CURLYSBY_s As String = "{SBY}"
Private Const cm_CURLYFUT_s As String = "{FUT}"
Private Const cm_CURLYSTAR_s As String = "{STAR}"
Private Const cm_FUT_s As String = "FUT"
Private Const cm_K_s As String = "K"
Private Const cm_TO_s As String = "TO"
Private Const cm_FROM_s As String = "FROM"
Private Const cm_CURLYAPP_s As String = "{APP}"
Private Const cm_APP_s As String = "APP"
Private Const cm_MINUS_s As String = "-"

'name: ProcInputFile
'description: unzips the shomatch file and passes each file on for reading

Public Sub ProcInputFile(ByVal pr_ShoFile_s As String)

Dim vl_FileSystem_o As FileSystemObject
Dim vl_TempFolder_o As Folder
Dim vl_TempFolder_s As String
Dim vl_File_o As File
Dim vl_NumAddns_i As Integer
Dim vl_NumUpdates_i As Integer

On Error GoTo ErrorHandler

    'Initialise
    Set vl_FileSystem_o = New FileSystemObject
    Set vl_TempFolder_o = vl_FileSystem_o.GetSpecialFolder(TemporaryFolder)
    vl_TempFolder_s = vl_TempFolder_o.Path & cm_TEMPFOLDER_s
    vl_NumAddns_i = 0
    vl_NumUpdates_i = 0

    'Delete temp folder if it exists
    If (vl_FileSystem_o.FolderExists(vl_TempFolder_s)) Then
        vl_FileSystem_o.DeleteFolder vl_TempFolder_s
    End If

    'If it's a zip file, unzip and pass all .txt files on for processing
    If (LCase$(vl_FileSystem_o.GetExtensionName(pr_ShoFile_s)) = cm_ZIP_s) Then

        'UNZIP32.DLL Options
        uQuiet = 2              'no output
        uOverWriteFiles = 1     'overwrite files
        uCaseSensitivity = 1    'case insensitive
        uZipFileName = pr_ShoFile_s
        uExtractDir = vl_TempFolder_s
    
        If (VBUnZip32 <> 0) Then
            Err.Raise vbObjectError + 1001, "mdlShoFile.ProcInputFile", "Unzip error!"
        End If

        Set vl_TempFolder_o = vl_FileSystem_o.GetFolder(vl_TempFolder_s)

        'Loop through the files in vl_TempFolder_o, passing them to ProcShoFile
        For Each vl_File_o In vl_TempFolder_o.Files
            If (LCase$(vl_FileSystem_o.GetExtensionName(vl_File_o)) = cm_TXT_s) Then
                ProcShoFile vl_File_o.Path, vl_NumAddns_i, vl_NumUpdates_i
            End If
        Next

        'Delete the folder
        vl_TempFolder_o.Delete True

    ElseIf (LCase$(vl_FileSystem_o.GetExtensionName(pr_ShoFile_s)) = cm_TXT_s) Then
        
        ProcShoFile pr_ShoFile_s, vl_NumAddns_i, vl_NumUpdates_i
    
    End If

    'Bring in all the new teams
    PopulateTeams frmMain.cboLeagues.Text

    'Report numbers of updates
    MsgBox vl_NumAddns_i & " players added and " & vl_NumUpdates_i & " updated"

    'clear up
    If Not (vl_File_o Is Nothing) Then
        Set vl_File_o = Nothing
    End If
    
    Set vl_TempFolder_o = Nothing
    Set vl_FileSystem_o = Nothing

    Exit Sub

ErrorHandler:

    'clear up
    If Not (vl_File_o Is Nothing) Then
        Set vl_File_o = Nothing
    End If

    If Not (vl_TempFolder_o Is Nothing) Then
        Set vl_TempFolder_o = Nothing
    End If
    
    If Not (vl_FileSystem_o Is Nothing) Then
        Set vl_FileSystem_o = Nothing
    End If

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "MdlMain.ProcInputFile", Err.Description
    End If

End Sub

'name: ProcShowFile
'description: Processes the individual shomatch files, returns number of updates

Private Sub ProcShoFile(ByVal pr_ShoFile_s As String, _
                        ByRef pu_NumAddns_i As Integer, _
                        ByRef pu_NumUpdates_i As Integer)

Dim vl_FileSystem_o As FileSystemObject
Dim vl_TextStream_o As TextStream
Dim vl_CurrLine_s As String
Dim vl_LineLen_i As Integer
Dim vl_Posn_i As Integer
Dim vl_Index_i As Integer
Dim vl_CurrAsc_i As Integer
Dim vl_Words_s As String
Dim vl_WordArray_s() As String
Dim vl_SquadFile_b As Boolean
Dim vl_AuctionRes_b As Boolean
Dim vl_DealsFile_b As Boolean

On Error GoTo ErrorHandler

    'Open textfile for reading
    Set vl_FileSystem_o = New FileSystemObject
    Set vl_TextStream_o = vl_FileSystem_o.OpenTextFile(pr_ShoFile_s, ForReading)

    'Is it a shomatch file?
    If Not (vl_TextStream_o.AtEndOfStream) Then

        'Initialise flags
        vl_SquadFile_b = False
        vl_AuctionRes_b = False
        vl_DealsFile_b = False

        'Read the first line
        vl_CurrLine_s = UCase$(vl_TextStream_o.ReadLine)
        
        'What kind of file?
        If (InStr(vl_CurrLine_s, cm_INTSQUAD_s) <> 0) Then
            vl_SquadFile_b = True
        ElseIf (InStr(vl_CurrLine_s, cm_REPSQUAD_s) <> 0) Then
            vl_SquadFile_b = True
        ElseIf (InStr(vl_CurrLine_s, cm_AUCTION_s) <> 0) Then
            vl_AuctionRes_b = True
        ElseIf (InStr(vl_CurrLine_s, cm_DEALS_s) <> 0) Then
            vl_DealsFile_b = True
        Else

            'Not a recognised shomatch file - return
            Exit Sub

        End If

    End If

    Do While Not vl_TextStream_o.AtEndOfStream
    
        vl_CurrLine_s = UCase$(vl_TextStream_o.ReadLine)
        
        vl_Posn_i = 1
        vl_LineLen_i = Len(vl_CurrLine_s)
        vl_Words_s = vbNullString
        
        Do While vl_Posn_i <= vl_LineLen_i

            vl_CurrAsc_i = Asc(Mid$(vl_CurrLine_s, vl_Posn_i, 1))

            'Ignore Leading white space
            Do While ((vl_CurrAsc_i = 9 Or vl_CurrAsc_i = 13 Or vl_CurrAsc_i = 32) _
                      And vl_Posn_i <= vl_LineLen_i)
                vl_Posn_i = vl_Posn_i + 1
                If (vl_Posn_i <= vl_LineLen_i) Then
                    vl_CurrAsc_i = Asc(Mid$(vl_CurrLine_s, vl_Posn_i, 1))
                End If
            Loop

            'Ignore <$blah> type formatting
            Do While (Mid$(vl_CurrLine_s, vl_Posn_i, 2) = cm_FORMATSTART_s)

                'Try and find the other end
                vl_Index_i = InStr(vl_Posn_i, vl_CurrLine_s, cm_FORMATEND_s)
                If (vl_Index_i <> 0 And vl_Index_i <= vl_LineLen_i) Then
                    
                    vl_Posn_i = vl_Index_i + 1
                    If (vl_Posn_i <= vl_LineLen_i) Then
                        vl_CurrAsc_i = Asc(Mid$(vl_CurrLine_s, vl_Posn_i, 1))
                    End If
                    
                    'Ignore any white space
                    Do While (vl_CurrAsc_i = 9 Or vl_CurrAsc_i = 13 Or vl_CurrAsc_i = 32 _
                              And vl_Posn_i <= vl_LineLen_i)
                    
                        vl_Posn_i = vl_Posn_i + 1
                        If (vl_Posn_i <= vl_LineLen_i) Then
                            vl_CurrAsc_i = Asc(Mid$(vl_CurrLine_s, vl_Posn_i, 1))
                        End If
                        
                    Loop
                    
                End If
            
            Loop
                    
            Do While Not (vl_CurrAsc_i = 9 Or vl_CurrAsc_i = 13 Or vl_CurrAsc_i = 32) _
                     And (vl_Posn_i <= vl_LineLen_i)
            
                vl_Words_s = vl_Words_s & Mid$(vl_CurrLine_s, vl_Posn_i, 1)
                vl_Posn_i = vl_Posn_i + 1
                If (vl_Posn_i <= vl_LineLen_i) Then
                    vl_CurrAsc_i = Asc(Mid$(vl_CurrLine_s, vl_Posn_i, 1))
                End If
            
            Loop
            
            vl_Words_s = vl_Words_s & cm_SPACE_s
        
        Loop

        'Remove the trailing space
        vl_Words_s = RTrim$(vl_Words_s)

        'Create the word array
        vl_WordArray_s = Split(vl_Words_s)

        'Process the line
        If (UBound(vl_WordArray_s) < 4) Then
            'Ignore line - no player data has less than 5 words
        ElseIf vl_SquadFile_b Then
            ProcSquad vl_WordArray_s, pu_NumAddns_i, pu_NumUpdates_i
        ElseIf vl_AuctionRes_b Then
            ProcAuction vl_WordArray_s, pu_NumAddns_i, pu_NumUpdates_i
        ElseIf vl_DealsFile_b Then
            ProcDeals vl_WordArray_s, pu_NumAddns_i, pu_NumUpdates_i
        End If

    Loop

    'Clear up
    Set vl_FileSystem_o = Nothing
    Set vl_TextStream_o = Nothing
    
    Exit Sub

ErrorHandler:

    'clear up
    If Not (vl_FileSystem_o Is Nothing) Then
        Set vl_FileSystem_o = Nothing
    End If
    
    If Not (vl_TextStream_o Is Nothing) Then
        Set vl_TextStream_o = Nothing
    End If

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "mdlShoFile.ProcShoFile", Err.Description
    End If

End Sub

'name: ProcSquad
'description: Takes a line from the shomatch squad file, creates a player object
'             from it, and adds it to the database. Returns the number of players
'             added or updated.

Private Sub ProcSquad(ByRef pr_Words_s() As String, _
                      ByRef pu_NumAddns_i As Integer, _
                      ByRef pu_NumUpdates_i As Integer)

Dim vl_Words_i As Integer
Dim vl_Index1_i As Integer
Dim vl_Index2_i As Integer
Dim vl_Player_o As Player
Dim vl_Team_s As String
Dim vl_Name_s As String
Dim vl_Age_i As Integer
Dim vl_Level_i As Integer
Dim vl_Cat_s As String
Dim vl_Pos_s As String
Dim vl_Type_s As String
Dim vl_PlayerKey_l As Long

On Error GoTo ErrorHandler

    'How many words?
    vl_Words_i = UBound(pr_Words_s)

    'Work out if this is a player or not
    For vl_Index1_i = 3 To vl_Words_i
        
        If (Mid$(pr_Words_s(vl_Index1_i), 3, 1) = cm_DASH_s And _
                   ((pr_Words_s(vl_Index1_i + 1) = cm_SKILL_s Or _
                     pr_Words_s(vl_Index1_i + 1) = cm_POWER_s Or _
                     pr_Words_s(vl_Index1_i + 1) = cm_POWERSKILL_s) Or _
                     pr_Words_s(vl_Index1_i - 1) = cm_KEEPER_s)) Then
            
            Exit For
        
        End If
    
    Next
    
    'It's a player if we didn't reach the end of the array
    If (vl_Index1_i < vl_Words_i) Then

        'Create the player object
        Set vl_Player_o = New Player

        'Team
        vl_Team_s = vbNullString
        
        If (pr_Words_s(vl_Index1_i - 1) = cm_KEEPER_s And _
            Not (pr_Words_s(vl_Index1_i + 1) = cm_POWER_s Or _
                 pr_Words_s(vl_Index1_i + 1) = cm_SKILL_s Or _
                 pr_Words_s(vl_Index1_i + 1) = cm_POWERSKILL_s)) Then

            For vl_Index2_i = vl_Index1_i + 1 To vl_Words_i
                vl_Team_s = vl_Team_s & pr_Words_s(vl_Index2_i) & cm_SPACE_s
            Next

        Else
        
            For vl_Index2_i = vl_Index1_i + 2 To vl_Words_i
                vl_Team_s = vl_Team_s & pr_Words_s(vl_Index2_i) & cm_SPACE_s
            Next

        End If
        
        vl_Team_s = RTrim$(vl_Team_s)

        'Name
        vl_Name_s = vbNullString
        
        For vl_Index2_i = 0 To vl_Index1_i - 2
        
            If Not (Right$(pr_Words_s(vl_Index2_i), 1) = cm_CLOSEBRACKET_s And _
                    Len(pr_Words_s(vl_Index2_i)) <= 3) Then
                vl_Name_s = vl_Name_s & pr_Words_s(vl_Index2_i) & cm_SPACE_s
            End If
            
        Next
        
        vl_Name_s = RTrim$(vl_Name_s)

        'Age, Level and Pos
        vl_Age_i = CInt(Left$(pr_Words_s(vl_Index1_i), 2))
        vl_Level_i = CInt(Mid$(pr_Words_s(vl_Index1_i), 4))
        vl_Pos_s = pr_Words_s(vl_Index2_i)
        
        'Cat
        If (pr_Words_s(vl_Index1_i - 1) = cm_KEEPER_s) Then
            vl_Cat_s = cm_NOCAT_s
        Else
            vl_Cat_s = pr_Words_s(vl_Index1_i + 1)
        End If

        'Type
        If (vl_Level_i > 12) Then
            vl_Type_s = cm_STAR_s
        ElseIf (vl_Age_i = 17) Then
            vl_Type_s = cm_SBY_s
        Else
            vl_Type_s = cm_NOCAT_s
        End If

        'Add player to database
        vl_Player_o.Create frmMain.cboLeagues.Text, vl_Team_s, vl_Name_s, vl_Age_i, _
                           vl_Level_i, vl_Cat_s, vl_Pos_s, vl_Type_s, 0, vbNullString
        vl_PlayerKey_l = vp_DBManager_o.AddPlayer(vl_Player_o)
        
        If (vl_PlayerKey_l > vp_MaxPlayerKey_l) Then
            vp_MaxPlayerKey_l = vl_PlayerKey_l
            pu_NumAddns_i = pu_NumAddns_i + 1
        Else
            pu_NumUpdates_i = pu_NumUpdates_i + 1
        End If

        'Destroy player object
        Set vl_Player_o = Nothing
        
    End If

    Exit Sub

ErrorHandler:

    'Clear up
    If Not (vl_Player_o Is Nothing) Then
        Set vl_Player_o = Nothing
    End If

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "mdlShoFile.ProcSquad", Err.Description
    End If

End Sub

'name: ProcAuction
'description: Takes a line from the shomatch squad file, creates the necessary player
'             objects from it, and adds them to the database. Returns the number of
'             players added or updated.

Private Sub ProcAuction(ByRef pr_Words_s() As String, _
                        ByRef pu_NumAddns_i As Integer, _
                        ByRef pu_NumUpdates_i As Integer)

Dim vl_Words_i As Integer
Dim vl_Index1_i As Integer
Dim vl_Index2_i As Integer
Dim vl_Player_o As Player
Dim vl_Team_s As String
Dim vl_Name_s As String
Dim vl_Age_i As Integer
Dim vl_Level_i As Integer
Dim vl_Cat_s As String
Dim vl_Pos_s As String
Dim vl_Type_s As String
Dim vl_PlayerKey_l As Long

On Error GoTo ErrorHandler

    'Get number of words
    vl_Words_i = UBound(pr_Words_s)

    'Work out if this is a player or not
    For vl_Index1_i = 2 To vl_Words_i
        If (Mid$(pr_Words_s(vl_Index1_i), 3, 1) = cm_DASH_s And _
            Right$(pr_Words_s(vl_Index1_i - 2), 1) = cm_CLOSEBRACKET_s) Then
            Exit For
        End If
    Next

    'It's a player if we didn't reach the end of the array and the player was bought
    If (vl_Index1_i < vl_Words_i And Not (pr_Words_s(3) = "NOT" And pr_Words_s(4) = "SOLD.")) Then

        'Create the player object
        Set vl_Player_o = New Player

        'Team
        vl_Index2_i = vl_Words_i - 1
        vl_Team_s = Left$(pr_Words_s(vl_Index2_i), Len(pr_Words_s(vl_Index2_i)) - 1)
        
        Do While (Left$(vl_Team_s, 1) <> cm_OPENBRACKET_s)
            vl_Index2_i = vl_Index2_i - 1
            vl_Team_s = pr_Words_s(vl_Index2_i) & cm_SPACE_s & vl_Team_s
        Loop
        
        vl_Team_s = Mid$(vl_Team_s, 2)
        
        'Name
        vl_Name_s = vbNullString
        
        Do While (vl_Index2_i > 3)
            vl_Index2_i = vl_Index2_i - 1
            vl_Name_s = pr_Words_s(vl_Index2_i) & cm_SPACE_s & vl_Name_s
        Loop
        
        vl_Name_s = RTrim$(vl_Name_s)

        'Age, Level and Pos
        vl_Age_i = CInt(Left$(pr_Words_s(2), 2))
        vl_Level_i = CInt(Mid$(pr_Words_s(2), 4))
        vl_Pos_s = pr_Words_s(1)
        
        'Cat
        If (vl_Pos_s = cm_KEEPER_s) Then
            vl_Cat_s = cm_NOCAT_s
        Else
            vl_Cat_s = cm_QUESTION_s
        End If

        'Type
        If (vl_Level_i > 12) Then
            vl_Type_s = cm_STAR_s
        ElseIf (vl_Age_i = 17) Then
            vl_Type_s = cm_SBY_s
        Else
            vl_Type_s = cm_NOCAT_s
        End If

        'Add player to database
        vl_Player_o.Create frmMain.cboLeagues.Text, vl_Team_s, vl_Name_s, vl_Age_i, _
                           vl_Level_i, vl_Cat_s, vl_Pos_s, vl_Type_s, 0, vbNullString
        vl_PlayerKey_l = vp_DBManager_o.AddPlayer(vl_Player_o)
        
        If (vl_PlayerKey_l > vp_MaxPlayerKey_l) Then
            vp_MaxPlayerKey_l = vl_PlayerKey_l
            pu_NumAddns_i = pu_NumAddns_i + 1
        Else
            pu_NumUpdates_i = pu_NumUpdates_i + 1
        End If

        'Destroy player object
        Set vl_Player_o = Nothing
        
    End If

    Exit Sub

ErrorHandler:

    'Clear up
    If Not (vl_Player_o Is Nothing) Then
        Set vl_Player_o = Nothing
    End If

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "mdlShoFile.ProcAuction", Err.Description
    End If

End Sub

'name: ProcDeals
'description: Takes a line from the shomatch squad file, creates the necessary player
'             objects from it, and adds them to the database. Returns the number of
'             players added or updated.

Private Sub ProcDeals(ByRef pr_Words_s() As String, _
                      ByRef pu_NumAddns_i As Integer, _
                      ByRef pu_NumUpdates_i As Integer)

Dim vl_Words_i As Integer
Dim vl_Index1_i As Integer
Dim vl_Index2_i As Integer
Dim vl_Player_o As Player
Dim vl_Squad_o As Squad
Dim vl_League_s As String
Dim vl_OldTeam_s As String
Dim vl_NewTeam_s As String
Dim vl_Name_s As String
Dim vl_Age_i As Integer
Dim vl_Level_i As Integer
Dim vl_Cat_s As String
Dim vl_Pos_s As String
Dim vl_Type_s As String
Dim vl_PlayerKey_l As Long
Dim vl_Switch_b As Boolean

On Error GoTo ErrorHandler

    'Get number of words
    vl_Words_i = UBound(pr_Words_s)

    'Initialise
    vl_Switch_b = False
    vl_League_s = frmMain.cboLeagues.Text
    Set vl_Squad_o = New Squad

    'Try and find the selling club's name
    vl_Index1_i = 0
    Do While (vl_Index1_i < vl_Words_i And _
              pr_Words_s(vl_Index1_i) <> cm_SWAPPED_s And _
              pr_Words_s(vl_Index1_i) <> cm_SOLD_s)
              
        vl_OldTeam_s = vl_OldTeam_s & pr_Words_s(vl_Index1_i) & cm_SPACE_s
        vl_Index1_i = vl_Index1_i + 1
        
    Loop
    vl_OldTeam_s = RTrim$(vl_OldTeam_s)

    'It's a player if we didn't reach the end of the array
    Do While (vl_Index1_i < vl_Words_i And _
              (pr_Words_s(vl_Index1_i) = cm_SWAPPED_s Or _
               pr_Words_s(vl_Index1_i) = cm_SOLD_s Or _
               pr_Words_s(vl_Index1_i) = cm_AND_s Or _
               pr_Words_s(vl_Index1_i) = cm_FOR_s))
    
        vl_Index1_i = vl_Index1_i + 1
    
        'Name
        vl_Name_s = vbNullString
        
        'Step through the word array until we find the age and level word
        'ie something that matches "(nn-"
        Do While (vl_Index1_i < vl_Words_i And _
                  Not (Left$(pr_Words_s(vl_Index1_i), 1) = cm_OPENBRACKET_s And _
                       IsNumeric(Mid$(pr_Words_s(vl_Index1_i), 2, 2)) And _
                       Mid$(pr_Words_s(vl_Index1_i), 4, 1) = cm_MINUS_s))
            vl_Name_s = vl_Name_s & pr_Words_s(vl_Index1_i) & cm_SPACE_s
            vl_Index1_i = vl_Index1_i + 1
        Loop

        vl_Name_s = RTrim$(vl_Name_s)

        'Age and Level
        vl_Age_i = CInt(Mid$(pr_Words_s(vl_Index1_i), 2, 2))
        vl_Level_i = CInt(Mid$(pr_Words_s(vl_Index1_i), _
                               InStr(pr_Words_s(vl_Index1_i), cm_DASH_s) + 1))
        
        'Type
        If (vl_Level_i > 12) Then
            vl_Type_s = cm_STAR_s
        ElseIf (vl_Age_i = 17) Then
            vl_Type_s = cm_SBY_s
        Else
            vl_Type_s = cm_NOCAT_s
        End If
        
        'Pos
        vl_Index1_i = vl_Index1_i + 1
        vl_Index2_i = InStr(pr_Words_s(vl_Index1_i), cm_OPENSQUAREB_s)
        
        If (vl_Index2_i = 0) Then
            vl_Pos_s = Left$(pr_Words_s(vl_Index1_i), Len(pr_Words_s(vl_Index1_i)) - 1)
        Else
            vl_Pos_s = Left$(pr_Words_s(vl_Index1_i), vl_Index2_i - 1)
            vl_Index1_i = vl_Index1_i + 1
        End If
    
        vl_Index1_i = vl_Index1_i + 1
    
        'Cat
        vl_Cat_s = cm_QUESTION_s
        
        If (pr_Words_s(vl_Index1_i) = cm_KEEPER_s) Then
            vl_Cat_s = cm_NOCAT_s
        End If

        If (pr_Words_s(vl_Index1_i) = cm_CURLYPS_s) Then
            vl_Cat_s = cm_POWERSKILL_s
            vl_Index1_i = vl_Index1_i + 1
        End If
        
        If (pr_Words_s(vl_Index1_i) = cm_CURLYSBY_s) Then
            vl_Type_s = cm_SBY_s
            vl_Index1_i = vl_Index1_i + 1
        ElseIf (pr_Words_s(vl_Index1_i) = cm_CURLYFUT_s) Then
            vl_Type_s = cm_FUT_s
            vl_Index1_i = vl_Index1_i + 1
        ElseIf (pr_Words_s(vl_Index1_i) = cm_CURLYAPP_s) Then
            vl_Type_s = cm_APP_s
            vl_Index1_i = vl_Index1_i + 1
        ElseIf (pr_Words_s(vl_Index1_i) = cm_CURLYSTAR_s) Then
            vl_Type_s = cm_STAR_s
            vl_Index1_i = vl_Index1_i + 1
        End If
    
        'Add the Player to the squad
        If vl_Switch_b Then
            vl_Squad_o.Add vl_League_s, vl_OldTeam_s, vl_Name_s, vl_Age_i, _
                           vl_Level_i, vl_Cat_s, vl_Pos_s, vl_Type_s, 0, vbNullString
        Else
            vl_Squad_o.Add vl_League_s, vbNullString, vl_Name_s, vl_Age_i, _
                           vl_Level_i, vl_Cat_s, vl_Pos_s, vl_Type_s, 0, vbNullString
        End If
        
        'End of first half of the deal?
        If (pr_Words_s(vl_Index1_i) = cm_FOR_s) Then
            vl_Switch_b = True
        End If
        
        'Show me the money!
        If (IsNumeric(Left$(pr_Words_s(vl_Index1_i + 1), Len(pr_Words_s(vl_Index1_i + 1)) - 1)) And _
            (Right$(pr_Words_s(vl_Index1_i + 1), 1) = cm_K_s)) Then
            vl_Index1_i = vl_Index1_i + 2
        End If

    Loop
    
    If (pr_Words_s(vl_Index1_i) <> cm_TO_s) Then  'SWAP rather than SALE

        Do While (pr_Words_s(vl_Index1_i) <> cm_FROM_s)
            vl_Index1_i = vl_Index1_i + 1
        Loop
        
        vl_Index1_i = vl_Index1_i + 1
        
        vl_NewTeam_s = vbNullString
        Do While (vl_Index1_i <= vl_Words_i)
            vl_NewTeam_s = vl_NewTeam_s & pr_Words_s(vl_Index1_i) & cm_SPACE_s
            vl_Index1_i = vl_Index1_i + 1
        Loop

        vl_NewTeam_s = Left$(vl_NewTeam_s, Len(vl_NewTeam_s) - 2)

    Else
    
        vl_Index1_i = vl_Index1_i + 1
        
        vl_NewTeam_s = vbNullString
        Do While (pr_Words_s(vl_Index1_i) <> cm_FOR_s)
            vl_NewTeam_s = vl_NewTeam_s & pr_Words_s(vl_Index1_i) & cm_SPACE_s
            vl_Index1_i = vl_Index1_i + 1
        Loop
        
        vl_NewTeam_s = RTrim$(vl_NewTeam_s)

    End If

    'Loop through the squad, updating the database
    For Each vl_Player_o In vl_Squad_o

        'Which way round?
        If (vl_Player_o.PlayerTeam = vbNullString) Then
            vl_Player_o.PlayerTeam = vl_NewTeam_s
            vl_PlayerKey_l = vp_DBManager_o.MovePlayer(vl_Player_o, vl_OldTeam_s)
        Else
            vl_PlayerKey_l = vp_DBManager_o.MovePlayer(vl_Player_o, vl_NewTeam_s)
        End If

        If (vl_PlayerKey_l > vp_MaxPlayerKey_l) Then
            vp_MaxPlayerKey_l = vl_PlayerKey_l
            pu_NumAddns_i = pu_NumAddns_i + 1
        Else
            pu_NumUpdates_i = pu_NumUpdates_i + 1
        End If

    Next

    'Clear up
    Set vl_Squad_o = Nothing
    Set vl_Player_o = Nothing

    Exit Sub

ErrorHandler:

    'Clear up
    If Not (vl_Player_o Is Nothing) Then
        Set vl_Player_o = Nothing
    End If
    
    If Not (vl_Squad_o Is Nothing) Then
        Set vl_Squad_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "mdlShoFile.ProcDeals", Err.Description
    End If

End Sub
