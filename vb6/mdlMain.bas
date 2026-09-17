Attribute VB_Name = "mdlMain"
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

Public vp_DBManager_o As DBManager
Public vp_Squad_o As Squad
Public vp_MaxPlayerKey_l As Long

Public Sub Main()
    
Dim vl_FileSystem_o As FileSystemObject
Dim vl_UDLPath_s As String
Dim vl_MDBPath_s As String
Dim vl_MDBDistPath_s As String
    
On Error GoTo ErrorHandler

    Set vp_DBManager_o = New DBManager
    Set vl_FileSystem_o = New FileSystemObject

    vl_UDLPath_s = App.Path & "\markinfo.udl"

    'If the UDL file exists, use it
    If (vl_FileSystem_o.FileExists(vl_UDLPath_s)) Then
        
        vp_DBManager_o.Create vl_UDLPath_s
    
    Else
        
        'Check that MarkInfoDB.mdb exists
        
        vl_MDBPath_s = App.Path & "\MarkInfo21.mdb"
        vl_MDBDistPath_s = App.Path & "\MarkInfo21.mdb.dist"
        
        If Not (vl_FileSystem_o.FileExists(vl_MDBPath_s)) Then
            
            If (vl_FileSystem_o.FileExists(vl_MDBDistPath_s)) Then
                vl_FileSystem_o.CopyFile vl_MDBDistPath_s, vl_MDBPath_s
            Else
                Err.Raise vbObjectError + 1001, "mdlMain.Main", _
                "Neither MarkInfo21.mdb nor MarkInfo21.mdb.dist exists in application folder"
            End If
        
        End If
        
        vp_DBManager_o.Create
    
    End If

    'Check whether Players table exists
    If Not vp_DBManager_o.TableExists("Players") Then
        Err.Raise vbObjectError + 1001, "mdlMain.Main", "MarkInfo v2 tables not found"
    End If
    
    'Show the main form
    frmMain.txtSession = CStr(vp_DBManager_o.Session)
    frmMain.Show

    'Get the current Maximum PlayerKey
    vp_MaxPlayerKey_l = vp_DBManager_o.GetMaxPlayerKey
    
    'Offer to import v1.0x or v2.0x database if there are no players
    'and old database is available
    If (vp_MaxPlayerKey_l = 0) Then
        If vl_FileSystem_o.FileExists(App.Path & "\MarkInfoDB.mdb") Then
            If (MsgBox("Import players from v2.0x database?", vbYesNo) = vbYes) Then
                ImportData App.Path & "\MarkInfoDB.mdb"
                PopulateLeagues
                Unload frmImport
            End If
        ElseIf vl_FileSystem_o.FileExists(App.Path & "\MarkInfo.mdb") Then
            If (MsgBox("Import players from v1.0x database?", vbYesNo) = vbYes) Then
                ImportData App.Path & "\MarkInfo.mdb"
                PopulateLeagues
                Unload frmImport
            End If
        End If
    End If
    
    'Clear up
    Set vl_FileSystem_o = Nothing

    Exit Sub

ErrorHandler:

    If Not (vl_FileSystem_o Is Nothing) Then
        Set vl_FileSystem_o = Nothing
    End If
    
    If Not (vp_DBManager_o Is Nothing) Then
        Set vp_DBManager_o = Nothing
    End If

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "mdlMain.Main", Err.Description
    End If

End Sub

Public Sub PopulateLeagues()

Dim vl_Leagues_s() As String
Dim vl_Index_i As Integer

On Error GoTo ErrorHandler

    'Empty the combo box
    frmMain.cboLeagues.Clear

    'Get the League Names
    vl_Leagues_s = vp_DBManager_o.GetLeagues
    
    'Add the League Names to the combo box
    For vl_Index_i = LBound(vl_Leagues_s) To UBound(vl_Leagues_s)
        frmMain.cboLeagues.AddItem vl_Leagues_s(vl_Index_i)
    Next vl_Index_i

    'Select the first league
    frmMain.cboLeagues.Text = frmMain.cboLeagues.List(0)

    Exit Sub

ErrorHandler:

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "mdlMain.PopulateLeagues", Err.Description
    End If

End Sub

Public Sub PopulateTeams(ByVal pr_League_s As String)

Dim vl_Teams_s() As String
Dim vl_Index_i As Integer

On Error GoTo ErrorHandler

    'Empty the combo box
    frmMain.cboTeams.Clear
    
    'Get the Team Names
    vl_Teams_s = vp_DBManager_o.GetTeams(pr_League_s)
    
    'Add the Team Names to the combo box
    For vl_Index_i = LBound(vl_Teams_s) To UBound(vl_Teams_s)
        If (vl_Teams_s(vl_Index_i) <> vbNullString) Then
            frmMain.cboTeams.AddItem vl_Teams_s(vl_Index_i)
        End If
    Next vl_Index_i

    'Select the first team
    frmMain.cboTeams.Text = frmMain.cboTeams.List(0)

    Exit Sub

ErrorHandler:

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "mdlMain.PopulateTeams", Err.Description
    End If

End Sub

Public Sub PopulateSquad(ByVal pr_Team_s As String, _
                         ByVal pr_League_s As String)

Dim vl_Row_i As Integer
Dim vl_Player_o As Player
Dim vl_Player_s As String
Dim vl_Index_l As Long
Dim vl_Word_s As String
Dim vl_Tick_b As Boolean

On Error GoTo ErrorHandler

    'Retrieve the squad
    Set vp_Squad_o = vp_DBManager_o.GetSquad(pr_Team_s, pr_League_s)

    'Empty the flexgrid
    frmMain.fgrSquad.Rows = 1

    'Add the players
    vl_Row_i = 1
    For Each vl_Player_o In vp_Squad_o
        
        vl_Player_s = vl_Player_o.PlayerName & vbTab & _
                      vl_Player_o.PlayerAge & vbTab & _
                      vl_Player_o.PlayerLevel & vbTab & _
                      vl_Player_o.PlayerCat & vbTab & _
                      vl_Player_o.PlayerPos & vbTab & _
                      vl_Player_o.PlayerType & vbTab & _
                      vl_Player_o.PlayerSess & vbTab
        
        'Are there any notes?
        If (vl_Player_o.PlayerNotes <> vbNullString) Then
        
            'Grab the first word - find the first space
            vl_Index_l = InStr(vl_Player_o.PlayerNotes, " ")
            
            If (vl_Index_l = 0) Then 'no spaces
                vl_Word_s = vl_Player_o.PlayerNotes
            Else
                vl_Word_s = Left$(vl_Player_o.PlayerNotes, vl_Index_l - 1)
            End If
            
            'If the first word is 3 chars or less, insert it
            'Otherwise insert a tick
            If (Len(vl_Word_s) <= 3 And _
                (IsNumeric(vl_Word_s) Or _
                (LCase$(Left$(vl_Word_s, 1)) = "i" And IsNumeric(Mid$(vl_Word_s, 2))) Or _
                (LCase$(Left$(vl_Word_s, 1)) = "s" And IsNumeric(Mid$(vl_Word_s, 2))))) Then
                vl_Player_s = vl_Player_s & vl_Word_s
                vl_Tick_b = False
            Else
                vl_Player_s = vl_Player_s & "ü"
                vl_Tick_b = True
            End If
        
        Else
            vl_Tick_b = False
        End If
        
        frmMain.fgrSquad.AddItem vl_Player_s
        
        'Change the font to display a tick properly
        If vl_Tick_b Then
            frmMain.fgrSquad.Row = vl_Row_i
            frmMain.fgrSquad.Col = 7
            frmMain.fgrSquad.CellFontName = "Wingdings"
        End If

        vl_Row_i = vl_Row_i + 1
    
    Next

    'Select the headings row
    frmMain.fgrSquad.Row = 0
    frmMain.fgrSquad.Col = 0
    frmMain.fgrSquad.ColSel = 7

    Exit Sub

ErrorHandler:

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "mdlMain.PopulateSquad", Err.Description
    End If

End Sub

Public Sub ImportData(ByVal pr_MDBPath_s As String)

Dim vl_Players_o As Squad
Dim vl_Player_o As Player
Dim vl_Counter_i As Integer
Dim vl_PlayerKey_l As Long

On Error GoTo ErrorHandler

    'Get the players from the V1 database
    Set vl_Players_o = vp_DBManager_o.GetPlayers(pr_MDBPath_s)
    
    'Disable the main form
    frmMain.Enabled = False
    
    'Set frmImport's attributes and show it
    frmImport.lblImport.Caption = "Importing " & vl_Players_o.Count & " players"
    frmImport.prbImport.Min = 0
    frmImport.prbImport.Max = vl_Players_o.Count
    frmImport.prbImport.Value = 0
    frmImport.prbImport.Visible = True
    frmImport.MousePointer = vbHourglass
    frmImport.Show

    'Loop through the players, adding them to the V2 Database
    vl_Counter_i = 0
    For Each vl_Player_o In vl_Players_o

        vl_PlayerKey_l = vp_DBManager_o.AddPlayer(vl_Player_o)
        
        If (vl_PlayerKey_l > vp_MaxPlayerKey_l) Then
            vp_MaxPlayerKey_l = vl_PlayerKey_l
            vl_Counter_i = vl_Counter_i + 1
        End If
        
        frmImport.prbImport.Value = frmImport.prbImport.Value + 1
        DoEvents

    Next

    MsgBox vl_Counter_i & " players added or updated"

    'Re-enable the main form
    frmMain.Enabled = True
    frmImport.MousePointer = vbNormal

    Exit Sub

ErrorHandler:

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "mdlMain.ImportV1Data", Err.Description
    End If

End Sub
