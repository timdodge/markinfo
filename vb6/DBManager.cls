VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private vm_Conn_o As ADODB.Connection
Private vm_Session_i As Integer

'Constants
Private Const cm_DEFAULT_PROVIDER_s As String = "Microsoft.Jet.OLEDB.4.0"
Private Const cm_MARKINFO_s As String = "MarkInfo21.mdb"

'name: Create
'description: Creates the connection and ensures that the table is available

Public Sub Create(Optional ByVal pr_UDL_s As String = vbNullString)

On Error GoTo ErrorHandler

    'Open the connection
    If (pr_UDL_s = vbNullString) Then
        vm_Conn_o.Provider = cm_DEFAULT_PROVIDER_s
        vm_Conn_o.Open App.Path & "/" & cm_MARKINFO_s
    Else
        vm_Conn_o.Open "File Name=" & pr_UDL_s
    End If
    
    Exit Sub

ErrorHandler:

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.Create", Err.Description
    End If

End Sub

'name: AddLeague
'description: Adds a League to the database, returns the LeagueKey

Public Function AddLeague(ByVal pr_League_s As String) As Long

Dim vl_League_s As String
Dim vl_SQL_s As String
Dim vl_RecordSet_o As ADODB.Recordset

On Error GoTo ErrorHandler

    'Double up on single quotes in LeagueName
    vl_League_s = Replace(pr_League_s, "'", "''")

    'Check whether League is already in database
    vl_SQL_s = "SELECT Leaguekey FROM Leagues WHERE LeagueName = '" & _
               vl_League_s & "';"
    Set vl_RecordSet_o = vm_Conn_o.Execute(vl_SQL_s)
    
    If Not vl_RecordSet_o.EOF Then

        'League already in database
        vl_RecordSet_o.MoveFirst
        AddLeague = vl_RecordSet_o("LeagueKey")

    Else

        vl_RecordSet_o.Close

        'Add the League
        vm_Conn_o.Execute "INSERT INTO Leagues (LeagueName) VALUES ('" & _
                          vl_League_s & "');"
        
        'Get the LeagueKey of the league just added
        Set vl_RecordSet_o = vm_Conn_o.Execute(vl_SQL_s)
        vl_RecordSet_o.MoveFirst
        AddLeague = vl_RecordSet_o("LeagueKey")

    End If

    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing

    Exit Function

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.AddLeague", Err.Description
    End If

End Function

'name: AddTeam
'description: Adds a team to the database, returns the TeamKey

Public Function AddTeam(ByVal pr_Team_s As String, _
                        ByVal pr_League_s As String) As Long

Dim vl_Team_s As String
Dim vl_League_s As String
Dim vl_SQL_s As String
Dim vl_RecordSet_o As ADODB.Recordset
Dim vl_LeagueKey_l As Long

On Error GoTo ErrorHandler

    'Double up on single quotes
    vl_Team_s = Replace(pr_Team_s, "'", "''")
    vl_League_s = Replace(pr_League_s, "'", "''")

    'Check to see whether the Team exists
    vl_SQL_s = "SELECT t.TeamKey FROM Teams t, Leagues l " & _
               "WHERE t.LeagueKey = l.LeagueKey " & _
               "AND t.TeamName = '" & vl_Team_s & "' " & _
               "AND l.LeagueName = '" & vl_League_s & "';"
    Set vl_RecordSet_o = vm_Conn_o.Execute(vl_SQL_s)
    
    If Not vl_RecordSet_o.EOF Then

        vl_RecordSet_o.MoveFirst
        AddTeam = vl_RecordSet_o("TeamKey")

    Else
    
        vl_RecordSet_o.Close
       
        'Get the LeagueKey (also adds the league if it doesn't exist)
        vl_LeagueKey_l = AddLeague(pr_League_s)
        
        'Add the Team
        vm_Conn_o.Execute "INSERT INTO Teams (LeagueKey, TeamName) VALUES (" & _
                          vl_LeagueKey_l & ", '" & vl_Team_s & "');"

        'Get the TeamKey of the added team
        Set vl_RecordSet_o = vm_Conn_o.Execute(vl_SQL_s)
        vl_RecordSet_o.MoveFirst
        AddTeam = vl_RecordSet_o("TeamKey")

    End If

    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing

    Exit Function

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.AddTeam", Err.Description
    End If

End Function

'name: AddPlayer
'description: Adds the player to the database, returns the PlayerKey

Public Function AddPlayer(ByVal pr_Player_o As Player) As Long

'Player details
Dim vl_Player_s As String
Dim vl_Team_s As String
Dim vl_League_s As String
Dim vl_Cat_s As String
Dim vl_Pos_s As String
Dim vl_Type_s As String
Dim vl_TeamKey_l As Long
Dim vl_PlayerKey_l As Long
Dim vl_PlayerSess_i As Integer
Dim vl_CurrSess_i As Integer
Dim vl_CurrCat_s As String
Dim vl_CurrPos_s As String
Dim vl_PlayerNotes_s As String

Dim vl_SQL_s As String
Dim vl_RecordSet_o As ADODB.Recordset

On Error GoTo ErrorHandler

    'Double up on single quotes
    vl_Player_s = Replace(pr_Player_o.PlayerName, "'", "''")
    vl_Team_s = Replace(pr_Player_o.PlayerTeam, "'", "''")
    vl_League_s = Replace(pr_Player_o.PlayerLeague, "'", "''")

    'Validate Pos, Cat and Type
    Set vl_RecordSet_o = vm_Conn_o.Execute("SELECT PlayerPos FROM Positions " & _
                         "WHERE PlayerPos = '" & pr_Player_o.PlayerPos & "';")

    If Not vl_RecordSet_o.EOF Then
        vl_RecordSet_o.MoveFirst
        vl_Pos_s = Trim$(vl_RecordSet_o("PlayerPos"))
    Else
        vl_Pos_s = "?"
    End If

    vl_RecordSet_o.Close
    
    If (vl_Pos_s = "GK") Then
        vl_Cat_s = "---"
    Else
        Set vl_RecordSet_o = vm_Conn_o.Execute("SELECT PlayerCat FROM Categories " & _
                             "WHERE PlayerCat = '" & pr_Player_o.PlayerCat & "';")

        If Not vl_RecordSet_o.EOF Then
            vl_RecordSet_o.MoveFirst
            vl_Cat_s = Trim$(vl_RecordSet_o("PlayerCat"))
        Else
            vl_Cat_s = "?"
        End If
        vl_RecordSet_o.Close
    End If

    Set vl_RecordSet_o = vm_Conn_o.Execute("SELECT PlayerType FROM Types " & _
                         "WHERE PlayerType = '" & pr_Player_o.PlayerType & "';")

    If Not vl_RecordSet_o.EOF Then
        vl_RecordSet_o.MoveFirst
        vl_Type_s = Trim$(vl_RecordSet_o("PlayerType"))
    Else
        vl_Type_s = "---"
    End If

    vl_RecordSet_o.Close
    
    'If player session number was optional (=0) make it this session
    If (pr_Player_o.PlayerSess = 0) Then
        vl_PlayerSess_i = vm_Session_i
    Else
        vl_PlayerSess_i = pr_Player_o.PlayerSess
    End If

    'If there were no notes, insert a NULL
    If (pr_Player_o.PlayerNotes = vbNullString) Then
        vl_PlayerNotes_s = "NULL"
    Else
        vl_PlayerNotes_s = "'" & Replace(pr_Player_o.PlayerNotes, "'", "''") & "'"
    End If
    
    'Check to see whether the player exists
    vl_SQL_s = "SELECT p.* FROM Players p, Teams t, Leagues l " & _
               "WHERE p.TeamKey = t.TeamKey " & _
               "AND t.LeagueKey = l.LeagueKey " & _
               "AND p.PlayerName = '" & vl_Player_s & "' " & _
               "AND t.TeamName = '" & vl_Team_s & "' " & _
               "AND l.LeagueName = '" & vl_League_s & "';"
    Set vl_RecordSet_o = New ADODB.Recordset
    vl_RecordSet_o.Open vl_SQL_s, vm_Conn_o, adOpenDynamic, adLockPessimistic, adCmdText
        
    If Not vl_RecordSet_o.EOF Then

        'Player exists in database
        vl_RecordSet_o.MoveFirst
        vl_PlayerKey_l = vl_RecordSet_o("PlayerKey")
        
        'Update the player
        If (vl_PlayerSess_i >= vl_RecordSet_o("PlayerSess")) Then
        
            'Use the current Category if the new one is not known
            If (vl_Cat_s = "?" And Trim$(vl_RecordSet_o("PlayerCat")) <> "?") Then
                vl_Cat_s = vl_RecordSet_o("PlayerCat")
            End If
            
            'Use the current Position if the new one is not known
            If (vl_Pos_s = "?" And Trim$(vl_RecordSet_o("PlayerPos")) <> "?") Then
                vl_Pos_s = vl_RecordSet_o("PlayerPos")
            End If
            
            vm_Conn_o.Execute "UPDATE Players " & _
                              "SET PlayerAge = " & pr_Player_o.PlayerAge & ", " & _
                                "PlayerLevel = " & pr_Player_o.PlayerLevel & ", " & _
                                "PlayerCat = '" & vl_Cat_s & "', " & _
                                "PlayerPos = '" & vl_Pos_s & "', " & _
                                "PlayerType = '" & vl_Type_s & "', " & _
                                "PlayerSess = " & vl_PlayerSess_i & ", " & _
                                "PlayerNotes = " & vl_PlayerNotes_s & " " & _
                              "WHERE PlayerKey = " & vl_PlayerKey_l & ";"

        End If
        vl_RecordSet_o.Close

    Else

        'New player
        vl_RecordSet_o.Close
        vl_TeamKey_l = AddTeam(pr_Player_o.PlayerTeam, pr_Player_o.PlayerLeague)

        'Add the player
        vm_Conn_o.Execute "INSERT INTO Players (TeamKey, PlayerName, " & _
                          "PlayerAge, PlayerLevel, PlayerCat, PlayerPos, " & _
                          "PlayerType, PlayerSess, PlayerNotes) VALUES (" & _
                          vl_TeamKey_l & ", '" & vl_Player_s & "', " & _
                          pr_Player_o.PlayerAge & ", " & pr_Player_o.PlayerLevel & ", '" & _
                          vl_Cat_s & "', '" & vl_Pos_s & "', '" & vl_Type_s & _
                          "', " & vl_PlayerSess_i & ", " & vl_PlayerNotes_s & ");"
        
        'Get the PlayerKey of the added player
        Set vl_RecordSet_o = vm_Conn_o.Execute(vl_SQL_s)
        vl_RecordSet_o.MoveFirst
        vl_PlayerKey_l = vl_RecordSet_o("PlayerKey")

        'Clear up
        vl_RecordSet_o.Close
        Set vl_RecordSet_o = Nothing

    End If

    AddPlayer = vl_PlayerKey_l

    Exit Function
    
ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.AddPlayer", Err.Description
    End If

End Function

'name: DelPlayer
'description: Deletes the player!

Public Sub DelPlayer(ByVal pr_Name_s As String, _
                     ByVal pr_Team_s As String, _
                     ByVal pr_League_s As String)

Dim vl_Name_s As String
Dim vl_Team_s As String
Dim vl_League_s As String
Dim vl_RecordSet_o As ADODB.Recordset
Dim vl_PlayerKey_l As Long

On Error GoTo ErrorHandler

    vl_Name_s = Replace(pr_Name_s, "'", "''")
    vl_Team_s = Replace(pr_Team_s, "'", "''")
    vl_League_s = Replace(pr_League_s, "'", "''")

    'Find the PlayerKey
    Set vl_RecordSet_o = vm_Conn_o.Execute("SELECT p.PlayerKey FROM " & _
                                           "Players p, Teams t, Leagues l " & _
                                           "WHERE p.TeamKey = t.TeamKey " & _
                                           "AND t.LeagueKey = l.LeagueKey " & _
                                           "AND p.PlayerName = '" & vl_Name_s & "' " & _
                                           "AND t.TeamName = '" & vl_Team_s & "' " & _
                                           "AND l.LeagueName = '" & vl_League_s & "';")
    If Not vl_RecordSet_o.EOF Then
    
        'Player exists - delete him
        vl_RecordSet_o.MoveFirst
        vl_PlayerKey_l = vl_RecordSet_o("PlayerKey")
    
        'Delete the player
        vm_Conn_o.Execute "DELETE FROM Players WHERE PlayerKey = " & vl_PlayerKey_l & ";"
    
    End If
    
    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing
    
    Exit Sub

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.DelPlayer", Err.Description
    End If

End Sub

'name: DelTeam
'description: Deletes the whole team

Public Sub DelTeam(ByVal pr_Team_s As String, _
                   ByVal pr_League_s As String)

Dim vl_TeamKey_l As Long
Dim vl_Team_s As String
Dim vl_League_s As String
Dim vl_RecordSet_o As ADODB.Recordset

On Error GoTo ErrorHandler

    vl_Team_s = Replace(pr_Team_s, "'", "''")
    vl_League_s = Replace(pr_League_s, "'", "''")

    'Find the TeamKey
    Set vl_RecordSet_o = vm_Conn_o.Execute("SELECT t.TeamKey FROM Teams t, Leagues l " & _
                                           "WHERE l.LeagueKey = t.LeagueKey " & _
                                           "AND t.TeamName = '" & vl_Team_s & "' " & _
                                           "AND l.LeagueName = '" & vl_League_s & "';")

    If Not vl_RecordSet_o.EOF Then
        
        'Team found - delete it
        vl_RecordSet_o.MoveFirst
        vl_TeamKey_l = vl_RecordSet_o("TeamKey")
    
        'Delete the players
        vm_Conn_o.Execute "DELETE FROM Players WHERE TeamKey = " & vl_TeamKey_l & ";"
    
        'Delete the team
        vm_Conn_o.Execute "DELETE FROM Teams WHERE TeamKey = " & vl_TeamKey_l & ";"
    
    End If

    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing

    Exit Sub

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.DelTeam", Err.Description
    End If

End Sub

'name: DelLeague
'description: Deletes the whole league (Should cascade down to Teams and Players)

Public Sub DelLeague(ByVal pr_League_s As String)

Dim vl_League_s As String
Dim vl_RecordSet_o As ADODB.Recordset

On Error GoTo ErrorHandler
    
    vl_League_s = Replace(pr_League_s, "'", "''")
    
    'Get a recordset containing the TeamKeys to be deleted
    Set vl_RecordSet_o = vm_Conn_o.Execute("SELECT t.TeamKey FROM " & _
                                           "Teams t, Leagues l " & _
                                           "WHERE t.LeagueKey = l.LeagueKey " & _
                                           "AND l.LeagueName = '" & vl_League_s & "';")
    If Not vl_RecordSet_o.EOF Then
    
        vl_RecordSet_o.MoveFirst
        Do While Not vl_RecordSet_o.EOF
            
            'Delete the Players in this Team
            vm_Conn_o.Execute "DELETE FROM Players " & _
                              "WHERE TeamKey = " & vl_RecordSet_o("TeamKey") & ";"

            'Now delete the team
            vm_Conn_o.Execute "DELETE FROM Teams " & _
                              "WHERE TeamKey = " & vl_RecordSet_o("TeamKey") & ";"

            vl_RecordSet_o.MoveNext
            
        Loop
        
    End If

    'Delete the league
    vm_Conn_o.Execute "DELETE FROM Leagues WHERE LeagueName = '" & vl_League_s & "';"
    
    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing
    
    Exit Sub

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.DelLeague", Err.Description
    End If

End Sub

'name: MovePlayer
'description: Deletes the player from his old team, adds him to the new

Public Function MovePlayer(ByVal pr_Player_o As Player, _
                           ByVal pr_OldTeam_s As String) As Long

    'Remove old player
    DelPlayer pr_Player_o.PlayerName, pr_OldTeam_s, pr_Player_o.PlayerLeague
    
    'Add the new player
    MovePlayer = AddPlayer(pr_Player_o)

End Function

'name: NewSeason
'description: Updates the session numbers for the new season

Public Sub NewSeason(ByVal pr_League_s As String, _
                     Optional ByVal pr_SessionLimit_i As Integer = 2)

Dim vl_SQL_s As String
Dim vl_RecordSet_o As ADODB.Recordset
Dim vl_PlayerKey_l As Long
Dim vl_Session_i As Integer

On Error GoTo ErrorHandler

    'Grab a recordset with all the PlayerKeys and Session Numbers of the players
    'in the selected League
    vl_SQL_s = "SELECT p.PlayerKey, p.PlayerSess FROM " & _
               "Players p, Teams t, Leagues l " & _
               "WHERE p.TeamKey = t.TeamKey " & _
               "AND t.LeagueKey = l.LeagueKey " & _
               "AND l.LeagueName = '" & pr_League_s & "';"
    Set vl_RecordSet_o = vm_Conn_o.Execute(vl_SQL_s)

    'Loop Through the recordset, updating the Players as necessary
    If Not vl_RecordSet_o.EOF Then
    
        vl_RecordSet_o.MoveFirst
        Do While Not vl_RecordSet_o.EOF
        
            'Get PlayerKey of current player
            If Not IsNull(vl_RecordSet_o("PlayerKey")) Then
                vl_PlayerKey_l = vl_RecordSet_o("PlayerKey")
            Else
                Err.Raise vbObjectError + 1001, "DBManager.NewSeason", _
                          "PlayerKey of current player is NULL"
            End If
        
            'Get Session number of current player
            If Not IsNull(vl_RecordSet_o("PlayerSess")) Then
                vl_Session_i = vl_RecordSet_o("PlayerSess")
            Else
                Err.Raise vbObjectError + 1001, "DBManager.NewSeason", _
                          "Session Number of player number " & _
                          vl_RecordSet_o("PlayerKey") & " is NULL"
            End If
            
            If (vl_Session_i < (-1) * (pr_SessionLimit_i - 1)) Then
            
                'remove all players older than the SessionLimit
                vl_SQL_s = "DELETE FROM Players  " & _
                           "WHERE PlayerKey = " & vl_PlayerKey_l & ";"
            
            ElseIf (vl_Session_i < 0) Then
            
                'Decrement the session number of players already > 1 session old
                vl_SQL_s = "UPDATE Players " & _
                           "SET PlayerSess = PlayerSess - 1 " & _
                           "WHERE PlayerKey = " & vl_PlayerKey_l & ";"
            
            Else
            
                'Set the session number of all current session players to -1
                vl_SQL_s = "UPDATE Players " & _
                           "SET PlayerSess = - 1 " & _
                           "WHERE PlayerKey = " & vl_PlayerKey_l & ";"
            
            End If
        
            'Execute the DELETE / UPDATE and move to next player
            vm_Conn_o.Execute vl_SQL_s
            vl_RecordSet_o.MoveNext
            
        Loop
        
    End If

    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing

    Exit Sub

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.NewSeason", Err.Description
    End If

End Sub

'name: GetSquad
'description: Returns a squad object containing the team

Public Function GetSquad(ByVal pr_Team_s As String, _
                         ByVal pr_League_s As String) As Squad

Dim vl_Team_s As String
Dim vl_League_s As String
Dim vl_SQL_s As String
Dim vl_RecordSet_o As Recordset
Dim vl_Squad_o As Squad
    
On Error GoTo ErrorHandler

    vl_Team_s = Replace(pr_Team_s, "'", "''")
    vl_League_s = Replace(pr_League_s, "'", "''")

    vl_SQL_s = "SELECT p.PlayerName, p.PlayerAge, p.PlayerLevel, p.PlayerCat, " & _
               "p.PlayerPos, p.PlayerType, p.PlayerSess, p.PlayerNotes " & _
               "FROM Players p, Teams t, Leagues l " & _
               "WHERE p.TeamKey = t.TeamKey " & _
               "AND t.LeagueKey = l.LeagueKey " & _
               "AND t.TeamName = '" & vl_Team_s & "' " & _
               "AND l.LeagueName = '" & vl_League_s & "' " & _
               "ORDER BY p.PlayerName;"

    Set vl_RecordSet_o = New ADODB.Recordset
    vl_RecordSet_o.Open vl_SQL_s, vm_Conn_o, adOpenStatic, adLockReadOnly, adCmdText

    Set vl_Squad_o = New Squad
    
    If Not (vl_RecordSet_o.EOF) Then
    
        If Not (vl_RecordSet_o.RecordCount = 1 And _
                vl_RecordSet_o("PlayerName") = vbNullString) Then
        
            vl_RecordSet_o.MoveFirst
            Do While Not vl_RecordSet_o.EOF
                
                If IsNull(vl_RecordSet_o("PlayerNotes")) Then
                    
                    vl_Squad_o.Add pr_League_s, pr_Team_s, vl_RecordSet_o("PlayerName"), _
                                   vl_RecordSet_o("PlayerAge"), vl_RecordSet_o("PlayerLevel"), _
                                   vl_RecordSet_o("PlayerCat"), vl_RecordSet_o("PlayerPos"), _
                                   vl_RecordSet_o("PlayerType"), vl_RecordSet_o("PlayerSess")
                
                Else
                    
                    vl_Squad_o.Add pr_League_s, pr_Team_s, vl_RecordSet_o("PlayerName"), _
                                   vl_RecordSet_o("PlayerAge"), vl_RecordSet_o("PlayerLevel"), _
                                   vl_RecordSet_o("PlayerCat"), vl_RecordSet_o("PlayerPos"), _
                                   vl_RecordSet_o("PlayerType"), vl_RecordSet_o("PlayerSess"), _
                                   vl_RecordSet_o("PlayerNotes")
                
                End If
    
                vl_RecordSet_o.MoveNext
            
            Loop
    
        End If
    
    End If
    
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing

    Set GetSquad = vl_Squad_o

    Exit Function

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.GetSquad", Err.Description
    End If

End Function

'name: GetTeams
'description: Returns a string array containing all the teams in a league

Public Function GetTeams(pr_League_s As String) As String()

Dim vl_RecordSet_o As ADODB.Recordset
Dim vl_SQL_s As String
Dim vl_League_s As String
Dim vl_Teams_s As String

On Error GoTo ErrorHandler

    vl_League_s = Replace(pr_League_s, "'", "''")

    vl_SQL_s = "SELECT t.TeamName " & _
               "FROM Teams t, Leagues l " & _
               "WHERE t.LeagueKey = l.LeagueKey " & _
               "AND l.LeagueName = '" & vl_League_s & "' " & _
               "ORDER BY t.TeamName;"
    
    Set vl_RecordSet_o = New ADODB.Recordset
    vl_RecordSet_o.Open vl_SQL_s, vm_Conn_o, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not (vl_RecordSet_o.EOF) Then
    
        vl_RecordSet_o.MoveFirst
        Do While Not vl_RecordSet_o.EOF
        
            vl_Teams_s = vl_Teams_s & vl_RecordSet_o("TeamName") & "||"
            vl_RecordSet_o.MoveNext
            
        Loop
    
        'Trim string and remove final "||"
        vl_Teams_s = Trim$(vl_Teams_s)
        vl_Teams_s = Left$(vl_Teams_s, Len(vl_Teams_s) - 2)
    
        'if there are no teams, return "NO TEAMS"
        If (Len(vl_Teams_s) = 0) Then
            vl_Teams_s = "NO TEAMS"
        End If
    
    Else
        
        vl_Teams_s = "NO TEAMS"
    
    End If

    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing

    GetTeams = Split(vl_Teams_s, "||")

    Exit Function

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.GetTeams", Err.Description
    End If

End Function

'name: GetLeagues
'description: Returns a string array containing all the leagues

Public Function GetLeagues() As String()

Dim vl_RecordSet_o As ADODB.Recordset
Dim vl_Leagues_s As String

On Error GoTo ErrorHandler

    Set vl_RecordSet_o = New ADODB.Recordset
    vl_RecordSet_o.Open "SELECT LeagueName FROM Leagues ORDER BY LeagueName;", _
                        vm_Conn_o, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not (vl_RecordSet_o.EOF) Then
        
        vl_RecordSet_o.MoveFirst
        Do While Not vl_RecordSet_o.EOF
            
            vl_Leagues_s = vl_Leagues_s & vl_RecordSet_o("LeagueName") & "||"
            vl_RecordSet_o.MoveNext
    
        Loop
    
        'Trim string and remove final "||"
        vl_Leagues_s = Trim$(vl_Leagues_s)
        vl_Leagues_s = Left$(vl_Leagues_s, Len(vl_Leagues_s) - 2)
    
        'if there are no leagues, return "NO LEAGUES"
        If (Len(vl_Leagues_s) = 0) Then
            vl_Leagues_s = "NO LEAGUES"
        End If
    
    Else
        
        vl_Leagues_s = "NO LEAGUES"
    
    End If

    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing

    GetLeagues = Split(vl_Leagues_s, "||")

    Exit Function

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.GetLeagues", Err.Description
    End If

End Function

'name: RenameTeam
'description: Renames the team!

Public Sub RenameTeam(ByVal pr_OldName_s As String, _
                      ByVal pr_NewName_s As String, _
                      ByVal pr_League_s As String)

Dim vl_SQL_s As String
Dim vl_RecordSet_o As Recordset
Dim vl_LeagueKey_l As Long

On Error GoTo ErrorHandler

    'Find League Key
    vl_SQL_s = "SELECT LeagueKey FROM Leagues WHERE LeagueName = '" & pr_League_s & "';"
    
    Set vl_RecordSet_o = vm_Conn_o.Execute(vl_SQL_s)
    
    If Not vl_RecordSet_o.EOF Then
        vl_RecordSet_o.MoveFirst
        vl_LeagueKey_l = vl_RecordSet_o("LeagueKey")
    Else
        Err.Raise vbObjectError + 1001, "DBManager.RenameTeam", "Team Not Found!"
    End If

    vl_SQL_s = "UPDATE Teams " & _
               "SET TeamName = '" & pr_NewName_s & "' " & _
               "WHERE TeamName = '" & pr_OldName_s & "' " & _
               "AND LeagueKey = " & vl_LeagueKey_l & ";"
    
    vm_Conn_o.Execute vl_SQL_s

    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing

    Exit Sub

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.RenameTeam", Err.Description
    End If

End Sub

'name: RenameLeague
'description: Renames the League!

Public Sub RenameLeague(ByVal pr_OldName_s As String, _
                        ByVal pr_NewName_s As String)

Dim vl_SQL_s As String

On Error GoTo ErrorHandler

    vl_SQL_s = "UPDATE Leagues " & _
               "SET LeagueName = '" & pr_NewName_s & "' " & _
               "WHERE LeagueName = '" & pr_OldName_s & "';"

    vm_Conn_o.Execute vl_SQL_s
    
    Exit Sub

ErrorHandler:

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.RenameLeague", Err.Description
    End If

End Sub

'name: GetPlayers
'description: Returns a squad full of the players from the v1 database

Public Function GetPlayers(ByVal pr_MDBPath_s As String) As Squad

Dim vl_Conn_o As ADODB.Connection
Dim vl_RecordSet_o As ADODB.Recordset
Dim vl_Field_o As ADODB.Field
Dim vl_Squad_o As Squad
Dim vl_SQL_s As String
Dim vl_PlayerNotes_s As String
Dim vl_v1mdb_b As Boolean

On Error GoTo ErrorHandler

    'Open connection
    Set vl_Conn_o = New ADODB.Connection
    vl_Conn_o.Provider = cm_DEFAULT_PROVIDER_s
    vl_Conn_o.Open pr_MDBPath_s
    
    'Open recordset
    Set vl_RecordSet_o = New ADODB.Recordset
    vl_RecordSet_o.Open "Players", vl_Conn_o, adOpenForwardOnly, adLockReadOnly, adCmdTable

    'Work out what version database it is
    vl_v1mdb_b = False
    
    'Look for v1 fields
    For Each vl_Field_o In vl_RecordSet_o.Fields
        If vl_Field_o.Name = "Position" Then
            vl_v1mdb_b = True
            Exit For
        End If
    Next

    Set vl_Squad_o = New Squad
    
    If vl_v1mdb_b Then
        
        'v1.0x database
        If Not vl_RecordSet_o.EOF Then
            vl_RecordSet_o.MoveFirst
            Do While Not vl_RecordSet_o.EOF
                
                vl_Squad_o.Add vl_RecordSet_o("League"), vl_RecordSet_o("Team"), _
                               vl_RecordSet_o("Name"), vl_RecordSet_o("Age"), _
                               vl_RecordSet_o("Level"), vl_RecordSet_o("Cat"), _
                               vl_RecordSet_o("Position"), vl_RecordSet_o("Type"), _
                               vl_RecordSet_o("Session")
                
                vl_RecordSet_o.MoveNext
            
            Loop
        End If

    Else

        'Close the previous recordset
        vl_RecordSet_o.Close

        vl_SQL_s = "SELECT l.LeagueName, t.TeamName, p.* " & _
                   "FROM Players p, Teams t, Leagues l " & _
                   "WHERE p.TeamKey = t.TeamKey " & _
                   "AND t.LeagueKey = l.LeagueKey;"
        
        vl_RecordSet_o.Open vl_SQL_s, vl_Conn_o, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        If Not vl_RecordSet_o.EOF Then
            vl_RecordSet_o.MoveFirst
            Do While Not vl_RecordSet_o.EOF
                
                If IsNull(vl_RecordSet_o("PlayerNotes")) Then
                    vl_PlayerNotes_s = vbNullString
                Else
                    vl_PlayerNotes_s = vl_RecordSet_o("PlayerNotes")
                End If
                
                vl_Squad_o.Add vl_RecordSet_o("LeagueName"), vl_RecordSet_o("TeamName"), _
                               vl_RecordSet_o("PlayerName"), vl_RecordSet_o("PlayerAge"), _
                               vl_RecordSet_o("PlayerLevel"), vl_RecordSet_o("PlayerCat"), _
                               vl_RecordSet_o("PlayerPos"), vl_RecordSet_o("PlayerType"), _
                               vl_RecordSet_o(10), vl_PlayerNotes_s
                
                vl_RecordSet_o.MoveNext
            
            Loop
        End If
    
    End If
    
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing
    vl_Conn_o.Close
    Set vl_Conn_o = Nothing
    
    Set GetPlayers = vl_Squad_o
    Set vl_Squad_o = Nothing

    Exit Function
    
ErrorHandler:
    
    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If Not (vl_Conn_o Is Nothing) Then
        If (vl_Conn_o.State = adStateOpen) Then
            vl_Conn_o.Close
        End If
        Set vl_Conn_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.GetPlayers", Err.Description
    End If

End Function

'name: Session (Get)
'description: returns the current session number

Public Property Get Session() As Integer

    If (vm_Session_i = 0) Then
        Session = GetSession
    Else
        Session = vm_Session_i
    End If

End Property

'name: Session (Let)
'description: Sets the current session number

Public Property Let Session(ByVal pr_Session_i As Integer)

    If (vm_Session_i <> pr_Session_i) Then
        vm_Session_i = pr_Session_i
        LetSession (pr_Session_i)
    End If

End Property

'name: GetSession
'description: Retreives the current session number from the database

Private Function GetSession() As Integer

Dim vl_SQL_s As String
Dim vl_RecordSet_o As ADODB.Recordset
    
On Error GoTo ErrorHandler
    
    vl_SQL_s = "SELECT CurrSession FROM Admin;"
    Set vl_RecordSet_o = vm_Conn_o.Execute(vl_SQL_s)
    
    If vl_RecordSet_o.EOF Then
        GetSession = 1
    Else
        vl_RecordSet_o.MoveFirst
        GetSession = vl_RecordSet_o("CurrSession")
    End If
    
    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing

    Exit Function

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.GetSession", Err.Description
    End If

End Function

'name: LetSession
'description: Sets the current session number to the database

Private Sub LetSession(pr_Session_i As Integer)

Dim vl_SQL_s As String
Dim vl_RecordSet_o As ADODB.Recordset
    
On Error GoTo ErrorHandler

    'Check to see if there is any session info in the Admin table
    vl_SQL_s = "SELECT CurrSession FROM Admin"
    Set vl_RecordSet_o = vm_Conn_o.Execute(vl_SQL_s)
    
    If vl_RecordSet_o.EOF Then
        'Nope - INSERT new value
        vl_SQL_s = "INSERT INTO Admin (CurrSession) VALUES (" & pr_Session_i & ");"
    Else
        'Yep - UPDATE with new value
        vl_SQL_s = "UPDATE Admin SET CurrSession = " & pr_Session_i & ";"
    End If
    
    'Clear up
    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing
    
    'Update table
    vm_Conn_o.Execute vl_SQL_s
    
    Exit Sub

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.LetSession", Err.Description
    End If

End Sub

'name: TableExists
'description: Returns true if it exists

Public Function TableExists(ByVal pr_Table_s As String) As Boolean

Dim vl_SQL_s As String
On Error GoTo ErrorHandler

    vl_SQL_s = "SELECT COUNT(*) FROM " & pr_Table_s & ";"
    vm_Conn_o.Execute vl_SQL_s
    
    TableExists = True
    
    Exit Function

ErrorHandler:

    TableExists = False

End Function

'name: GetMaxPlayerKey
'description: Returns the current maximum PlayerKey

Public Function GetMaxPlayerKey() As Long

Dim vl_RecordSet_o As ADODB.Recordset

On Error GoTo ErrorHandler

    Set vl_RecordSet_o = vm_Conn_o.Execute("SELECT MAX(PlayerKey) AS MaxPlayerKey FROM Players;")

    If Not vl_RecordSet_o.EOF Then
        vl_RecordSet_o.MoveFirst
        If IsNull(vl_RecordSet_o("MaxPlayerKey")) Then
            GetMaxPlayerKey = 0
        Else
            GetMaxPlayerKey = vl_RecordSet_o("MaxPlayerKey")
        End If
    Else
        GetMaxPlayerKey = 0
    End If

    vl_RecordSet_o.Close
    Set vl_RecordSet_o = Nothing

    Exit Function

ErrorHandler:

    'Clear up
    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If
    
    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.GetMaxPlayerKey", Err.Description
    End If

End Function

'name: CreateV2Tables
'description: Attempts to create the V2 tables

Public Sub CreateV2Tables()

Dim vl_SQL_s As String

On Error GoTo ErrorHandler

    'Create the tables
    vm_Conn_o.Execute "CREATE TABLE Admin (CurrSession TINYINT NOT NULL PRIMARY KEY);"
    vm_Conn_o.Execute "CREATE TABLE Categories (PlayerCat VARCHAR(3) NOT NULL " & _
                      "PRIMARY KEY);"
    vm_Conn_o.Execute "CREATE TABLE Positions (PlayerPos VARCHAR(6) NOT NULL " & _
                      "PRIMARY KEY);"
    vm_Conn_o.Execute "CREATE TABLE Types (PlayerType VARCHAR(4) NOT NULL " & _
                      "PRIMARY KEY);"
    vm_Conn_o.Execute "CREATE TABLE Leagues (LeagueKey INT NOT NULL AUTO_INCREMENT " & _
                      "PRIMARY KEY, LeagueName VARCHAR(20) NOT NULL, UNIQUE " & _
                      "INDEX IX_Leagues_LeagueName (LeagueName));"
    vm_Conn_o.Execute "CREATE TABLE Teams (TeamKey INT NOT NULL AUTO_INCREMENT " & _
                      "PRIMARY KEY, LeagueKey INT NOT NULL, TeamName VARCHAR(20) " & _
                      "NOT NULL, UNIQUE INDEX IX_Teams_TeamName (LeagueKey, " & _
                      "TeamName));"
    vm_Conn_o.Execute "CREATE TABLE Players (PlayerKey INT NOT NULL AUTO_INCREMENT " & _
                      "PRIMARY KEY, TeamKey INT NOT NULL, " & _
                      "PlayerName VARCHAR(20) NOT NULL, PlayerAge TINYINT NOT NULL, " & _
                      "PlayerLevel TINYINT NOT NULL, PlayerCat VARCHAR(3) NOT NULL, " & _
                      "PlayerPos VARCHAR(6) NOT NULL, PlayerType VARCHAR(4) NOT " & _
                      "NULL, PlayerSess SMALLINT NOT NULL, PlayerNotes VARCHAR(255) " & _
                      "NULL, UNIQUE INDEX IX_Players_PlayerName (TeamKey, PlayerName));"

    'Populate the dimensions
    vm_Conn_o.Execute "INSERT INTO Categories (PlayerCat) VALUES ('?');"
    vm_Conn_o.Execute "INSERT INTO Categories (PlayerCat) VALUES ('---');"
    vm_Conn_o.Execute "INSERT INTO Categories (PlayerCat) VALUES ('P');"
    vm_Conn_o.Execute "INSERT INTO Categories (PlayerCat) VALUES ('S');"
    vm_Conn_o.Execute "INSERT INTO Categories (PlayerCat) VALUES ('P/S');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('?');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('GK');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('SW');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('DF');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('DF/A');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('MF');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('MF/D');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('MF/G');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('MF/A');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('MF/A/D');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('FW');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('FWT');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('FWS');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('WG');"
    vm_Conn_o.Execute "INSERT INTO Positions (PlayerPos) VALUES ('UT');"
    vm_Conn_o.Execute "INSERT INTO Types (PlayerType) VALUES ('---');"
    vm_Conn_o.Execute "INSERT INTO Types (PlayerType) VALUES ('APP');"
    vm_Conn_o.Execute "INSERT INTO Types (PlayerType) VALUES ('FUT');"
    vm_Conn_o.Execute "INSERT INTO Types (PlayerType) VALUES ('SBY');"
    vm_Conn_o.Execute "INSERT INTO Types (PlayerType) VALUES ('STAR');"

    Exit Sub

ErrorHandler:

    'Try and drop the tables if they exist
    If TableExists("Players") Then
        vm_Conn_o.Execute "DROP TABLE Players;"
    End If

    If TableExists("Teams") Then
        vm_Conn_o.Execute "DROP TABLE Teams;"
    End If

    If TableExists("Leagues") Then
        vm_Conn_o.Execute "DROP TABLE Leagues;"
    End If

    If TableExists("Categories") Then
        vm_Conn_o.Execute "DROP TABLE Categories;"
    End If

    If TableExists("Types") Then
        vm_Conn_o.Execute "DROP TABLE Types;"
    End If

    If TableExists("Positions") Then
        vm_Conn_o.Execute "DROP TABLE Positions;"
    End If

    If TableExists("Admin") Then
        vm_Conn_o.Execute "DROP TABLE Admin;"
    End If

    MsgBox "Unable to create the tables - please create them manually"
    Err.Raise vbObjectError + 1001, "DBManager.CreateV2Tables", "Unable to create MarkInfo V2 tables"

End Sub

'name: GetCats
'description: Returns a string array containing all the Player Categories

Public Function GetCats() As String()

Dim vl_RecordSet_o As ADODB.Recordset
Dim vl_Cats_s As String

On Error GoTo ErrorHandler

    Set vl_RecordSet_o = New ADODB.Recordset
    vl_RecordSet_o.Open "Categories", vm_Conn_o, adOpenStatic, adLockReadOnly, adCmdTable

    If Not vl_RecordSet_o.EOF Then
        
        vl_RecordSet_o.MoveFirst
        Do While Not vl_RecordSet_o.EOF
        
            vl_Cats_s = vl_Cats_s & vl_RecordSet_o("PlayerCat") & " "
            vl_RecordSet_o.MoveNext
            
        Loop
        
    End If
    
    GetCats = Split(Trim$(vl_Cats_s))
    
    Exit Function

ErrorHandler:

    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.GetCats", Err.Description
    End If

End Function

'name: GetPosns
'description: Returns a string array containing all the Player Categories

Public Function GetPosns() As String()
Dim vl_RecordSet_o As ADODB.Recordset
Dim vl_Posns_s As String

On Error GoTo ErrorHandler

    Set vl_RecordSet_o = New ADODB.Recordset
    vl_RecordSet_o.Open "Positions", vm_Conn_o, adOpenStatic, adLockReadOnly, adCmdTable

    If Not vl_RecordSet_o.EOF Then
        
        vl_RecordSet_o.MoveFirst
        Do While Not vl_RecordSet_o.EOF
        
            vl_Posns_s = vl_Posns_s & vl_RecordSet_o("PlayerPos") & " "
            vl_RecordSet_o.MoveNext
            
        Loop
        
    End If
    
    GetPosns = Split(Trim$(vl_Posns_s))
    
    Exit Function

ErrorHandler:

    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.GetPosns", Err.Description
    End If

End Function

'name: GetTypes
'description: Returns a string array containing all the Player Types

Public Function GetTypes() As String()
Dim vl_RecordSet_o As ADODB.Recordset
Dim vl_Types_s As String

On Error GoTo ErrorHandler

    Set vl_RecordSet_o = New ADODB.Recordset
    vl_RecordSet_o.Open "Types", vm_Conn_o, adOpenStatic, adLockReadOnly, adCmdTable

    If Not vl_RecordSet_o.EOF Then
        
        vl_RecordSet_o.MoveFirst
        Do While Not vl_RecordSet_o.EOF
        
            vl_Types_s = vl_Types_s & vl_RecordSet_o("PlayerType") & " "
            vl_RecordSet_o.MoveNext
            
        Loop
        
    End If
    
    GetTypes = Split(Trim$(vl_Types_s))
    
    Exit Function

ErrorHandler:

    If Not (vl_RecordSet_o Is Nothing) Then
        If (vl_RecordSet_o.State = adStateOpen) Then
            vl_RecordSet_o.Close
        End If
        Set vl_RecordSet_o = Nothing
    End If

    If (Err.Number > vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        Err.Raise vbObjectError + 1001, "DBManager.GetCats", Err.Description
    End If

End Function

'name: Class_Initialize
'description: Initialises the connection

Private Sub Class_Initialize()
    
    Set vm_Conn_o = New ADODB.Connection
    vm_Session_i = 0

End Sub

'name: Class_Terminate
'description: Closes and destroys the connection if it's still open

Private Sub Class_Terminate()

    If Not (vm_Conn_o Is Nothing) Then
        If (vm_Conn_o.State = adStateOpen) Then
            vm_Conn_o.Close
        End If
        Set vm_Conn_o = Nothing
    End If

End Sub
