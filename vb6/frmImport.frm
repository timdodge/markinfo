VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Database"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar prbImport 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2640
      Picture         =   "frmImport.frx":0000
      Top             =   225
      Width           =   480
   End
   Begin VB.Label lblImport 
      Caption         =   "Importing Records"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
