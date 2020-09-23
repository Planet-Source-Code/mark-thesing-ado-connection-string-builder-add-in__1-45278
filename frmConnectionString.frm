VERSION 5.00
Begin VB.Form frmConnectionString 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Add In"
   ClientHeight    =   960
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   1995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleMode       =   0  'User
   ScaleWidth      =   2011.625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmConnectionString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VBInstance As VBIDE.VBE
Public Connect As ConnectionString
