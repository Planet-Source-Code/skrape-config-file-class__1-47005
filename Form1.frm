VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim CF As New ConfigFile
    
    
    CF.Create (App.Path + "\test.txt")
    Call CF.Add("First Name", "Bob")
    Call CF.Add("Maximize on Load", "True")
    Call CF.Add("Start Server on Load", "False")
    
    CF.Delete ("First Name")
    Call CF.Update("Start Server on Load", "True")
    
    MsgBox CF.Find("Maximize on Load")
    
    End
    
End Sub
