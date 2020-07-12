VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "NOMBRE DEL PROCESO"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6810
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Worksheets("ESTIMADO").Range("B1") = NombreProceso

NombreProceso = Clear

UserForm1.Hide

End Sub
