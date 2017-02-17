VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'************************************************************************
'   The code of this page permits to :
'       1 - create a userform thanks to DTPickerCalendar
'       2 - update the date number of a given Range
'   Before using this code, please activate the DTPickerCalendar module with the Tools/Reference menu
'
'   Version : 01
'   Description : First release of the code
'   Author : Florent Tainturier - florent.tainturier@gmail.com
'
'   Instruction :
'       In a basis Worksheet, please pay attention to add dummy code (see other files). This dummy code permits to call this module.
'       This code works onlys with : DateCalendarPicker VBA module and with the Calendar VBA Userform module
'************************************************************************

Dim cellAddress As String

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)






    Dim myPlage As Range
    Dim myDateCell As Date


    Set myPlage = Range("I4:K999")

    'Si colonne J

    'Selectionner une seule cellule, sinon le code s'arrête
    If Target.Cells.Count > 1 Then
        Debug.Print "Sélectionnez une seule cellule."
        Exit Sub
    End If


    If Application.Intersect(Target, myPlage) Is Nothing Then

    Else
        'Si cellule non vide - alors on récupère la date
        'puis on affiche le calendrier
        myDateCell = Target.Cells.Value
        cellAddress = Target.Cells.Address
'        MsgBox "Vous avez double cliqué sur la cellule " & Target.Address
        Cancel = True
        GestionCalendrier myDateCell, cellAddress

    End If

End Sub


