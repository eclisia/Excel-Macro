VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar 
   Caption         =   "Sélection d'une date"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6540
   OleObjectBlob   =   "Calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'       This code works onlys with : DateCalendarPicker VBA module
'************************************************************************

Public dateSelected As String


Private Sub DTPickerCalendar_Change()

End Sub





Private Sub DTPickerCalendar_Click()
'    MsgBox DTPickerCalendar.Value
    TextBoxDate.Value = DTPickerCalendar.Value
    DTPickerCalendar.CustomFormat = "dd/MM/yyyy"
    dateSelected = DTPickerCalendar.Value
    Debug.Print "Date sélectionnée non formatée" & dateSelected
End Sub

Private Sub UserForm_Click()
    Range(mycellAdressGlobal).NumberFormat = "dd/MM/yyyy"
    
    
    If dateSelected <> "" Then
        'Permit to check if the string is empty. Whereas, the CDate conversion will raise on error
        Range(mycellAdressGlobal).Value = CDate(dateSelected)   'Cdate pour la convertion de type
        Debug.Print "Date à charger " & CDate(dateSelected)
    Else
        Debug.Print "DATE VIDE " & dateSelected
    End If
    
    Unload Me
End Sub

Private Sub UserForm_Initialize()



End Sub
