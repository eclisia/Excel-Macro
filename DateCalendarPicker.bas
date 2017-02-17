Attribute VB_Name = "DateCalendarPicker"
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
'       This code works onlys with : Calendar VBA UserForm module
'************************************************************************

Public mycellAdressGlobal As String

Sub GestionCalendrier(myDate As Date, mycellAdress As String)


    mycellAdressGlobal = mycellAdress
    
    'Gestion de la date par défaut
    If myDate = 0 Then
        'aucune date présente dans la cellule sélectionnée
        Calendar.DTPickerCalendar.Value = Now
    Else
        Calendar.DTPickerCalendar.Value = myDate
    End If
    Calendar.Show
    
    

    

End Sub
