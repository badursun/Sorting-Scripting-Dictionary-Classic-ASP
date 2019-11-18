<%
Class sortk
	'---------------------------------------------
	' Scripting Dictionary Sort Class
	' For Classic ASP
	' by Anthony Burak DURSUN 2019 (c) 
	' badursun@gmail.com
	' https://github.com/badursun/
	' Original Raw Code "Sorting a Dictionary Object" By Aaron A.
	' http://www.4guysfromrolla.com/webtech/062701-1.shtml
	'---------------------------------------------

	' Class Init
	'-----------------------------------
	Private Sub Class_Initialize()
	End Sub

	' Class Terminate
	'-----------------------------------
	Private Sub Class_Terminate()
	End Sub

	' Grab Forms Data And Add Dictionary
	'-----------------------------------
	Public Property Get GrabForms()
		Set d = Server.CreateObject("Scripting.Dictionary")
			For Each Item in Request.Form
				d.Add Item, Request.Form(Item)
			Next 
		Set GrabForms = d 
		Set d = Nothing
	End Property

	' Build Array
	'-----------------------------------
	Private Sub BuildArray(objDict, aTempArray)
		Dim nCount, strKey
		nCount = 0

		Redim aTempArray(objDict.Count - 1)

		For Each strKey In objDict.Keys
			aTempArray(nCount) = strKey 
			nCount = nCount + 1
		Next 
	End Sub

	' Sort Array by KeyName
	'-----------------------------------
	Private Sub SortArray(aTempArray) 
		Dim iTemp, jTemp, strTemp

		For iTemp = 0 To UBound(aTempArray)  
			For jTemp = 0 To iTemp  

				If strComp(aTempArray(jTemp),aTempArray(iTemp)) > 0 Then
					strTemp = aTempArray(jTemp) 
					aTempArray(jTemp) = aTempArray(iTemp) 
					aTempArray(iTemp) = strTemp 
				End If 

			Next 
		Next 
	End Sub

	' 
	'-----------------------------------
	Private Sub PrintDictionary(objDict, aTempArray) 
		Dim iTemp 
		For iTemp = 0 To UBound(aTempArray) 
			Response.Write(aTempArray(iTemp) & " - " & objDict.Item(aTempArray(iTemp)) & "<br>") 
		Next 
	End Sub

	'
	'-----------------------------------
	Public Sub PrintSortedDictionary(objDict)
		Dim aTemp
		Call BuildArray(objDict, aTemp)
		Call SortArray(aTemp)
		Call PrintDictionary(objDict, aTemp)
	End Sub

	'
	'-----------------------------------
	Public Sub PrintUnSortedDictionary(objDict)
		Dim aTemp
		Call BuildArray(objDict, aTemp)
		'Call SortArray(aTemp)
		Call PrintDictionary(objDict, aTemp)
	End Sub
End Class
%>
