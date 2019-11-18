<%
Sub BuildArray(objDict, aTempArray) 
  Dim nCount, strKey
  nCount = 0
  
  '-- Redim the array to the number of keys we need 
  Redim aTempArray(objDict.Count - 1)

  '-- Load the array 
  For Each strKey In objDict.Keys

    '-- Set the array element to the key 
    aTempArray(nCount) = strKey 

    '-- Increment the count 
    nCount = nCount + 1

  Next 
End Sub


Sub SortArray(aTempArray) 
  Dim iTemp, jTemp, strTemp

  For iTemp = 0 To UBound(aTempArray)  
    For jTemp = 0 To iTemp  

      If strComp(aTempArray(jTemp),aTempArray(iTemp)) > 0 Then
        'Swap the array positions
        strTemp = aTempArray(jTemp) 
        aTempArray(jTemp) = aTempArray(iTemp) 
        aTempArray(iTemp) = strTemp 
      End If 

    Next 
  Next 
End Sub


Sub PrintDictionary(objDict, aTempArray) 
  Dim iTemp 
  For iTemp = 0 To UBound(aTempArray) 
    Response.Write(aTempArray(iTemp) & " - " & _
                   objDict.Item(aTempArray(iTemp)) & "<br>") 
  Next 
End Sub


Sub PrintSortedDictionary(objDict)
  Dim aTemp
  Call BuildArray(objDict, aTemp) 	' Build the array
  Call SortArray(aTemp) 	' Sort the array 
  Call PrintDictionary(objDict, aTemp) ' Print the dictionary using the array as an index 
End Sub



Dim dObj, aTemp 	' Create our dictionary variable name and the temporary array

Set dObj = Server.CreateObject("Scripting.Dictionary") 

'-- Get some values
dObj.Add "Apple", "Value1" 
dObj.Add "Orange", "Value2" 
dObj.Add "Banana", "Value3"
dObj.Add "Grapefruit", "Value4" 
dObj.Add "Avacado", "Value5"

Call PrintSortedDictionary(dObj)

Set dObj = Nothing	' Dereference the object
%>
