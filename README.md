# Sorting-Scripting-Dictionary-Classic-ASP
Sorting a Dictionary Object By Aaron A.

## How to sort dictionary by Key
This code and class Sort an arrays / form collections / collections by key name like PHP ksort function.
See GrabForms property, grab each Request.Form collecions key and value, add scripting dictionary and sorting.

You are free to develop as you wish. It is prepared to show you and simplify the way you need it.

# How To Use

## Using Auto Grab POST FORM Datas
<%
```asp
Set SortForm = New sortk

	' Grab Form / First Create Dictionary
	Set MyDictionary = SortForm.GrabForms( Null )

	' Print Values Non-Ordered
	Response.Write "<h4>UNSORTED VALUES</h4>"
	SortForm.PrintSortedDictionary( MyDictionary )

	Response.Write "<hr />"

	' Print Values Ordered
	Response.Write "<h4>SORTED VALUES</h4>"
	SortForm.PrintUnSortedDictionary( MyDictionary )

Set SortForm = Nothing
```
%>

## Using Auto Grab POST FORM Datas And Add Manually
<%
```asp
Set SortForm = New sortk
	
	' Using Grab Form And Add Manually Data 
	'---------------------------------------------

	' First Create Dictionary
	Set MyDictionary = SortForm.CreateDictionary

	' Add Some Data To Dictionary
	With SortForm
		.AddData MyDictionary ,"C_ORDER_VAL_1", "Manuel Test Value"
		.AddData MyDictionary ,"Z_ORDER_VAL_2", "Manuel Tewt Value"
		.AddData MyDictionary ,"G_ORDER_VAL_3", "Manuel Tewt Value"
		.AddData MyDictionary ,"A_ORDER_VAL_1", "Manuel Test Value"
	End With

	' Grab Form
	SortForm.GrabForms( MyDictionary )


	' Print Values Non-Ordered
	Response.Write "<h4>UNSORTED VALUES</h4>"
	SortForm.PrintSortedDictionary( MyDictionary )

	Response.Write "<hr />"

	' Print Values Ordered
	Response.Write "<h4>SORTED VALUES</h4>"
	SortForm.PrintUnSortedDictionary( MyDictionary )

Set SortForm = Nothing
```
%>
