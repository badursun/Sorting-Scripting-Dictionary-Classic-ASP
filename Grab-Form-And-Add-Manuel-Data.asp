<!--#include file="sort-class.asp"--><%
Set SortForm = New sortk
	
	' Using Grab Form And Add Manuel Data 
	'---------------------------------------------

	' First Create Dictionary
	Set MyDictionary = SortForm.CreateDictionary

	' Add Some Data To Dictionary Manuelly
	With SortForm
		.AddData MyDictionary ,"C_ORDER_VAL_1", "Manuel Test Value"
		.AddData MyDictionary ,"Z_ORDER_VAL_2", "Manuel Tewt Value"
		.AddData MyDictionary ,"G_ORDER_VAL_3", "Manuel Tewt Value"
		.AddData MyDictionary ,"A_ORDER_VAL_1", "Manuel Test Value"
	End With

	' Grab Form Datas
	SortForm.GrabForms( MyDictionary )


	' Print Values Non-Ordered
	Response.Write "<h4>UNSORTED VALUES</h4>"
	SortForm.PrintSortedDictionary( MyDictionary )

	Response.Write "<hr />"

	' Print Values Ordered
	Response.Write "<h4>SORTED VALUES</h4>"
	SortForm.PrintUnSortedDictionary( MyDictionary )

Set SortForm = Nothing
%>
