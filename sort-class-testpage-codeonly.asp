<!--#include file="sort-class.asp"--><%
Set SortForm = New sortk
	
	Response.Write "<h4>UNSORTED FORM VALUES</h4>"
	SortForm.PrintSortedDictionary( SortForm.GrabForms() )

	Response.Write "<hr />"

  Response.Write "<h4>SORTED FORM VALUES</h4>"
	SortForm.PrintUnSortedDictionary( SortForm.GrabForms() )
  
Set SortForm = Nothing
%>
