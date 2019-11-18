<!--#include file="sort-class.asp"-->
<%Set SortForm = New sortk%>

<!DOCTYPE html>
<html>
<head>
	<title>Scripting Dictionary Sort Class</title>
</head>
<body>
<p>Post This Page Form Value End See Result</p>
<table width="800" align="center" border="1">
	<tr>
		<td width="50%"><h4>UNSORTED FORM VALUES</h4></td>
		<td width="50%"><h4>SORTED FORM VALUES</h4></td>
	</tr>
	<tr>
		<td width="50%"><%SortForm.PrintUnSortedDictionary( SortForm.GrabForms() )%></td>
		<td width="50%"><%SortForm.PrintSortedDictionary( SortForm.GrabForms() )%></td>
	</tr>
</table>
</body>
</html>
<% Set SortForm = Nothing %>
