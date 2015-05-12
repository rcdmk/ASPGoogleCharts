<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8">
	<title>Google Charts</title>
</head>
<body>
	<!--#include file="GoogleCharts.class.asp" -->
	<h1>Google Charts</h1>
	<%
	dim chart
	set chart = new GoogleCharts
	
	chart.type = CHART_COLUMN
	
	chart.title = ""
	
	chart.addColumn CTYPE_STRING, "Type"
	chart.addColumn CTYPE_NUMBER, "Qty"
	chart.addColumn CTYPE_NUMBER, "Price"
	
	chart.addRow Array("Peperony", 2, 1.2)
	chart.addRow Array("Marguerita", 1, 3.5)
	chart.addRow Array("Bacon", 4, 2.25)
	
	chart.draw
	%>
</body>
</html>