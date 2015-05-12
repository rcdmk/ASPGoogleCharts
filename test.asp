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
	
	chart.id = 1 ' if not defined, it uses a random number
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
	<h2>Data from ADO Recordset</h2>
	<!--
	   METADATA    
	   TYPE="TypeLib"    
	   NAME="Microsoft ActiveX Data Objects 2.5 Library"    
	   UUID="{00000205-0000-0010-8000-00AA006D2EA4}"    
	   VERSION="2.5"
	-->
	<%
	dim rs
	set rs = createObject("ADODB.Recordset")
	
	' prepera an in memory recordset 
	' could be, and mostly, loaded from a database
	rs.CursorType = adOpenKeyset
	rs.CursorLocation = adUseClient
	rs.LockType = adLockOptimistic
	
	rs.Fields.Append "Type", adVarChar, 50, adFldKeyColumn
	rs.Fields.Append "Qty", adInteger, , adFldMayBeNull
	rs.Fields.Append "Price", adDecimal, 14, adFldMayBeNull
	rs.Fields("Price").NumericScale = 2
	
	rs.Open
	
	rs.AddNew
	rs("Type") = "Peperony"
	rs("Qty") = 2
	rs("Price") = 1.2
	rs.Update
	
	rs.AddNew
	rs("Type") = "Marguerita"
	rs("Qty") = 1
	rs("Price") = 3.5
	rs.Update
	
	rs.AddNew
	rs("Type") = "Bacon"
	rs("Qty") = 4
	rs("Price") = 2.25
	rs.Update
	 
	chart.id = 2 ' change id to draw a new chart with the same object
	chart.loadRecordSet rs
	
	set rs = nothing
	
	chart.draw
	%>
</body>
</html>