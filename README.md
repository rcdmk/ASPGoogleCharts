# ASP GoogleCharts
A classic ASP wrapper for the Google Charts JS API

Simple to use:

	dim chart
	set chart = new GoogleCharts

	chart.type = CHART_COLUMN
	
	chart.title = "Friday Night Pizza"
	
	chart.addColumn CTYPE_STRING, "Type"
	chart.addColumn CTYPE_NUMBER, "Qty"
	chart.addColumn CTYPE_NUMBER, "Price"
	
	chart.addRow Array("Peperony", 2, 1.2)
	chart.addRow Array("Marguerita", 1, 3.5)
	chart.addRow Array("Bacon", 4, 2.25)
	
	chart.draw

Accepts loading data from bidimensional arrays, like the ones provided from the `GetRows()` method of `ADODB Recordset`s:

	dim chart
	set chart = new GoogleCharts

	chart.type = CHART_COLUMN
	
	chart.title = "Friday Night Pizza"
	
	chart.addColumn CTYPE_STRING, "Type"
	chart.addColumn CTYPE_NUMBER, "Qty"
	chart.addColumn CTYPE_NUMBER, "Price"
	
	dim rs
	set rs = createObject("ADODB.Recordset")
	rs.open "SQL HERE", yourConnection, 0, 1
	
	chart.loadArray rs.getRows()
	chart.draw
	
	set rs = nothing
	
It can also load data from `Recordset`s:

	dim chart
	set chart = new GoogleCharts

	chart.type = CHART_COLUMN
	
	chart.title = "Friday Night Pizza"

	dim rs
	set rs = createObject("ADODB.Recordset")
	rs.open "SQL HERE", yourConnection, 0, 1
	
	' No need to declare columns. It gets the type and label from the `Recordset.Fields` property.
	
	chart.loadRecordSet rs
	chart.draw
	
	set rs = nothing
	
## Licence

The MIT License (MIT)
Copyright (c) 2012 RCDMK - rcdmk[at]hotmail[dot]com

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


	
