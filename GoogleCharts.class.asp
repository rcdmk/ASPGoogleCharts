<%
' GoogleChart utility class 0.1 - May, 11th - 2015
'
' Licence:
' The MIT License (MIT)
' Copyright (c) 2012 RCDMK - rcdmk[at]hotmail[dot]com
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
' associated documentation files (the "Software"), to deal in the Software without restriction,
' including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
' subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all copies or substantial
' portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
' NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
' SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


' Constants
const CHART_BAR		= "BarChart"
const CHART_COLUMN 	= "ColumnChart"
const CHART_LINE 	= "LineChart"
const CHART_PIE 	= "PieChart"

const CTYPE_BOOL		= "boolean"
const CTYPE_DATE		= "date"
const CTYPE_DATETIME	= "datetime"
const CTYPE_NUMBER 		= "number"
const CTYPE_STRING 		= "string"
const CTYPE_TIME		= "timeofday"

class GoogleCharts
	dim i_type, i_title, i_subtitle, i_id, i_width, i_height
	dim i_columns, i_data, i_scriptPath
	
	public property get [Type]()
		[Type] = i_type
	end property
	
	public property let [Type](byval value)
		select case value
			case CHART_BAR, CHART_COLUMN, CHART_LINE, CHART_PIE			
				i_type = value
			case else:
				err.raise 1, typeName(me), "Invalid chart type. Valid types are: CHART_BAR, CHART_COLUMN, CHART_LINE, CHART_PIE"
		end select		
	end property
	
	public property get ID()
		ID = i_id
	end property
	
	public property let ID(byval value)
		i_id = value
	end property
	
	public property get Title()
		Title = i_title
	end property
	
	public property let Title(byval value)
		i_title = value
	end property
	
	public property get Subtitle()
		Subtitle = i_subtitle
	end property
	
	public property let Subtitle(byval value)
		i_subtitle = value
	end property
	
	public property get Width()
		Width = i_width
	end property
	
	public property let Width(byval value)
		i_width = value
	end property
	
	public property get Height()
		Height = i_height
	end property
	
	public property let Height(byval value)
		i_height = value
	end property
	
	public property get ScriptPath()
		ScriptPath = i_scriptPath
	end property
	
	public property let ScriptPath(byval value)
		i_scriptPath = value
	end property
	
	
	sub Class_Initialize()
		i_type = CT_BAR
		Randomize
		i_id = int(rnd() * timer)
		i_title = "Google Chart " & i_id
		i_subtitle = ""
		i_width = 400
		i_height = 300
		
		i_scriptPath = "https://www.google.com/jsapi"
		
		redim i_columns(-1)
		redim i_data(-1)
	end sub
	
	sub Class_Terminate()
		redim i_data(-1,-1)
	end sub
	
	public function AddColumn(byval dataType, byval label)
		dim col
		set col = new GoogleChartsColumn
		col.Type = dataType
		col.Label = label
		
		arrayPush i_columns, col
		
		set AddColumn = col
	end function
	
	public sub AddRow(byval values)
		dim row
		
		if isArray(values) then
			row = values
		else
			row = Array(values)
		end if
		
		arrayPush i_data, row
	end sub
	
	public sub ClearColumns()
		dim col
		
		for each col in i_columns
			set col = nothing
		next
	
		redim i_columns(-1)
	end sub
	
	public sub ClearData()
		redim i_data(-1)
	end sub
	
	public sub LoadRecordset(byref rs)
		if TypeName(rs) <> "Recordset" then err.raise 3, TypeName(me), "Invalid object type. Accepted type is ADODB.Recordset"
	
		ClearColumns()
		ClearData()
		
		dim field, row
		
		for each field in rs.fields
			AddColumn getFieldType(field.type), field.Name
		next
		
		if not rs.bof then rs.moveFirst
		
		while not rs.eof
			redim row(-1)
			
			for each field in rs.fields
				ArrayPush row, field.value
			next
			
			AddRow row
			
			rs.movenext
		wend
	end sub
	
	function getFieldType(byval fieldType)
		select case fieldType
			case 2, 3, 4, 5, 6, 7, 14, 16, 17, 18, 19, 20, 21, 131, 139
				getFieldType = CTYPE_NUMBER
				
			case 11
				getFieldType = CTYPE_BOOL
				
			case 64, 134, 135
				getFieldType = CTYPE_TIME
				
			case 133
				getFieldType = CTYPE_DATE
			
			case else
				getFieldType = CTYPE_STRING
		end select
	end function
	
	
	public sub LoadArray(byref arr)
		if TypeName(arr) <> "Variant()" then err.raise 4, TypeName(me), "Invalid object type. Accepted type is Array() - Variant"
		
		if arrayDimensions(arr) <> 2 then err.raise 5, TypeName(me), "Invalid Array Size. Array must be bidimensional"
		
		ClearData()
		
		dim r, c, row
		
		for r = 0 to ubound(arr, 2)
			redim row(-1)
			
			for c = 0 to ubound(arr, 1)
				arrayPush row, arr(c, r)
			next
			
			AddRow row
		next
	end sub
	
	' Pushes (adds) a value to an array, expanding it
	private function arrayPush(byref arr, byref value)
		redim preserve arr(ubound(arr) + 1)
		
		if isobject(value) then
			set arr(ubound(arr)) = value
		else
			arr(ubound(arr)) = value
		end if
		
		ArrayPush = arr
	end function
	
	private function arrayDimensions(byref arr)
		dim dimensions
		dimensions = 0
		
		on error resume next
		while err.number = 0
			dimensions = ubound(arr, dimensions + 1) + 1
		wend
		on error goto 0
		
		arrayDimensions = dimensions - 1
	end function
	
	public sub Draw()
		dim curLCID
		curLCID = Session.LCID
		Session.LCID = 1033
		%>
<script type="text/javascript"  id="GC_script_tag_<%= i_id %>">
	if (typeof(google) == "undefined") {
		var s = document.createElement("script");
		s.src = '<%= i_scriptPath %>?callback=GC<%= i_id %>';
		
		var body = document.getElementsByTagName('body')[0];
		body.insertBefore(s, body.childNodes[0]);
	} else {
		GC<%= i_id %>();
	}

	function GC<%= i_id %>() {
		if (typeof(google) == "undefined") return;
	
		// Load the Visualization API and the piechart package.
		google.load('visualization', '1.0', {'packages':['corechart', 'bar'], 'callback':drawChart});

		// Callback that creates and populates a data table,
		// instantiates the chart, passes in the data and
		// draws it.
		function drawChart() {
			// Create the data table.
			var data = new google.visualization.DataTable();
			<%
			dim col
			for each col in i_columns
				%>
			data.addColumn({
				type: '<%= col.Type %>',
				label: '<%= escapeText(col.Label) %>',
				id: '<%= col.ID %>',
				role: '<%= col.Role %>',
				pattern: '<%= col.Pattern %>'
			});
				<%
			next
			%>
			data.addRows([
				<%
				dim r, row, c
				for r = 0 to ubound(i_data)
					row = i_data(r)
					%>
					[
						<%
						for c = 0 to ubound(row)
							if c <= ubound(i_columns) then
								if i_columns(c).Type = CTYPE_STRING then
									response.write "'"
									response.write escapeText(row(c))
									response.write "'"
								else
									response.write row(c)
								end if
							end if
							
							if c < ubound(row) then response.write ","
						next
						%>
					]<%
					if r < ubound(i_data) then response.write ","
				next
				%>
			]);

			// Set chart options
			var options = {
				'title':'<%= escapeText(i_title) %>',
				'subtitle':'<%= escapeText(i_subtitle) %>',
				'width':<%= i_width %>,
				'height':<%= i_height %>
			};
			
			// Instantiate and draw our chart, passing in some options.
			var element = document.getElementById('chart_div_<%= i_id %>');
			var chart = new google.visualization.<%= i_type %>(element);
			chart.draw(data, options);
		}
	};
</script>
<!--Div that will hold the chart-->
<div id="chart_div_<%= i_id %>" class="GoogleChart"></div>	
<%
		Session.LCID = curLCID
	end sub
	
	function escapeText(byval text)
		dim result
		result = text
		
		result = replace(result, "'", "\'")
		result = replace(result, """", "\""")
		result = replace(result, vbCr, "\r")
		result = replace(result, vbLf, "\n")
		
		escapeText = result
	end function
end class

class GoogleChartsColumn
	dim i_type, i_label, i_role, i_id, i_pattern
	
	public property get [Type]()
		[Type] = i_type
	end property
	
	public property let [Type](byval value)
		select case value
			case CTYPE_BOOL, CTYPE_DATE, CTYPE_DATETIME, CTYPE_NUMBER, CTYPE_STRING, CTYPE_TIME
				i_type = value
			case else
				err.raise 2, typeName(me), "Invalid column type. Valid types are: CTYPE_BOOL, CTYPE_DATE, CTYPE_DATETIME, CTYPE_NUMBER, CTYPE_STRING, CTYPE_TIME" 
		end select
	end property
	
	public property get Label
		Label = i_label
	end property
	
	public property let Label(byval value)
		i_label = value
	end property
	
	public property get Role()
		Role = i_role
	end property
	
	public property let Role(byval value)
		i_role = value
	end property
	
	public property get ID()
		ID = i_id
	end property
	
	public property let ID(byval value)
		i_id = value
	end property
	
	public property get Pattern()
		Pattern = i_pattern
	end property
	
	public property let Pattern(byval value)
		i_pattern = value
	end property
	
	
	sub Class_Initialize()
		i_type = CTYPE_STRING
		i_label = "column"
		i_role = ""
		i_id = ""
		i_pattern = ""
	end sub
	
	sub Class_Terminate()
	end sub
end class
%>