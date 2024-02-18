<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%><%

Dim test As System.Drawing.Image

test.NewImage 200, 200, "FFFFFF"
test.Rectangle 10, 10, 90, 90
test.PenColor = "FF0000"
test.PenThickness = 5
test.BrushColor = "0000FF"
test.BrushStyle = 0
test.Rectangle 110, 10, 190, 90
test.PenThickness = 1
test.Line 0, 100, 100, 199
test.Line 100, 199, 199, 100
test.BrushColor = "00FF00"
test.FloodFill 10, 190

%>
