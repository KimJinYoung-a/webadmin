<%
dim yyyy2

yyyy2   = request("yyyy2")

If (UBound(Split(yyyy2, ",")) > 0) Then
	response.Write "ERROR"
	response.End
End If

%>
