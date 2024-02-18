<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%

dim arr, i, v, t

arr = Array(1, 2, 3, Array(2,3,4), 6)

for i = 0 to UBound(arr)
	if IsNumeric(arr(i)) then
		response.write "숫자 : " & arr(i) & "<br />"
	elseif IsArray(arr(i)) then
		v = 0
		for each t in  arr(i)
			v = v + t
		next
		response.write "배열 : " & v & "<br />"
	end if
next

%>12
