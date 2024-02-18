<%@ codepage="65001" language="VBScript" %>
<%
option Explicit
Response.CharSet = "utf-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seachkeywordCls.asp" -->
<%

dim iUserIP : iUserIP = Request.ServerVariables ("REMOTE_ADDR")
if NOT (iUserIP="172.16.0.206" or LEFT(iUserIP,11)="61.252.133." or iUserIP="192.168.50.10") then
    dbget.Close() : response.end
end if

'// ============================================================================
dim osearchKeyword

set osearchKeyword = new CSearchKeyword

osearchKeyword.GetRelatedKeywordList_API

dim arr : arr = osearchKeyword.FResultArray

If Not IsArray(arr) Then
	dbget.Close() : response.end
end if

dim rows, row, currKeyword
rows = UBound(arr, 2)
currKeyword = ""
For row = 0 to rows
	if (LCASE(currKeyword) <> LCASE(arr(0, row))) then
		if (currKeyword = "") then
			Response.Write(arr(0, row) & ":" & arr(1, row))
		else
			Response.Write(vbCrLf & arr(0, row) & ":" & arr(1, row))
		end if
		currKeyword = arr(0, row)
	else
		Response.Write("," & arr(1, row))
	end if
	
	if (row mod 10000)=0 then response.flush  ''2018/07/18
Next

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
