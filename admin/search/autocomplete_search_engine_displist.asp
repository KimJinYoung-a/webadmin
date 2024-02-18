<%@ codepage="65001" language="VBScript" %>
<%
option Explicit
Response.CharSet = "utf-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim iUserIP : iUserIP = request.ServerVariables("REMOTE_ADDR")

if NOT (iUserIP="172.16.0.206" or LEFT(iUserIP,10)="192.168.1.") then
    dbget.Close() : response.end
end if

Dim vQuery, vArr, i, vUseYN
vUseYN = requestCheckVar(Request("useyn"),1)
If vUseYN = "" Then
	vUseYN = "y"
End If

vQuery = ""
vQuery = vQuery & "SELECT catename, catecode as meta1, isNull(([db_item].[dbo].[getCateCodeFullDepthName](catecode)),'') as meta2 "
vQuery = vQuery & "FROM [db_item].[dbo].[tbl_display_cate] WHERE useyn = '" & vUseYN & "' "
vQuery = vQuery & "Order by depth asc, catecode asc, sortNo asc"
rsget.CursorLocation = adUseClient
rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly

If Not rsget.Eof Then
	vArr = rsget.getRows()
End If
rsget.close
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
If isArray(vArr) Then
	For i=0 To UBound(vArr,2)
		If vUseYN = "y" Then
			Response.Write fnReplaceWord(vArr(0,i)) & ":" & vArr(1,i) & ":" & fnReplaceWord(vArr(2,i)) & vbCrLf
		ElseIf vUseYN = "n" Then
			Response.Write vArr(0,i) & vbCrLf
		End If
	Next
End If

Function fnReplaceWord(v)
	Dim vTmp
	vTmp = v
	vTmp = Replace(vTmp, "&nbsp;", "")
	vTmp = Replace(vTmp, ":", "")
	fnReplaceWord = vTmp
End Function
%>