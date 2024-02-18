<%@ codepage="65001" language="VBScript" %>
<%
option Explicit
Response.CharSet = "utf-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim iUserIP : iUserIP = request.ServerVariables("REMOTE_ADDR")

if NOT (iUserIP="172.16.0.206" or LEFT(iUserIP,10)="192.168.1.") then
    dbget.Close() : response.end
end if

Dim vQuery, vArr, i, vUseYN, vArrKy
vUseYN = requestCheckVar(Request("useyn"),1)
If vUseYN = "" Then
	vUseYN = "y"
End If

'vQuery = ""
'vQuery = vQuery & "SELECT title, (isNull(url_pc,'') + '$$' + isNull(url_m,'')) as meta1, (autotype + '$$' + icon) as meta2 "
'vQuery = vQuery & "FROM [db_sitemaster].[dbo].[tbl_search_autocomplete] WHERE useyn = '" & vUseYN & "' and autotype = 'ky' "
'vQuery = vQuery & "Order by idx asc"
'rsget.CursorLocation = adUseClient
'rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
'
'If Not rsget.Eof Then
'	vArrKy = rsget.getRows()
'End If
'rsget.close

vQuery = "EXEC db_EVT.dbo.usp_konan_autocomplete_keyword_get"
rsEVTget.CursorLocation = adUseClient
rsEVTget.Open vQuery,dbEVTget,adOpenForwardOnly,adLockReadOnly

If Not rsEVTget.Eof Then
	vArrKy = rsEVTget.getRows()
End If
rsEVTget.close
'''---------------------------------------------------------------------------------------

vQuery = ""
vQuery = vQuery & "SELECT title, (isNull(url_pc,'') + '$$' + isNull(url_m,'')) as meta1, (autotype + '$$' + icon) as meta2 "
vQuery = vQuery & "FROM [db_sitemaster].[dbo].[tbl_search_autocomplete] WHERE useyn = '" & vUseYN & "' and autotype <> 'ky' "
vQuery = vQuery & "Order by idx desc"
rsget.CursorLocation = adUseClient
rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly

If Not rsget.Eof Then
	vArr = rsget.getRows()
End If
rsget.close
%>
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
If isArray(vArrKy) Then
	For i=0 To UBound(vArrKy,2)
		If vUseYN = "y" Then
			Response.Write fnReplaceWord(vArrKy(0,i)) & ":" & vArrKy(1,i) & ":" & vArrKy(2,i) & vbCrLf
		ElseIf vUseYN = "n" Then
			Response.Write fnReplaceWord(vArrKy(0,i)) & vbCrLf
		End If
	Next
End If

If isArray(vArr) Then
	For i=0 To UBound(vArr,2)
		If vUseYN = "y" Then
			Response.Write fnReplaceWord(vArr(0,i)) & ":" & vArr(1,i) & ":" & vArr(2,i) & vbCrLf
		ElseIf vUseYN = "n" Then
			Response.Write fnReplaceWord(vArr(0,i)) & vbCrLf
		End If
	Next
End If


Function fnReplaceWord(v)
	Dim vTmp
	vTmp = v
	'vTmp = Replace(vTmp, "&nbsp;", "")
	vTmp = Replace(vTmp, ":", "")
	fnReplaceWord = vTmp
End Function
%>