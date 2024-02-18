<%@  codepage="949" language="VBScript" %><% option explicit %><?xml version="1.0" encoding="euc-kr"?>
<% Session.CodePage = 949 %>
<% Response.contentType = "text/xml; charset=euc-kr" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/displaycateCls.asp"-->
<response>
<%
'// 값비교 후 Return 값 like iif function
Function ChkIIF(trueOrFalse, trueVal, falseVal)
	if (trueOrFalse) then
	    ChkIIF = trueVal
	else
	    ChkIIF = falseVal
	end if
End Function
function db2html(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&amp;", "&")
    v = replace(v, "&lt;", "<")
    v = replace(v, "&gt;", ">")
    v = replace(v, "&quot;", "'")
    v = Replace(v, "", "<br>")
    v = Replace(v, "\0x5C", "\")
    v = Replace(v, "\0x22", "'")
    v = Replace(v, "\0x25", "'")
    v = Replace(v, "\0x27", "%")
    v = Replace(v, "\0x2F", "/")
    v = Replace(v, "\0x5F", "_")
    ''checkvalue = Replace(checkvalue, vbcrlf,"<br>")
    db2html = v
end Function

dim param1
param1 = request("param1")

dim cDisp, i

	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 2
	cDisp.FRectCateCode = param1
	cDisp.GetDispCateList()

	If cDisp.FResultCount > 0 Then
		For i=0 To cDisp.FResultCount-1
			response.write "<item>" + VbCrlf
			response.write "<value1>" + Cstr(cDisp.FItemList(i).FCateCode) + "</value1>" + VbCrlf
			if cDisp.FItemList(i).FUseYN="N" then
				response.write "<value2>" & (cDisp.FItemList(i).FCateName) & " (오픈예정)</value2>" & VbCrlf
			else
				response.write "<value2>" & (cDisp.FItemList(i).FCateName) &"</value2>" & VbCrlf
			end if
			response.write "</item>" + VbCrlf
		Next
	End If

SET cDisp = Nothing
%>
</response>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->