<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'####################################################
' Description : 포토그래퍼 , 스타일리스트 검색
' History : 2018.01.26 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
<%
dim searchtype
	searchtype = requestcheckvar(request("searchtype"),32)

if searchtype="" then
	dbget.close() : response.end
end if

'/ 아작스라서 utf8이 기본인데.. 클래스 파일을 utf8 로 바꿀수가 없어서 따로땀.
Sub SelectUserajax(AA, BB, CC)
	Dim query1
	query1 = " select user_no, user_id, user_name from [db_partner].[dbo].tbl_photo_user"
	query1 = query1 + " where user_type='"&AA&"' and user_useyn = 'Y'"
	rsget.Open query1,dbget,1
%>
	<select class="select" name='<%=BB%>'>
		<%= chkIIF(AA = "1","<option value='0'>-- 포토그래퍼 선택 --</option>","<option value=0>-- 스타일리스트 선택 --</option>") %>
<%
	If not rsget.EOF Then
		rsget.Movefirst
		Do until rsget.EOF
			response.write("<option value='"&rsget("user_id")& "' "& chkIIF(CC = rsget("user_id"),"selected","") &">" & rsget("user_name") & "" & "</option>")
			rsget.MoveNext
		Loop
	End If
	rsget.close
	response.write("</select>")
End Sub
%>
<% if searchtype = "req_photo" then %>
	<% call SelectUserajax("1", "req_photo", "") %>
<% elseif searchtype = "req_Stylist" then %>
	<% call SelectUserajax("2", "req_Stylist", "") %>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<% session.codePage = 949 %>