<% option Explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim prdno, tenitemid
prdno = requestCheckVar(request("prdno"),10)

dim sqlStr
sqlStr = "select top 1 itemid from  [db_item].[dbo].tbl_interpark_reg_item"
sqlStr = sqlStr & " where interparkPrdNo='" & prdno & "'"

rsget.Open sqlStr, dbget, 1
if Not rsget.Eof then
    tenitemid = rsget("itemid")
end if
rsget.Close
%>
<p>
<small>
<table width="100%" height="20" border=0 cellspacing=0 >
<form name="frm" method=get action="">
<tr>
    <td><input type="text" name="prdno" value="<%= prdno %>" size="8" maxlength="10"><small>(8자리)</small></td>
</tr>
</form>
</table>
<%
dim iframeSRC
if (tenitemid<>"") then
    iframeSRC = "http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & tenitemid
    
    response.write "상품번호 : " & tenitemid
else
    iframeSRC = ""
    response.write "검색결과가 없습니다."
end if
%>

<table width="100%" height="100%" border=1 cellspacing=1 >
    <tr>
        <td>
            <iframe src="<%= iframeSRC %>" width="100%" height="100%" frameborder=1 scrolling=yes marginheight=0 marginwidth=0 align=center></iframe>
        </td>
    </tr>
</table>
</small>
<!-- #include virtual="/lib/db/dbclose.asp" --> 