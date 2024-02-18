<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/topbrand/topbrandnewscls.asp" -->
<%

dim i, page, iscurrtopbrand

page = requestCheckVar(request("page"),10)
if (page = "") then
    page = "1"
end if


'==============================================================================
dim otopbrandnewslist
set otopbrandnewslist = New CTopBrandNews

iscurrtopbrand = otopbrandnewslist.IsCurrentTopBrand(session("ssBctId"))

otopbrandnewslist.FRectMakerID = session("ssBctId")
otopbrandnewslist.FCurrPage = page
'otopbrandnewslist.FRectIsCurrentTopBrand = "Y"

otopbrandnewslist.GetTopBrandNewsList


if ((iscurrtopbrand = false) and (session("ssBctId") <> "test")) then
    response.write "<script>alert('탑브랜드 전용 메뉴입니다.');</script>"
    dbget.close()	:	response.End
end if

%>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="f" action="brandnews_write.asp" method=get onsubmit="return false">
    <input type=hidden name=menupos value="<%= menupos %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="icon" value="#">
			<font color="red"><b>브랜드뉴스 리스트</b></font>
			&nbsp;
			<% if (iscurrtopbrand = false) then %>
	        	<font color=red><b>현재 탑브랜드가 아닙니다.</b></font>
			<% end if %>
			&nbsp;
	        <input type="button" class="button" value="등록하기" onClick="document.f.submit();">
		</td>
	</tr>
	</form>
	
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="40" align="center">Idx</td>
        <td>제목</td>
        <td width="80" align="center">등록일</td>
    </tr>
<% for i = 0 to (otopbrandnewslist.FResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF">
        <td height="25" align="center"><%= otopbrandnewslist.FItemList(i).Fidx %></td>
        <td align="left"><a href="brandnews_modify.asp?idx=<%= otopbrandnewslist.FItemList(i).Fidx %>"><%= DDotFormat(otopbrandnewslist.FItemList(i).Ftitle,40) %></a></td>
        <td align="center"><%= Left(otopbrandnewslist.FItemList(i).Fregdate,10) %></td>
    </tr>
<% next %>
<% if (otopbrandnewslist.FResultCount < 1) then %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="25" colspan="3">검색결과가 없습니다.</td>
    </tr>
<% end if %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			&nbsp;
		</td>
	</tr>
</table>


<%

set otopbrandnewslist = Nothing

%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->