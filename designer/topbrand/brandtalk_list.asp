<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/topbrand/topbrandtalkcls.asp" -->
<%

dim i,ix, page, iscurrtopbrand

page = requestCheckVar(request("page"),10)
if (page = "") then
    page = "1"
end if


'==============================================================================
dim otopbrandtalklist
set otopbrandtalklist = New CTopBrandTalk

iscurrtopbrand = otopbrandtalklist.IsCurrentTopBrand(session("ssBctId"))

otopbrandtalklist.FRectMakerID = session("ssBctId")
otopbrandtalklist.FCurrPage = page
'otopbrandtalklist.FRectIsCurrentTopBrand = "Y"

otopbrandtalklist.GetTopBrandTalkList


if ((iscurrtopbrand = false) and (session("ssBctId") <> "test")) then
    response.write "<script>alert('탑브랜드 전용 메뉴입니다.');</script>"
    dbget.close()	:	response.End
end if

%>

<script language="javascript" type="text/javascript">

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="f" action="brandtalk_write.asp" method=get onsubmit="return false">
    <input type=hidden name=menupos value="<%= menupos %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="icon" value="#">
			<font color="red"><b>브랜드토크 리스트</b></font>
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
        <td width="40">Idx</td>
        <td width="55">이미지</td>
        <td>내용</td>
        <td width="80">등록일</td>
    </tr>
<% for i = 0 to (otopbrandtalklist.FResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= otopbrandtalklist.FItemList(i).Fidx %></td>
        <td><a href="brandtalk_modify.asp?idx=<%= otopbrandtalklist.FItemList(i).Fidx %>"><img src='<%= otopbrandtalklist.FItemList(i).Ficon1 %>' border="0"></a></td>
        <td align="left"><a href="brandtalk_modify.asp?idx=<%= otopbrandtalklist.FItemList(i).Fidx %>"><%= DDotFormat(db2html(otopbrandtalklist.FItemList(i).Fimagetalk),40) %></a></td>
        <td><%= Left(otopbrandtalklist.FItemList(i).Fregdate,10) %></td>
    </tr>
<% next %>
<% if (otopbrandtalklist.FResultCount < 1) then %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="25" colspan="4">검색결과가 없습니다.</td>
    </tr>
<% end if %>


    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if otopbrandtalklist.HasPreScroll then %>
				<a href="javascript:NextPage('<%= otopbrandtalklist.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
	
			<% for ix=0 + otopbrandtalklist.StartScrollPage to otopbrandtalklist.FScrollCount + otopbrandtalklist.StartScrollPage - 1 %>
				<% if ix>otopbrandtalklist.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
			<% next %>
	
			<% if otopbrandtalklist.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>

<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
</form>
<!-- 표 하단바 끝-->
<%

set otopbrandtalklist = Nothing

%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->