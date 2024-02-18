<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/artistsroomcls.asp"-->
<%
dim oartistsroom
dim page
page = RequestCheckvar(request("page"),10)
if (page="") then page=1

set oartistsroom = new CArtistsRoom
oartistsroom.FPageSize=50
oartistsroom.FCurrPage = page
oartistsroom.GetArtistRoomList


dim i

%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
    <tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="30" >
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">

        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22"  border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#DDDDFF">
		<td colspan="1"><a href="/academy/artistroom/artistroomreg.asp" target="_blank"><img src="/images/icon_new_registration.gif" width="75" border="0"></a></td>
		<td colspan="12" align="right">검색건수 : <%= oartistsroom.FTotalCount %> 건 Page : <%= page %>/<%= oartistsroom.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="70">강사ID</td>
		<td align="center" width="70">강사명</td>
		<td align="center">강사타이틀</td>
		<td align="center" >등록일</td>
	</tr>
<% for i=0 to oartistsroom.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF">
		<td><a href="/academy/artistroom/artistroomreg.asp?lecuserid=<%= oartistsroom.FItemList(i).Flecuserid %>" target="_blank"><%= oartistsroom.FItemList(i).Flecuserid %></a></td>
		<td><%= oartistsroom.FItemList(i).Fsocname_kor %></td>
		<td><%= oartistsroom.FItemList(i).Ftitle %></td>
		<td align="center"><%= oartistsroom.FItemList(i).Fregdate  %></td>
	</tr>
<% next %>
</table>

<%
set oartistsroom = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbacademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->