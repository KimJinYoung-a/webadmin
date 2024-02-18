<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 매출집계- 관심등록전환매출
' History : 2016.10.06 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/report/maechul/statisticCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->

<%
Dim i, cStatistic, vSiteName, itemid, iTotCnt

	vSiteName = RequestCheckvar(request("sitename"),16)
	itemid = RequestCheckvar(request("itemid"),10)
	if vSiteName = "" then vSiteName = "diyitem"



Set cStatistic = New cacademyStatic_list
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRectItemid   = itemid 
	cStatistic.fStatistic_wish_UserList()

'response.end
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>Userid</td>
</tr>
<% if cStatistic.FTotalCount>0 then %>
	<% For i = 0 To cStatistic.FTotalCount -1 %>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
		<td><%= cStatistic.FItemList(i).FMakerID %></td>
	</tr>
	<% Next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">검색결과가 없습니다.</td>
	</tr>
<% end if %>
</table>

<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->