<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.04.15 한용민 생성
'	Description : 감성엽서
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/sens/image_commentcls.asp"-->
<%
dim page
page = request("page")
if page="" then page=1

dim oitem
set oitem = new CItemImage
oitem.FCurrPage=page
oitem.FPageSize=20
oitem.GetItemImageList

dim i
%>
<script language="javascript">
// 신규등록
function fnNew(){
	document.location.href="image_comment_edit.asp?mode=add&menupos=<%= menupos %>";
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">	
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="신규등록" onclick="fnNew();">
		</td>
		<td align="right">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oitem.fresultcount >0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oitem.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oitem.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>
			번호
		</td>
		<td>
			아이콘
		</td>
		<td>
			메인표시일
		</td>
		<td>
			등록일
		</td>
		<td>
			사용여부
		</td>
    </tr>
	<% for i=0 to oitem.FResultcount -1 %>
	    <tr align="center" bgcolor="#FFFFFF">
			<td>
				<%= oitem.FItemList(i).Fidx %>
			</td>
			<td>
				<a href="image_comment_edit.asp?mode=edit&reviewid=<%= oitem.FItemList(i).Fidx %>&menupos=<%= menupos %>"><img src="<%= oitem.FItemList(i).FIconUrl %>" width =40 height=40 border="0"></a>
			</td>
			<td>
				<%= FormatDateTime(oitem.FItemList(i).Fviewdate,2) %>
			</td>
			<td>
				<%= FormatDateTime(oitem.FItemList(i).FRegDate,2) %>
			</td>
			<td>
				<% if oitem.FItemList(i).FIsusing = "Y" then %>Y<% else %><font color="red">N</font><% end if %>
			</td>
		</tr>
	<% next %>

	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
	
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if oitem.HasPreScroll then %>
				<a href="?page=<%= oitem.StarScrollPage-1 %>&menupos=<%= menupos %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oitem.StarScrollPage to oitem.FScrollCount + oitem.StarScrollPage - 1 %>
				<% if i>oitem.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if oitem.HasNextScroll then %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->