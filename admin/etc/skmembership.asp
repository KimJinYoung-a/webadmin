<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/skmembershippointcls.asp"-->
<%
dim skuserid, page
skuserid = request("skuserid")
page = request("page")
if page="" then page=1

dim omembership
set omembership = new CSktSentence
omembership.FPageSize=200
omembership.FCurrpage=page
omembership.FRectOnlySended = "on"
omembership.FRectSkUserid = skuserid

omembership.getCheckSentenceList

dim i
%>
<script language='javascript'>
function CancelMember(iid){
	frmcancel.idx.value=iid;
	frmcancel.password.value=frm.pwd.value;
	if (confirm('취소하시겠습니까?')){
		frmcancel.submit();
	}
}

function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<tr>
		<td class="a" width="600">
		SK 아이디 <input type=text name="skuserid" value="<%= skuserid %>" size=10>

		&nbsp;&nbsp;&nbsp;
		PWD <input type=password name="pwd" value="" size=10>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit()"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td>idx</td>
	<td>주문번호</td>
	<td>결제수단</td>
	<td>skuserid</td>
	<td>userid</td>
	<td>regdate</td>
	<td>apprcode</td>
	<td>resultcode</td>
	<td>discountsum</td>
	<td>취소</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="10" >배송전취소</td>
</tr>
<% for i=0 to omembership.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
	<td><%= omembership.FItemList(i).Fidx %></td>
	<td><%= omembership.FItemList(i).Forderserial %></td>
	<td><%= omembership.FItemList(i).getAccountDivName %></td>

	<td><%= omembership.FItemList(i).Fskuserid %></td>
	<td><%= omembership.FItemList(i).Fuserid %></td>
	<td><%= omembership.FItemList(i).Fregdate %></td>
	<td><%= omembership.FItemList(i).Fapprcode %></td>
	<td><%= omembership.FItemList(i).Fresultcode %></td>
	<td><%= omembership.FItemList(i).Fdiscountsum %></td>
	<td align=center>
	<% if IsNULL(omembership.FItemList(i).FCancelIdx) then %>
	<a href="javascript:CancelMember('<%= omembership.FItemList(i).Fidx %>');">취소</a>
	<% else %>
	기취소건
	<% end if %>
	</td>
</tr>
<% next %>
<%
omembership.getCheckSentenceList2

%>
<tr bgcolor="#FFFFFF">
	<td colspan="10" >반품</td>
</tr>
<% for i=0 to omembership.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
	<td><%= omembership.FItemList(i).Fidx %></td>
	<td><%= omembership.FItemList(i).Forderserial %></td>
	<td><%= omembership.FItemList(i).getAccountDivName %></td>

	<td><%= omembership.FItemList(i).Fskuserid %></td>
	<td><%= omembership.FItemList(i).Fuserid %></td>
	<td><%= omembership.FItemList(i).Fregdate %></td>
	<td><%= omembership.FItemList(i).Fapprcode %></td>
	<td><%= omembership.FItemList(i).Fresultcode %></td>
	<td><%= omembership.FItemList(i).Fdiscountsum %></td>
	<td align=center>
	<% if IsNULL(omembership.FItemList(i).FCancelIdx) then %>
	<a href="javascript:CancelMember('<%= omembership.FItemList(i).Fidx %>');">취소</a>
	<% else %>
	기취소건
	<% end if %>
	</td>
</tr>
<% next %>
<!--
<tr>
	<td colspan="10" align="center" bgcolor="#FFFFFF">
		<% if omembership.HasPreScroll then %>
		<a href="javascript:NextPage('<%= omembership.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + omembership.StarScrollPage to omembership.FScrollCount + omembership.StarScrollPage - 1 %>
			<% if i>omembership.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if omembership.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
-->
</table>
<form name=frmcancel method=post action="docancelmembership.asp">
<input type=hidden name=idx value="">
<input type=hidden name=password value="">
<input type=hidden name=mode value="cancel">
</form>
<%
set omembership = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->