<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- # include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->

<%
	Dim clist, vPage, iTotCnt, i, vItemID
	vPage = NullFillWith(requestCheckVar(Request("page"),10),1)
	
	Set clist = New cOnlySys
	 	clist.FCurrPage = vPage
	 	clist.FPageSize = 10
		clist.fnAwardNotIncludeItemList
		iTotCnt = clist.ftotalcount
%>

<script>
function searchFrm(p){
	frm.page.value = p;
	frm.submit();
}

function jsSaveItem(){
	if(frm1.itemid.value == ""){
		alert("��ǰ�ڵ带 ��������.");
		frm1.itemid.focus();
		return;
	}
	
	frm1.submit();
}

function jsDelItem(i){
	if(confirm("������ ��ǰ�� ����忡 �ٽ� �߰��Ͻðڽ��ϱ�?") == true) {
		frm1.itemid.value = i;
		frm1.gubun.value = "delete";
		frm1.submit();
	}
}
</script>
<br>
<h2>* (����Ʈ�����, ����, GiftTalk ����Ʈ) ��ǰ �����ϱ�</h1>
<br>
<form name="frm1" action="award_notinclude_item_proc.asp" method="post">
<input type="hidden" name="gubun" value="insert">
<table cellpadding="0" cellspacing="0" border="0" class="a">
<tr>
	<td>
		��ǰ�ڵ� : <textarea name="itemid" cols="30" rows="5"><%=vItemID%></textarea>
		<input type="button" class="button" value="�� ��" onClick="jsSaveItem()">
	</td>
</tr>
<tr>
	<td>�� 1�� �̻��� ��� ��ǥ�� �����Ͽ� �Է�.</td>
</tr>
</table>
</form>
<br>
<form name="frm" action="<%=CurrURL%>" method="get">
<input type="hidden" name="page" value="">
</form>
<strong>- ���ܵ� ��ǰ����Ʈ</strong>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap>ItemID</td>
  	<td nowrap>��ǰ��</td>
  	<td nowrap>�����</td>
  	<td nowrap></td>
</tr>
<%
If clist.FResultCount > 0 Then
	For i=0 To clist.FResultCount-1
%>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%=clist.FItemList(i).Fitemid%></td>
		<td> <img src="<%=webImgUrl%>/image/small/<%=GetImageSubFolderByItemid(clist.FItemList(i).Fitemid)%>/<%=clist.FItemList(i).Fsmallimage%>" width="50" height="50" border="0"> <%=clist.FItemList(i).Fitemname%>
			[<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=clist.FItemList(i).Fitemid%>" target="_blank">��ǰ��ũ</a>]
		</td>
		<td align="center"><%=clist.FItemList(i).Fregdate%></td>
		<td align="center"><input type="button" value="�ٽ� �߰�" onClick="jsDelItem('<%=clist.FItemList(i).Fitemid%>');"></td>
	</tr>
<%
	Next
End If
%>
</table>

<table width="100%" border="0" align="center" class="a">
<tr height="50" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		<% if clist.HasPreScroll then %>
		<a href="javascript:searchFrm('<%= clist.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + clist.StartScrollPage to clist.FScrollCount + clist.StartScrollPage - 1 %>
			<% if i>clist.FTotalpage then Exit for %>
			<% if CStr(vPage)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if clist.HasNextScroll then %>
			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<% Set clist = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->