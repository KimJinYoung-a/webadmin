<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/myqnaCls.asp"-->
<%
'####################################################
' Description :  1:1 ������ ����Ʈ
' History : 2016.07.28 ������ ����
'####################################################
%>
<%
Dim page, i, research, isanswer, searchDiv, searchKey, searchString
Dim oMyqna
page			= RequestCheckvar(request("page"),10)
research		= RequestCheckvar(request("research"),10)
isanswer		= RequestCheckvar(request("isanswer"),1)
searchDiv		= RequestCheckvar(request("searchDiv"),2)
searchKey		= RequestCheckvar(request("searchKey"),16)
searchString	= RequestCheckvar(request("searchString"),128)

If page = "" Then page = 1
If (research = "") Then
	isanswer	= "N"
	searchKey	= "title"
End If

Set oMyqna = new CQna
	oMyqna.FCurrPage			= page
	oMyqna.FPageSize			= 18
	oMyqna.FRectisanswer		= isanswer
	oMyqna.FRectsearchDiv		= searchDiv
	oMyqna.FRectsearchKey		= searchKey
	oMyqna.FRectsearchString	= searchString
	oMyqna.getMyqnaList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function goView(vidx, vgridx){
	location.href='/academy/board/myqnaView.asp?menupos=<%=menupos%>&idx='+vidx+'&gridx='+vgridx;	
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		���� :
		<select name="isanswer" class="select" onChange="">
			<option value="" >��ü</option>
			<option value="Y" <%= chkiif(isanswer = "Y", "selected", "") %> >�Ϸ�</option>
			<option value="N" <%= chkiif(isanswer = "N", "selected", "") %> >���</option>
		</select>
		&nbsp;
		���Ǻо� : 
		<select name="searchDiv" class="select" onChange="">
			<option value="">����</option>
			<option value="1" <%= chkiif(searchDiv="1", "selected", "") %> >��ǰ(��ǰ) �ֹ�/����</option>
			<option value="2" <%= chkiif(searchDiv="2", "selected", "") %> >�ֹ� ���/��ǰ/��ȯ</option>
			<option value="3" <%= chkiif(searchDiv="3", "selected", "") %> >��ǰ ��� ���� ����</option>
			<option value="4" <%= chkiif(searchDiv="4", "selected", "") %> >������û/���� ����</option>
			<option value="5" <%= chkiif(searchDiv="5", "selected", "") %> >���� ���</option>
			<option value="6" <%= chkiif(searchDiv="6", "selected", "") %> >�������� ���� ����</option>
			<option value="7" <%= chkiif(searchDiv="7", "selected", "") %> >�̺�Ʈ/����/���ϸ��� ����</option>
			<option value="8" <%= chkiif(searchDiv="8", "selected", "") %> >ȸ��Ż��/�簡��</option>
			<option value="9" <%= chkiif(searchDiv="9", "selected", "") %> >��Ÿ ����</option>
		</select>
	</td>
	<td align="right">
		�˻� : 
		<select name="searchKey" class="select">
			<option value="idx" <%= chkiif(searchKey = "idx", "selected", "") %> >��ȣ</option>
			<option value="title" <%= chkiif(searchKey = "title", "selected", "") %> >����</option>
			<option value="comment" <%= chkiif(searchKey = "comment", "selected", "") %> >����</option>
			<option value="titlecomment" <%= chkiif(searchKey = "titlecomment", "selected", "") %> >����+����</option>
			<option value="regid" <%= chkiif(searchKey = "regid", "selected", "") %> >�����ID</option>
		</select>
		<input type="text" class="text" name="searchString" size="20" value="<%=searchString%>">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td align="left" colspan="5">
		�Ǽ� : <b><%= FormatNumber(oMyqna.FTotalCount,0) %>��</b>
	</td>
	<td align="right">
		Page : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oMyqna.FTotalPage,0) %> </b>
	</td>
</tr>
<tr height="35" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">��ȣ</td>
	<td width="250">���Ǻо�</td>
	<td width="80">����</td>
	<td>����</td>
	<td width="140">�����</td>
	<td width="140">�����</td>
</tr>
<% For i=0 to oMyqna.FResultCount - 1 %>
<tr height="30" style="cursor:pointer;" align="center" bgcolor='#FFFFFF'" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff'; onclick="goView('<%= oMyqna.FItemList(i).FIdx %>','<%= oMyqna.FItemList(i).FReply_group_idx %>')">
	<td align="center"><%= oMyqna.FItemList(i).FIdx %></td>
	<td align="center"><%= oMyqna.FItemList(i).getQnaGubunName %></td>
	<td align="center"><%= oMyqna.FItemList(i).getAnswerName %></td>
	<td align="left"><%= oMyqna.FItemList(i).FTitle %></td>
	<td align="center"><%= oMyqna.FItemList(i).FUserid %></td>
	<td align="center"><%= FormatDate(oMyqna.FItemList(i).FRegdate,"0000.00.00") %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oMyqna.HasPreScroll then %>
		<a href="javascript:goPage('<%= oMyqna.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oMyqna.StartScrollPage to oMyqna.FScrollCount + oMyqna.StartScrollPage - 1 %>
    		<% if i>oMyqna.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oMyqna.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<% Set oMyqna = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->