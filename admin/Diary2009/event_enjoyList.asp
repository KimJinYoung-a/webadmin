<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/Classes/DiaryEnjoyCls.asp"-->
<%
'###############################################
' PageName : event_enjoyList.asp
' Discription : �۰����� �׷��� ���
' History : 2009.09.30 ������ : ����
'###############################################

dim page, i, lp
dim makerid, isusing

page = request("page")
if page = "" then page=1
makerid = request("makerid")
isusing = request("isusing")
if isusing = "" then isusing="Y"

dim oEnjoy
set oEnjoy = New CEnjoy
oEnjoy.FCurrPage = page
oEnjoy.FPageSize=20
oEnjoy.FRectMaker = makerid
oEnjoy.FRectUsing = isusing
oEnjoy.GetDiaryEnjoyList

%>
<script language='javascript'>
<!--
// ������ �̵�
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="event_enjoyList.asp";
	document.refreshFrm.submit();
}

//-->
</script>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="event_enjoyList.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		�귣��:
	    <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20" >
	    <input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID(this.form.name,'makerid');" >
		/ ��뿩��
		<select name="isusing">
			<option value="Y">���</option>
			<option value="N">����</option>
		</select>
		<script language="javascript">
		refreshFrm.isusing.value="<%=isusing%>";
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right"><input type="button" value="�۰� �߰�" onclick="self.location='event_enjoyWrite.asp?mode=add&menupos=<%= menupos %>'" class="button"></td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="6">
		�˻���� : <b><%=oEnjoy.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oEnjoy.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ȣ</td>
	<td>�귣��</td>
	<td>�̹���</td>
	<td>����</td>
	<td>�ڸ�Ʈ</td>
	<td>�����</td>
</tr>
<%	if oEnjoy.FResultCount < 1 then %>
<tr>
	<td colspan="6" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �������� �����ϴ�.</td>
</tr>
<%
	else
		for i=0 to oEnjoy.FResultCount-1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><a href="event_enjoyWrite.asp?mode=edit&menupos=<%= menupos %>&denjSn=<%= oEnjoy.FItemList(i).FdenjSn %>"><%= oEnjoy.FItemList(i).FdenjSn %></a></td>
	<td align="center"><%= oEnjoy.FItemList(i).Fbrandname %></td>
	<td align="center"><a href="event_enjoyWrite.asp?mode=edit&menupos=<%= menupos %>&denjSn=<%= oEnjoy.FItemList(i).FdenjSn %>"><img src="<%= webImgUrl & "/diary_collection/enjoy/" & oEnjoy.FItemList(i).FsmallImage %>" width="100" border="0"></a></td>
	<td align="center"><a href="event_enjoyWrite.asp?mode=edit&menupos=<%= menupos %>&denjSn=<%= oEnjoy.FItemList(i).FdenjSn %>"><%= oEnjoy.FItemList(i).Fsubject %></a></td>
	<td align="center"><%= oEnjoy.FItemList(i).FcmtCnt %>��</td>
	<td align="center"><%= left(oEnjoy.FItemList(i).Fregdate,10) %></td>
</tr>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<tr bgcolor="#FFFFFF">
	<td colspan="6" align="center">
	<!-- ������ ���� -->
	<%
		if oEnjoy.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oEnjoy.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oEnjoy.StartScrollPage to oEnjoy.FScrollCount + oEnjoy.StartScrollPage - 1

			if lp>oEnjoy.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oEnjoy.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- ������ �� -->
	</td>
</tr>
</table>
<%
set oEnjoy = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->