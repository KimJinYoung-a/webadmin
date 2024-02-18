<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : MY�˸�
' Hieditor : 2009.04.17 ������ ����
'			 2016.07.19 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/my10x10/myAlarmCls.asp" -->
<%
dim research, useYN, page, i
	research= request("research")
	useYN = request("useYN")

if ((research = "") and (useYN = "")) then
	useYN = "Y"
end if

if page="" then page=1

dim oCMyAlarm
set oCMyAlarm = new CMyAlarm
	oCMyAlarm.FPageSize = 20
	oCMyAlarm.FCurrPage = page
	oCMyAlarm.FRectUseYN = useYN
	oCMyAlarm.GetMyAlarmByLevel

%>
<script type="text/javascript">

function NextPage(page) {
    frm.page.value = page;
    frm.submit();
}

function AddNewMyAlarm() {
    var popwin = window.open("popMyAlarmEdit.asp?idx=0","AddNewMyAlarm","width=600,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function ModiMyAlarm(idx) {
    var popwin = window.open("popMyAlarmEdit.asp?idx=" + idx,"ModiMyAlarm","width=600,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
}

</script>

<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td height="30" align="left">
	    ��뿩�� :
		<select class="select" name="useYN">
			<option value="">��ü</option>
			<option value="Y" <% if useYN = "Y" then response.write "selected" %> >�����</option>
			<option value="N" <% if useYN = "N" then response.write "selected" %> >������</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button" value=" �˻� " onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>

<input type="button" class="button" value="�űԵ��" onClick="AddNewMyAlarm()">

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oCMyAlarm.FtotalCount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCMyAlarm.FtotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td height="25" width="40">IDX</td>
    <td width="80">�˸���¥</td>
    <td width="200">����</td>
    <td width="200">������</td>
    <td width="250">����</td>
    <td width="100">Ÿ�ٵ��</td>
    <td>Ÿ��URL</td>
    <td width="40">����<br>����</td>
    <td width="40">���<br>����</td>
    <td width="80">�����</td>
	<td width="80">��������</td>
    <td></td>
</tr>
<%
	for i = 0 to oCMyAlarm.FResultCount - 1
%>
<% if (oCMyAlarm.FItemList(i).FuseYN = "N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
	<td align="center" height="25"><%= oCMyAlarm.FItemList(i).FlevelAlarmIdx %></td>
	<td align="center"><%= oCMyAlarm.FItemList(i).Fyyyymmdd %></td>
	<td align="left"><a href="javascript:ModiMyAlarm(<%= oCMyAlarm.FItemList(i).FlevelAlarmIdx %>)"><%= oCMyAlarm.FItemList(i).Ftitle %></a></td>
	<td align="left"><%= oCMyAlarm.FItemList(i).Fsubtitle %></td>
	<td align="left"><%= oCMyAlarm.FItemList(i).Fcontents %></td>
	<td align="center">
		<% if oCMyAlarm.FItemList(i).fUserLevel="100" then %>
			���ȸ�� ��ü
		<% else %>
			<%= getUserLevelStr(oCMyAlarm.FItemList(i).fUserLevel) %>
		<% end if %>
	</td>
	<td align="left"><%= oCMyAlarm.FItemList(i).FwwwTargetURL %></td>
	<td align="center"><%= oCMyAlarm.FItemList(i).FopenYN %></td>
	<td align="center"><%= oCMyAlarm.FItemList(i).FuseYN %></td>
	<td align="center"><%= oCMyAlarm.FItemList(i).Freguserid %></td>
	<td align="center"><%= Left(oCMyAlarm.FItemList(i).Flastupdate, 10) %></td>
	<td></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="13" align="center" height="30">
    <% if oCMyAlarm.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCMyAlarm.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oCMyAlarm.StarScrollPage to oCMyAlarm.FScrollCount + oCMyAlarm.StarScrollPage - 1 %>
		<% if i>oCMyAlarm.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oCMyAlarm.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>

<form name="frmAct" method="post">
</form>

<%
set oCMyAlarm = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
