<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2007.11.09 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/othermall/othermall_main_contents_managecls.asp" -->

<%
dim linktype, fixtype
dim poscode, page

poscode = request("poscode")
page = request("page")

if poscode="" then poscode=0
if page="" then page=1

dim oposcode,oposcodeList

set oposcode = new CMainContentsCode
	oposcode.FRectPosCode = poscode
	oposcode.GetOneContentsCode

set oposcodeList = new CMainContentsCode
	oposcodeList.FPageSize=10
	oposcodeList.FCurrPage= page
	oposcodeList.GetposcodeList

dim i
%>
<script language='javascript'>
function SavePosCode(frm){
    if (frm.poscode.value.length<1){
        alert('���� �ڵ� ���� �Է��ϼ���.');
        frm.poscode.focus();
        return;
    }
    
    if (frm.poscode.value*1<1){
        alert('���� �ڵ� ���� 1 �̻��Դϴ�.');
        frm.poscode.focus();
        return;
    }
    
    if (frm.posname.value.length<1){
        alert('���и��� �Է��ϼ���.');
        frm.posname.focus();
        return;
    }
    
    if (frm.posVarname.value.length<1){
        alert('��������  �Է��ϼ���.');
        frm.posVarname.focus();
        return;
    }
    
    if (frm.linktype.value.length<1){
        alert('��ũ������ �����ϼ���.');
        frm.linktype.focus();
        return;
    }
    
    if (frm.imagewidth.value.length<1){
        alert('�̹��� ������W�� �Է��ϼ���.');
        frm.imagewidth.focus();
        return;
    }
    
    if (frm.imageheight.value.length<1){
        alert('�̹��� ������H�� �Է��ϼ���.');
        frm.imageheight.focus();
        return;
    }

    if (frm.useSet.value.length<1){
        alert('�̹��� ��밳���� �Է��ϼ���.');
        frm.useSet.focus();
        return;
    }
    
    if (frm.fixtype.value.length<1){
        alert('�ݿ��ֱ⸦ �����ϼ���.');
        frm.fixtype.focus();
        return;
    }
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    
}

function ChangeLinktype(){
    // Do nothing
}
</script>

<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>�ܺθ� �ڵ� �Է�</strong></font>
			</td>		
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
</table>
<!--ǥ ��峡-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frmposcode" method="post" action="othermall_do_mainPosCode.asp" >
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�����ڵ�</td>
    <td>
        <% if oposcode.FOneItem.Fposcode<>"" then %>
        <%= oposcode.FOneItem.Fposcode %>
        <input type="hidden" name="poscode" value="<%= oposcode.FOneItem.Fposcode %>" >
        <% else %>
        <input type="text" name="poscode" value="<%= oposcode.FOneItem.Fposcode %>" maxlength="7" size="5">
        (����)
        <% end if %>         
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">���и�</td>
    <td>
        <input type="text" name="posname" value="<%= oposcode.FOneItem.Fposname %>" maxlength="32" size="64">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">������</td>
    <td>
        <input type="text" name="posVarname" value="<%= oposcode.FOneItem.FposVarname %>" maxlength="32" size="20">
        <br>
        (����/ ���������� ��� : ���� ����, Ư������ ����)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�̹��� �ʺ�</td>
    <td>
        <input type="text" name="imagewidth" value="<%= oposcode.FOneItem.Fimagewidth %>" maxlength="16" size="8">
        (�̹��� Width Size ����)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�̹��� ����</td>
    <td>
        <input type="text" name="imageheight" value="<%= oposcode.FOneItem.Fimageheight %>" maxlength="16" size="8">
        (�̹��� Height Size ���� : 0 �ΰ�� height ���� ����)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�̹��� ��밳��</td>
    <td>
        <input type="text" name="useSet" value="<% if oposcode.FOneItem.FuseSet="" then Response.Write "1":Else Response.Write oposcode.FOneItem.FuseSet:End if %>" size="5">
        (�÷����� ���� ����, �Ϲ��� 1�� ����)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��ũ����</td>
    <td>
        
        <% call DrawLinktypeCombo ("linktype", oposcode.FOneItem.Flinktype, "") %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">���뱸��(�ݿ��ֱ�)</td>
    <td>
        <% call DrawFixTypeCombo ("fixtype", oposcode.FOneItem.Ffixtype, "") %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��뿩��</td>
    <td>
        <% if oposcode.FOneItem.Fisusing="N" then %>
        <input type="radio" name="isusing" value="Y">�����
        <input type="radio" name="isusing" value="N" checked >������
        <% else %>
        <input type="radio" name="isusing" value="Y" checked >�����
        <input type="radio" name="isusing" value="N">������
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SavePosCode(frmposcode);"></td>
</tr>
</form>
</table>
<%
set oposcode = Nothing
%>
<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="right"><a href="?poscode="><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr bgcolor="#DDDDFF">
    <td width="100">code</td>
    <td width="100">���и�</td>
    <td width="100">������</td>
    <td width="100">��ũ����</td>
    <td width="100">�ݿ��ֱ�</td>
    <td width="60">��뿩��</td>
</tr>
<% for i=0 to oposcodeList.FResultCount-1 %>
<% if (CStr(oposcodeList.FItemList(i).FposCode)=poscode) then %>
<tr bgcolor="#9999CC">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td ><%= oposcodeList.FItemList(i).FposCode %></td>
    <td ><a href="?poscode=<%= oposcodeList.FItemList(i).FposCode %>&page=<%= page %>"><%= oposcodeList.FItemList(i).FposName %></a></td>
    <td ><%= oposcodeList.FItemList(i).FposVarName %></td>
    <td ><%= oposcodeList.FItemList(i).getlinktypeName %></td>
    <td ><%= oposcodeList.FItemList(i).getfixtypeName %></td>
    <td ><%= oposcodeList.FItemList(i).Fisusing %></td>
</tr>
<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
       <td valign="bottom" align="center">
		    <% if oposcodeList.HasPreScroll then %>
				<a href="?page=<%= oposcodeList.StarScrollPage-1 %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + oposcodeList.StarScrollPage to oposcodeList.FScrollCount + oposcodeList.StarScrollPage - 1 %>
				<% if i>oposcodeList.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>">[<%= i %>]</a>
				<% end if %>
			<% next %>
		
			<% if oposcodeList.HasNextScroll then %>
				<a href="?page=<%= i %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<%
set oposcodeList = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->