<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.23 �ѿ�� ����
'	Description : ���ų�����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->


<%
dim linktype, fixtype
dim poscode, page

poscode = request("poscode")
page = request("page")

if poscode="" then poscode=0
if page="" then page=1

dim oposcode,oposcodeList

set oposcode = new organizerCls
	oposcode.FRectPosCode = poscode
	oposcode.fposcode_oneitem

set oposcodeList = new organizerCls
	oposcodeList.FPageSize=10
	oposcodeList.FCurrPage= page
	oposcodeList.fposcode_list

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
    
    
    if (frm.imagetype.value.length<1){
        alert('��ũŸ���� �����ϼ���.');
        frm.imagetype.focus();
        return;
    }
    
    if (frm.imagewidth.value.length<1){
        alert('�̹��� ������ ���θ� �Է��ϼ���.');
        frm.imagewidth.focus();
        return;
    }
    
    if (frm.imageheight.value.length<1){
        alert('�̹��� ������ ���θ� �Է��ϼ���.');
        frm.imageheight.focus();
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
	<tr height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
</table>
<!--ǥ ��峡-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frmposcode" method="post" action="/admin/organizer/imagemake/imagemake_poscode_process.asp" >
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
    <td width="150" bgcolor="#DDDDFF">������̹�����</td>
    <td>
        <input type="text" name="imagecount" value="<%= oposcode.FOneItem.fimagecount %>" maxlength="2" size="2">
        (���ڸ� �Է� �ϼ���)
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
    <td width="150" bgcolor="#DDDDFF">��ũŸ��</td>
    <td>
        <select name="imagetype">
	        <option value="" <% if oposcode.FOneItem.fimagetype = "" then response.write " selected" %>>����</option>
	        <option value="map" <% if oposcode.FOneItem.fimagetype = "map" then response.write " selected" %>>map</option>
	        <option value="link" <% if oposcode.FOneItem.fimagetype = "link" then response.write " selected" %>>link</option>                
	        <option value="flash" <% if oposcode.FOneItem.fimagetype = "flash" then response.write " selected" %>>flash</option>
	        <option value="multi" <% if oposcode.FOneItem.fimagetype = "multi" then response.write " selected" %>>multi</option>
        </select>
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
<% if oposcodeList.FResultCount > 0 then %>
	<tr bgcolor="#DDDDFF">
	    <td width="100">code</td>
	    <td width="100">���и�</td>
	    <td width="100">��ũŸ��</td>
	    <td width="60">Image��</td>
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
		    <td ><%= oposcodeList.FItemList(i).fimagetype %></td>
		    <td ><%= oposcodeList.FItemList(i).fimagecount %></td>
		    <td ><%= oposcodeList.FItemList(i).Fisusing %></td>
		</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td align="center">������ �����ϴ�.</td>
	</tr>	
<% end if %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
       <td valign="bottom" align="center">
		    <% if oposcodeList.HasPreScroll then %>
				<a href="?page=<%= oposcodeList.StartScrollPage-1 %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + oposcodeList.StartScrollPage to oposcodeList.FScrollCount + oposcodeList.StartScrollPage - 1 %>
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
