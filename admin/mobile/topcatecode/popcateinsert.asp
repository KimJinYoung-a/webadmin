<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/mobile/topcatecodeCls.asp" -->
<%
'###############################################
' PageName : popcateinsert.asp
' Discription : ����� ����Ʈ GNB���� �ڵ����
' History : 2015-09-15 ����ȭ
'###############################################

dim linktype, fixtype
dim gcode, page

gcode = request("gcode")
page = request("page")

if gcode="" then gcode=0
if page="" then page=1

dim ognbcode,ognbcodeList

set ognbcode = new GNBcode
ognbcode.FRectgnbcode = gcode
ognbcode.GetOneContentsCode

set ognbcodeList = new GNBcode
ognbcodeList.FPageSize=20
ognbcodeList.FCurrPage= page
ognbcodeList.GetgnbcodeList

dim i
%>
<script language='javascript'>
function Savegnbcode(frm){
    if (frm.gcode.value.length<1){
        alert('GNB �ڵ� ���� �Է��ϼ���.');
        frm.gcode.focus();
        return;
    }
    
    if (frm.gcode.value*1<1){
        alert('GNB �ڵ� ���� 1 �̻��Դϴ�.');
        frm.gcode.focus();
        return;
    }
    
    if (frm.gnbname.value.length<1){
        alert('GNB �̸��� �Է��ϼ���.');
        frm.gnbname.focus();
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

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmgnbcode" method="post" action="popcateproc.asp" >
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">GNB �ڵ�</td>
    <td>
        <% if ognbcode.FOneItem.Fgnbcode<>"" then %>
        <%= ognbcode.FOneItem.Fgnbcode %>
        <input type="hidden" name="gcode" value="<%= ognbcode.FOneItem.Fgnbcode %>" >
        <% else %>
        <input type="text" name="gcode" value="<%= ognbcode.FOneItem.Fgnbcode %>" maxlength="3" size="5">
        (����)
        <% end if %>
            
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">GNB �̸�</td>
    <td>
        <input type="text" name="gnbname" value="<%= ognbcode.FOneItem.Fgnbname %>" maxlength="32" size="64">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��뿩��</td>
    <td>
        <% if ognbcode.FOneItem.Fisusing="N" then %>
        <input type="radio" name="isusing" value="Y">�����
        <input type="radio" name="isusing" value="N" checked >������
        <% else %>
        <input type="radio" name="isusing" value="Y" checked >�����
        <input type="radio" name="isusing" value="N">������
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" �� �� " onClick="Savegnbcode(frmgnbcode);"></td>
</tr>
</form>
</table>
<%
set ognbcode = Nothing
%>
<br>

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="right"><a href="?gcode="><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr bgcolor="#DDDDFF">
    <td width="100">GNB �ڵ�</td>
    <td width="100">GNB �̸�</td>

    <td width="60">��뿩��</td>
</tr>
<% for i=0 to ognbcodeList.FResultCount-1 %>
<% if (CStr(ognbcodeList.FItemList(i).Fgnbcode)=gcode) then %>
<tr bgcolor="#9999CC">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td ><%= ognbcodeList.FItemList(i).Fgnbcode %></td>
    <td ><a href="?gcode=<%= ognbcodeList.FItemList(i).Fgnbcode %>&page=<%= page %>"><%= ognbcodeList.FItemList(i).Fgnbname %></a></td>
    <td ><%= ognbcodeList.FItemList(i).Fisusing %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="center">
    <% if ognbcodeList.HasPreScroll then %>
		<a href="?page=<%= ognbcodeList.StarScrollPage-1 %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ognbcodeList.StarScrollPage to ognbcodeList.FScrollCount + ognbcodeList.StarScrollPage - 1 %>
		<% if i>ognbcodeList.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ognbcodeList.HasNextScroll then %>
		<a href="?page=<%= i %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>

<%
set ognbcodeList = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->