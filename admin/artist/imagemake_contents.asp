<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��Ƽ��Ʈ ������ �̹����� ���� ���  
' History : 2012.03.26 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/artist/artist_class.asp"-->
<%
dim idx, reload , ix, gubun
	idx = request("idx")
	gubun = request("gubun")
	reload = request("reload")
	if idx="" then idx=0

if reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
end if

dim oMainContents
	set oMainContents = new cposcode_list
	oMainContents.FRectIdx = idx
	oMainContents.FGubun = gubun
	oMainContents.fcontents_oneitem
%>
<script language='javascript'>
function SaveMainContents(frm){
    if (frm.image_order.value.length<1){
        alert('�̹��� �켱������ �Է� �ϼ���.');
        frm.image_order.focus();
        return;
    }
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}
function SaveMainContents2(frm){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="<%=staticimgurl%>/linkweb/artist/image_proc2.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="gubun" value="<%=gubun%>">
<input type="hidden" name="ckUserId" value="<%=session("ssBctId")%>">
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">Idx :</td>
	    <td><%= oMainContents.FOneItem.Fidx %><input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>"></td>
	</tr>
<% If gubun = 1 Then %>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">�̹������Ŀ켱���� :</td>
	    <td>
			<select name="image_order">
				<option>����</option>
				<% For ix = 1 to 3 %>
					<option value="<%=ix%>" <% if cint(oMainContents.FOneItem.fimage_order) = cint(ix) then response.write " selected"%>><%= ix %></option>				
				<% Next %>
			</select>�Ǽ��� ����� ���ڰ� ������� �켱����
	    </td>
	</tr>
<% End If %>
	<tr bgcolor="#FFFFFF">
		<td width="150" align="center"><%=chkIIF(gubun="1","�Ѹ� ��� �̹��� : ","�̹��� : ")%></td>
		<td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
		<% if oMainContents.FOneItem.Fidx<>"" then %>
		<br><img src="<%=uploadUrl%>/artist/<%= oMainContents.FOneItem.fimagepath %>" border="0"> 
		<br><%=uploadUrl%>/artist/mainBanner/<%= oMainContents.FOneItem.fimagepath %>
		<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ũ�� :</td>
	    <td>
			<input type="text" name="linkpath" value="<%= oMainContents.FOneItem.flinkpath %>" maxlength="128" size="60"><br>(����η� ǥ���� �ּ���  ex: /artist/artist_sub.asp?designerid=0100)
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��뿩��</td>
	    <td>
			<input type="radio" name="isusing" value="Y" checked <%=chkIIF(oMainContents.FOneItem.fisusing = "Y","checked","")%> >Y
			<input type="radio" name="isusing" value="N" <%=chkIIF(oMainContents.FOneItem.fisusing = "N","checked","")%>>N
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align="center" colspan=2>
	    	<% If gubun = 1 Then %>
	    	<input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);" class="button">
	    	<% Else %>
	    	<input type="button" value=" �� �� " onClick="SaveMainContents2(frmcontents);" class="button">
	    	<% End If%>
	    </td>
	</tr>	
</form>
</table>
<%
set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
