<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/brandmainCls.asp" -->
<%
Dim idx, reload, ix, makerid
	idx = request("idx")
	reload = request("reload")
	if idx="" then idx=0

If reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
End If

Dim mbrand
Set mbrand = New cBrandMain
	mbrand.FRectIdx = idx
	mbrand.sMainTop3modify

	makerid = mbrand.FOneItem.fmakerid
%>
<script language='javascript'>
function SaveMainContents(frm){
    if (frm.makerid.value.length<1){
        alert('�귣�带 �Է� �ϼ���.');
        frm.makerid.focus();
        return;
    }
    if (frm.image_order.value.length<1){
        alert('�̹��� �켱������ �Է� �ϼ���.');
        frm.image_order.focus();
        return;
    }
    if (frm.linkpath.value.length<1){
        alert('��ũ���� �Է��ϼ���');
        frm.linkpath.focus();
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

//�귣�� ID �˻� �˾�â
function jsSearchBrandIDNew(frmName,compName){
	var compVal = "";
	try{
		compVal = eval("document.all." + frmName + "." + compName).value;
	}catch(e){
		compVal = "";
	}

	var popwin = window.open("popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="doBrandPick.asp" onsubmit="return false;">
<input type="hidden" name="ckUserId" value="<%=session("ssBctId")%>">
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">Idx :</td>
	    <td><%= mbrand.FOneItem.Fidx %><input type="hidden" name="idx" value="<%= mbrand.FOneItem.Fidx %>"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">�귣�� :</td>
	    <td><% NewDrawSelectBoxDesignerwithNameEvent "makerid", makerid %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">�̹������Ŀ켱���� :</td>
	    <td>
	    	<input type="text" size="2" maxlength="2" name="image_order"  class="text" value=<%= chkiif(mbrand.FOneItem.Fidx = "","99",mbrand.FOneItem.FImage_order) %>>
	    	&nbsp;�Ǽ��� ����� ���ڰ� ������� �켱����
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="150" align="center">�̹��� : </td>
		<td>
			<input type="hidden" name="Imagepath" value="<%=mbrand.FOneItem.fimagepath%>">
			<% If mbrand.FOneItem.Fidx <> "" Then %>
				<% if instr(mbrand.FOneItem.FImagepath,"http://")>0 then %>
					<br><img src="<%=mbrand.FOneItem.fimagepath%>" width="200" id="mainimg" border="0">
					<br><span id="imgurl"><%=mbrand.FOneItem.fimagepath%></span>
				<% else %>
					<br><img src="<%=uploadUrl%>/brandstreet/main/<%= mbrand.FOneItem.fimagepath %>" width="200" id="mainimg" border="0">
					<br><span id="imgurl"><%=uploadUrl%>/brandstreet/main/<%= mbrand.FOneItem.fimagepath %></span>
				<% end if %>
			<% else %>
				<br><img src="" width="200" border="0" id="mainimg">
				<br><span id="imgurl"></span>
			<% End If %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ũ�� :</td>
	    <td>
			<input type="text" name="linkpath" value="<%= mbrand.FOneItem.flinkpath %>" maxlength="128" size="60"><br>(����η� ǥ���� �ּ���  ex: /street/street_brand_sub06.asp?makerid=ithinkso)
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��뿩��</td>
	    <td>
			<input type="radio" name="isusing" value="Y" checked <%=chkIIF(mbrand.FOneItem.fisusing = "Y","checked","")%> >Y
			<input type="radio" name="isusing" value="N" <%=chkIIF(mbrand.FOneItem.fisusing = "N","checked","")%>>N
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align="center" colspan=2>
	    	<input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);" class="button">
	    </td>
	</tr>	
</form>
</table>
<%
set mbrand = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->