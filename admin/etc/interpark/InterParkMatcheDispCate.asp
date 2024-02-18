<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/interpark/interparkcls.asp"-->
<%
Dim cdl, cdm, cdn
cdl = request("cdl")
cdm = request("cdm")
cdn = request("cdn")

Dim oInterParkitem
Set oInterParkitem = new CInterpark
	oInterParkitem.FRectCate_large  = cdl
	oInterParkitem.FRectCate_mid    = cdm
	oInterParkitem.FRectCate_small  = cdn
	oInterParkitem.GetOneInterParkCategoryMaching
%>
<script language='javascript'>
function searchCate(frm){
    if (frm.sRect.value.length<1){
        alert('�˻�� �Է��ϼ���.');
        frm.sRect.focus();
        return;
    }
    frm.action="/admin/etc/iframeInterParkDispcateSelect.asp"
    frm.target = "iFrameDispCate";
    frm.submit();
}

function searchStoreCate(frm){
    frm.action="/admin/etc/iframeInterParkStoreCateSelect.asp"
    frm.target = "iFrameStoreCate";
    frm.submit();
}

function SvCode(frm){
    if (frm.interparkdispcategory.value.length<1){
        alert('���� ī�װ��� �Է��ϼ���.');
        frm.interparkdispcategory.focus();
        return;
    }
    
    if (frm.SupplyCtrtSeq.value.length<1){
        alert('���� ��� �ڵ带 �Է��ϼ���. == 2');
        frm.SupplyCtrtSeq.focus();
        return;
    }
 
    if (frm.interparkstorecategory.value.length<1){
        //alert('�귣�� ī�װ��� �Է��ϼ���.');
        //frm.interparkstorecategory.focus();
        //return;
    }
   
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}
</script>
<table width="100%" border="0" cellspacing="2" cellpadding="2" >
<tr>
	<td valign="top">
		<table width="400" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<form name="frmSvr" method="post" action="/admin/etc/interpark/interparkCate_Process.asp">
		<input type="hidden" name="mode" value="cateedit">
		<input type="hidden" name="tecdl" value="<%= cdl %>">
		<input type="hidden" name="tecdm" value="<%= cdm %>">
		<input type="hidden" name="tecdn" value="<%= cdn %>">
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">ī�װ�1</td> 
			<td><%= oInterParkitem.FOneItem.Fnmlarge %></td> 
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">ī�װ�2</td> 
			<td><%= oInterParkitem.FOneItem.Fnmmid %></td> 
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">ī�װ�3</td> 
			<td><%= oInterParkitem.FOneItem.FnmSmall %></td> 
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">iPark ����1</td> 
			<td>
				<input type="text" class="text" name="interparkdispcategory" value="<%= oInterParkitem.FOneItem.Finterparkdispcategory %>" size="32" maxlength="32">
				<input type="text" class="text_ro" name="interparkdispcategoryText" value="<%= oInterParkitem.FOneItem.FinterparkdispcategoryText %>" size="50" >
			</td> 
		</tr>
		<tr bgcolor="#FFFFFF">
			<td colspan="2" height="100"></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">���ް���ڵ�</td> 
			<td>
				<input type="text" class="text" name="SupplyCtrtSeq" value="<%= oInterParkitem.FOneItem.FSupplyCtrtSeq %>" size="1" maxlength="1">
				<input type="text" class="text" name="SupplyCtrtSeqName" value="<%= oInterParkitem.FOneItem.getSupplyCtrtSeqName %>" size="10" maxlength="32">
			</td> 
		</tr>
		<% If (FALSE) Then %> <!-- �귣�� ���� ��� ����. -->
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">iPark �귣������1</td> 
			<td>
				<input type="text" class="text" name="interparkstorecategory" value="<%= oInterParkitem.FOneItem.Finterparkstorecategory %>" size="32" maxlength="32">
				<input type="text" class="text_ro" name="interparkstorecategoryText" value="<%= oInterParkitem.FOneItem.FinterparkstorecategoryText %>" size="50" >
			</td>
		</tr>
		<% Else %>
			<input type="hidden" name="interparkstorecategory" value="<%= oInterParkitem.FOneItem.Finterparkstorecategory %>" >
			<input type="hidden" name="interparkstorecategoryText" value="<%= oInterParkitem.FOneItem.FinterparkstorecategoryText %>"  >
		<% End If %>
		<tr bgcolor="#FFFFFF">
			<td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SvCode(frmSvr);"></td>
		</tr>
		</form>
		</table>
	</td>
	<td valign="top">
		<table width="400" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<form name="frmDispSearch" >
		<input type="hidden" name="mode" value="all">
		<tr  bgcolor="#FFFFFF">    
			<td>
			<input type="text" name="sRect" value="" onKeyPress="if (event.keyCode == 13) searchCate(frmDispSearch);" ><input type="button" class="button" value="�˻�" onClick="searchCate(frmDispSearch);">  
		</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td><iframe name="iFrameDispCate" id="iFrameDispCate" src="/admin/etc/iframeInterParkDispcateSelect.asp" width="600" height="180" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe></td>
		</tr>
		</form>
		<!-- �귣�� ���� ��� ����.
		<form name="frmStoreCateSearch" >
		<input type="hidden" name="mode" value="all">
		<tr  bgcolor="#FFFFFF">    
			<td>
				<input type="text" name="sRect" value="" onKeyPress="if (event.keyCode == 13) searchStoreCate(frmStoreCateSearch);" ><input type="button" class="button" value="�˻�" onClick="searchStoreCate(frmStoreCateSearch);">  
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td><iframe name="iFrameStoreCate" id="iFrameStoreCate" src="/admin/etc/iframeInterParkStoreCateSelect.asp" width="600" height="180" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe></td>
		</tr>
		</form>
		-->
		</table>
	</td>
</tr>
</table>
<% SET oInterParkitem = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->