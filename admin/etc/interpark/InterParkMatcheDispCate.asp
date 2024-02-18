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
        alert('검색어를 입력하세요.');
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
        alert('전시 카테고리를 입력하세요.');
        frm.interparkdispcategory.focus();
        return;
    }
    
    if (frm.SupplyCtrtSeq.value.length<1){
        alert('공급 계약 코드를 입력하세요. == 2');
        frm.SupplyCtrtSeq.focus();
        return;
    }
 
    if (frm.interparkstorecategory.value.length<1){
        //alert('브랜드 카테고리를 입력하세요.');
        //frm.interparkstorecategory.focus();
        //return;
    }
   
    
    if (confirm('저장 하시겠습니까?')){
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
			<td width="100" bgcolor="#F3F3FF">카테고리1</td> 
			<td><%= oInterParkitem.FOneItem.Fnmlarge %></td> 
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">카테고리2</td> 
			<td><%= oInterParkitem.FOneItem.Fnmmid %></td> 
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">카테고리3</td> 
			<td><%= oInterParkitem.FOneItem.FnmSmall %></td> 
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">iPark 전시1</td> 
			<td>
				<input type="text" class="text" name="interparkdispcategory" value="<%= oInterParkitem.FOneItem.Finterparkdispcategory %>" size="32" maxlength="32">
				<input type="text" class="text_ro" name="interparkdispcategoryText" value="<%= oInterParkitem.FOneItem.FinterparkdispcategoryText %>" size="50" >
			</td> 
		</tr>
		<tr bgcolor="#FFFFFF">
			<td colspan="2" height="100"></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">공급계약코드</td> 
			<td>
				<input type="text" class="text" name="SupplyCtrtSeq" value="<%= oInterParkitem.FOneItem.FSupplyCtrtSeq %>" size="1" maxlength="1">
				<input type="text" class="text" name="SupplyCtrtSeqName" value="<%= oInterParkitem.FOneItem.getSupplyCtrtSeqName %>" size="10" maxlength="32">
			</td> 
		</tr>
		<% If (FALSE) Then %> <!-- 브랜드 전시 사용 안함. -->
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="#F3F3FF">iPark 브랜드전시1</td> 
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
			<td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SvCode(frmSvr);"></td>
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
			<input type="text" name="sRect" value="" onKeyPress="if (event.keyCode == 13) searchCate(frmDispSearch);" ><input type="button" class="button" value="검색" onClick="searchCate(frmDispSearch);">  
		</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td><iframe name="iFrameDispCate" id="iFrameDispCate" src="/admin/etc/iframeInterParkDispcateSelect.asp" width="600" height="180" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe></td>
		</tr>
		</form>
		<!-- 브랜드 전시 사용 안함.
		<form name="frmStoreCateSearch" >
		<input type="hidden" name="mode" value="all">
		<tr  bgcolor="#FFFFFF">    
			<td>
				<input type="text" name="sRect" value="" onKeyPress="if (event.keyCode == 13) searchStoreCate(frmStoreCateSearch);" ><input type="button" class="button" value="검색" onClick="searchStoreCate(frmStoreCateSearch);">  
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