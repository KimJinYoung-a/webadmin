<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/optionmanagecls.asp"-->
<%
dim cdl,cdm,cds
cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")


dim mode,ckitem, t_cdl, t_cdm, t_cds

mode = request("mode")
ckitem = request("ckitem")

t_cdl = request("t_cdl")
t_cdm = request("t_cdm")
t_cds = request("t_cds")
''response.write ckitem

dim sqlStr
if mode="addArr" then
	sqlStr = "insert into [db_temp].[dbo].tbl_temp_itemcategory"
	sqlStr = sqlStr + " (itemid,cdlarge,cdmid,cdsmall)"
	sqlStr = sqlStr + " select itemid, '" + t_cdl + "','" + t_cdm + "','" + t_cds + "'"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item"
	sqlStr = sqlStr + " where itemid in (" + ckitem + ")"
	rsget.Open sqlStr, dbget, 1

	response.write "<script>parent.parent.newcate.imatchitem.location.reload();</script>"
end if


dim oCateItem
set oCateItem = new CCatemanager
oCateItem.FPageSize = 300

	oCateItem.GetOrgCateNotMachItemList

dim i
%>
<script language='javascript'>
function SaveArr(){
	var cdl,cdm,cds;
	var catename = parent.newcate.AvailCategory();

	if (catename==""){
		alert('새로운카테고리를 3단계까지 선택하세요.');
		return;
	}
	cdl = catename.substr(0,2);
	cdm = catename.substr(2,2);
	cds = catename.substr(4,2);

	frm.t_cdl.value = cdl;
	frm.t_cdm.value = cdm;
	frm.t_cds.value = cds;

	var passed = false;
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="checkbox")) {
			passed = e.checked;
			if (passed) break;
		}
	}

	if (!passed){
		alert('선택 상품이 없습니다.');
		return;
	}

	if (confirm('선택 상품을 ' + catename + ' 카테고리로 저장하시겠습니까?')){
		frm.submit();
	}
}

function CheckAll(comp){
	var bool = comp.checked;

	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="checkbox")) {
			e.checked = bool;
		}
	}
}
</script>

<table width=300 cellspacing=0 cellpadding=0 class=a border=1>
<tr>
	<td colspan="3" align=right><%= oCateItem.FResultCount %>/<%= oCateItem.FTotalCount %></td>
</tr>
<tr>
	<td width=20><input type="checkbox" name="ckall" onClick="CheckAll(this);"></td>
	<td colspan=2 align=right><input type="button" value="새로운 카테고리로 저장" onclick="SaveArr();"></td>
</tr>
<form name=frm method=post action="">
<input type="hidden" name="mode" value="addArr">
<input type="hidden" name="t_cdl" value="">
<input type="hidden" name="t_cdm" value="">
<input type="hidden" name="t_cds" value="">
<% for i=0 to oCateItem.FResultCount-1 %>
<tr>
	<td><input type="checkbox" name="ckitem" value="<%= oCateItem.FITemList(i).FItemID %>"></td>
	<td width=50><img src="<%= oCateItem.FITemList(i).FImgSmall %>" width=50 height=50></td>
	<td><font color="#888888"><%= oCateItem.FITemList(i).FItemName %></font><br>(<%= oCateItem.FITemList(i).FMakerid %>) <%= oCateItem.FITemList(i).Fitemid %><br>
	<% if oCateItem.FITemList(i).Fdispyn="N" then %>
	<font color="blue">전시x</font>
	<% end if %>
	<% if oCateItem.FITemList(i).Fsellyn="N" then %>
	<font color="red">판매x</font>
	<% end if %>
	<% if oCateItem.FITemList(i).Fisusing="N" then %>
	사용x
	<% end if %>
	</td>
</tr>
<% next %>
</form>
</table>
<%
set oCateItem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->