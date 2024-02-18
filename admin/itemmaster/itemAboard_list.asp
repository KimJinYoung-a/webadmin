<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ؿܻ�ǰ�Ӽ�����
' History : �̻� ����
'			2018.10.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, itemname, makerid, sellyn, usingyn, mwdiv, limityn, overSeaYn, weightYn, itemrackcode,research
dim cdl, cdm, cds, sortDiv, page, limitrealstock, stocktype, i, pojangok, itemManageType, sizeYn, chdeliverOverseas
dim itemdivNotexists
	itemid		= request("itemid")
	itemname	= requestCheckVar(request("itemname"),128)
	makerid		= requestCheckVar(request("makerid"),32)
	sellyn		= requestCheckVar(request("sellyn"),1)
	usingyn		= requestCheckVar(request("usingyn"),1)
	mwdiv		= requestCheckVar(request("mwdiv"),32)
	limityn		= requestCheckVar(request("limityn"),1)
	overSeaYn	= requestCheckVar(request("overSeaYn"),1)
	weightYn	= requestCheckVar(request("weightYn"),1)
	itemrackcode= requestCheckVar(request("itemrackcode"),32)
	sortDiv		= requestCheckVar(request("sortDiv"),32)
	research	=requestCheckVar(Request("research"),1)
	pojangok	=requestCheckVar(Request("pojangok"),1)
	cdl = requestCheckVar(request("cdl"),32)
	cdm = requestCheckVar(request("cdm"),32)
	cds = requestCheckVar(request("cds"),32)
	page = requestCheckVar(request("page"),32)
	limitrealstock = requestCheckVar(request("limitrealstock"),32)
	stocktype = requestCheckVar(request("stocktype"),32)
	itemManageType = requestCheckVar(request("itemManageType"),32)
	sizeYn = requestCheckVar(request("sizeYn"),32)
	chdeliverOverseas = requestCheckVar(request("chdeliverOverseas"),10)
	itemdivNotexists = requestCheckVar(request("itemdivNotexists"),32)

'�⺻��
if chdeliverOverseas="" then chdeliverOverseas="Y"
if (page="") then page=1
if sortDiv="" then sortDiv="new"
if research="" then
	if mwdiv="" then mwdiv="MW"
	if overSeaYn="" then overSeaYn="Y"
	if weightYn="" then weightYn="Y"
	'if pojangok="" then pojangok="Y"
	itemManageType = "I"
end if
if research="" and itemdivNotexists="" then
	itemdivNotexists="on"
end if
if (stocktype = "") then
	stocktype = "sys"
end if

if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim oitem
set oitem = new CItem
	oitem.FPageSize         = 30
	oitem.FCurrPage         = page
	oitem.FRectMakerid      = makerid
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectLimityn      = limityn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectIsOversea	= overSeaYn
	oitem.FRectIsWeight		= weightYn
	oitem.FRectRackcode		= itemrackcode
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectSortDiv		= sortDiv
	oitem.FRectlimitrealstock = limitrealstock
	oitem.FRectStockType = stocktype
	oitem.FRectpojangok = pojangok
	oitem.FRecItemManageType = itemManageType
	oitem.FRectSizeYn = sizeYn

	if itemdivNotexists="on" then
		oitem.frectitemdivNotexists="'08','21'"
	end if

	oitem.GetItemAboardList

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function chgSort(srt){
	document.frm.sortDiv.value= srt;
	document.frm.submit();
}

// �ɼǼ��� -��ü
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//�Ǹż���
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

// �̹�������
function editItemImage(itemid, makerid) {
	var param = "itemid=" + itemid;

	//if(makerid =="ithinkso"){
		//popwin = window.open('/common/pop_itemimage_ithinkso.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	//}else{
		popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	//}
	popwin.focus();
}

// ��ǰ���� �̹��� ���/����
function popItemContImage(itemid)
{
	var popwin = window.open("/admin/shopmaster/item_imgcontents_write.asp?mode=edit&itemid=" + itemid + "&menupos=423","popitemContImage","width=600 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// �����Ȳ �˾�
function PopItemStock(itemid){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// �⺻���� ����
function editItemBasicInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// �ǸŰ� �� ���ް� ����
function editItemPriceInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('pop_ItemPriceInfo.asp?' + param ,'editItemPrice','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopItemWeightEdit(iitemid){
	var popwin = window.open('/warehouse/pop_ItemWeightEdit.asp?itembarcode=' + iitemid + '&menupos=<%=menupos%>','itemWeightEdit','width=800,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function jsSubmit(frm) {
	/*
	if ((frm.itemManageType.value == 'O') && (frm.makerid.value == '')) {
		alert('�ɼǺ� �˻��� �귣�带 �����ؾ߸� �˻������մϴ�.');
		return;
	}
	*/

	frm.submit();
}

function downloadexcel() {
	alert('200000�Ǳ��� �ٿ�ε� ����. �ε��� ��ٷ� �ּ���.');
	frm.action='/common/item/itemAboard_exceldownload.asp';
	frm.target='view';
	frm.submit();
	frm.action='';
	frm.target='';
	return false;
}

function regdeliverOverseas(){
	if (frm.chdeliverOverseas.value==""){
		alert("�ϰ������Ͻ� �ؿܹ�ۿ��θ� ������ �ּ���.");
		return;
	}
	frmlist.chdeliverOverseas.value=frm.chdeliverOverseas.value

    if ($('input[name="check"]:checked').length == 0) {
        alert('�ϰ������Ͻ� ��ǰ�� ������ �ּ���.');
        return;
    }

	frmlist.action="/warehouse/itemWeight_process.asp";
	frmlist.target="view";
	frmlist.submit();
}

function toggleChecked(status) {
    $('[name="check"]').each(function () {
        $(this).prop("checked", status);
    });
}

$(document).ready(function () {
    var checkAllBox = $("#ckall");

    checkAllBox.click(function () {
        var status = checkAllBox.prop('checked');
        toggleChecked(status);
    });
});

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<input type="hidden" name="sortDiv" value="<%=sortDiv%>">
<input type="hidden" name="research" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣�� : <%	drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br><br>
		���ڵ� :
		<input type="text" class="text" name="itemrackcode" value="<%= itemrackcode %>" size="12" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		��ǰ�ڵ� :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)
		&nbsp;
		��ǰ�� :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="jsSubmit(document.frm)">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		�Ǹ� : <% drawSelectBoxSellYN "sellyn", sellyn %>
		&nbsp;
		��� : <% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;
		���� : <% drawSelectBoxLimitYN "limityn", limityn %>
		&nbsp;
		�ŷ����� : <% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;
		�ؿܹ�� : <% drawSelectBoxUsingYN "overSeaYn", overSeaYn %>
		&nbsp;
		���尡�ɿ��� : <% drawSelectBoxUsingYN "pojangok", pojangok %>
		&nbsp;
		��� <select name="stocktype" class="select">
			<option value="sys" <% if (stocktype = "sys") then %>selected<% end if %> >�ý������</option>
			<option value="real" <% if (stocktype = "real") then %>selected<% end if %> >��ȿ���</option>
		</select>
		: <% drawSelectBoxexistsstock "limitrealstock", limitrealstock, "" %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		��Ϲ�� :
		<select name="itemManageType" class="select">
			<option value="I" <% if (itemManageType = "I") then %>selected<% end if %> >��ǰ��</option>
			<option value="O" <% if (itemManageType = "O") then %>selected<% end if %> >�ɼǺ�</option>
		</select>
		&nbsp;
		���Կ��� : <% drawSelectBoxUsingYN "weightYn", weightYn %>
		&nbsp;
		������� : <% drawSelectBoxUsingYN "sizeYn", sizeYn %>
		&nbsp;
		<input type="checkbox" name="itemdivNotexists" value="on" <% if itemdivNotexists="on" then response.write "checked" %> >Ƽ��/Ŭ����/����ǰ����
	</td>
</tr>
</table>

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			�ؿܹ�ۿ��� : <% drawSelectBoxUsingYN "chdeliverOverseas", chdeliverOverseas %>
			<input type="button" onclick="regdeliverOverseas();" value="�ؿܹ�ۿ����ϰ�����" class="button">
		</td>
		<td align="right">
			<input type="button" onclick="downloadexcel(); return false;" value="���� �ٿ�ε�" class="button">&nbsp;&nbsp;
			���Ĺ�� :
			<select name="sort" class="select" onchange="chgSort(this.value)">
				<option value="new" <% if sortDiv="new" then Response.Write "selected" %>>�Ż�ǰ��</option>
				<option value="rack" <% if sortDiv="rack" then Response.Write "selected" %>>���ڵ��</option>
				<option value="weight" <% if sortDiv="weight" then Response.Write "selected" %>>��ǰ���Լ�</option>
			</select>
		</td>
	</tr>
</table>
</form>
<!-- �׼� �� -->

<form name="frmlist" method="post" action="" style="margin:0px;">
<input type="hidden" name="mode" value="chdeliverOverseas">
<input type="hidden" name="chdeliverOverseas" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oitem.FTotalCount%></b>
		&nbsp;
		������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="30"><input type="checkbox" name="ckall" id="ckall"></td>
	<td width="50"> �̹���</td>
	<td width="60">Rack</td>
	<td width="60">No.</td>
	<td width="100">�귣��ID</td>
	<td>��ǰ��</td>
	<td width="60">�ǸŰ�</td>
	<td width="50">���<br>����</td>
	<td width="30">�Ǹ�<br>����</td>
	<td width="30">���<br>����</td>
	<td width="30">�ؿ�<br>����</td>
	<td width="50">���尡��<br>����</td>
	<td width="60">��ǰ����</td>
	<td width="120">��ǰ������</td>
	<td width="40">���</td>
</tr>

<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td align="center"><input type="checkbox" name="check" value="<%= oitem.FItemList(i).Fitemid %>" /></td>
		<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<td align="center"><%= oitem.FItemList(i).Fitemrackcode %></td>
		<td align="center">
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����">
			<%= oitem.FItemList(i).Fitemid %></a>
			</td>
		<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
		<td align="left"><% =oitem.FItemList(i).Fitemname %></td>
		<td align="right">
		<%
			Response.Write "<a href=""javascript:editItemPriceInfo('" & oitem.FItemList(i).Fitemid & "')"" title='�ǸŰ� �� ���ް� ����'>" & FormatNumber(oitem.FItemList(i).Forgprice,0) & "</a>"
			'���ΰ�
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>(��)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'������
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
				end Select
			end if
		%>
		</td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).FdeliverOverseas,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fpojangok,"yn") %></td>
		<td align="center"><%= FormatNumber(oitem.FItemList(i).FitemWeight,0) %>g</td>
		<td align="center">
			<%= oitem.FItemList(i).fvolX %> * <%= oitem.FItemList(i).fvolY %> * <%= oitem.FItemList(i).fvolZ %> cm
		</td>
	    <td align="center"><input type="button" onClick="PopItemWeightEdit('<%= oitem.FItemList(i).Fitemid %>');" value="����" class="button"></td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
    			<% if i>oitem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oitem.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height="300" frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height="0" frameborder="0" scrolling="no"></iframe>
<% end if %>
<%
set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
