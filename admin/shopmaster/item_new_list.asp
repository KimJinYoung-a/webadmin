<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : [ON]��ǰ����>>�ǸŴ���ǰLIST
' History : �̻� ����
'			2023.10.4 �ѿ�� ����(�����α��߰�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemlistcls.asp"-->
<%
Dim oitem,ix,page,makerid, lp, ipgoyn, itemid,research, cdl, cdm, cds, dispCate, ttlCount
	makerid = requestCheckvar(request("makerid"),32)
	page = requestCheckvar(request("page"),10)
	ipgoyn = requestCheckvar(request("ipgoyn"),10)
	itemid = requestCheckvar(request("itemid"),10)
	cdl = requestCheckvar(request("cdl"),10)
	cdm = requestCheckvar(request("cdm"),10)
	cds = requestCheckvar(request("cds"),10)
	dispCate = requestCheckvar(request("disp"),16)

if (page="") then page=1
if (ipgoyn="") then ipgoyn="Y"

set oitem = new CItemList
	oitem.FPageSize = 50
	oitem.FCurrPage = page
	oitem.FRectMakerid = makerid
	oitem.FRectItemid = itemid
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectDispCate = dispCate

	if ipgoyn="U" then
		oitem.FRectIpgoGubun = "N"
		oitem.FRectDeliverType = "U"
	else
		oitem.FRectIpgoGubun = ipgoyn
		oitem.FRectDeliverType = "T"
	end if

	''2016/12/27 by eastone �귣�庰 �˻����� �߰�. ��������/���� �귣�� ���� �����.
	if (ipgoyn="BY") or (ipgoyn="BN") then
		oitem.getSellWaitItemListByBrand
	elseif ipgoyn = "GIY" or ipgoyn = "GIN" then
		oitem.GetgiftItemIpgo
	else
		oitem.getSellWaitItemList
	end if

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function goPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
function ViewItemDetail(itemno){
	window.open('/admin/viewitem.asp?itemid='+itemno ,'window1','width=790,height=600,scrollbars=yes,status=no');
}
function insertdb(itemid,itemname){
 //if (confirm(itemname + "�� ����Ͻðڽ��ϱ�?") == true){
    //location.href("item_insertdb.asp?itemid="+itemid);
 //}
}
function WaitState(itemid){
	var ret = confirm('��ϴ��� �����Ͻðڽ��ϱ�?');

	if (ret){
		document.location = 'doitemregboru.asp?mode=waitstate&idx=' + itemid;
	}
}

function popoptionEdit(iid){
	var popwin = window.open('popitemoptionedit.asp?menupos=<%= menupos %>&itemid=' + iid,'popitemoptionedit','width=440 height=500 scrollbars=yes resizable=yes');
	popwin.focus();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function fnResearchSellWait(imakerid, iipgoyn) {
	var frm = document.frm;
    frm.makerid.value=imakerid;
    // frm.ipgoyn.value=iipgoyn;
	jsSetRadioValue('ipgoyn', iipgoyn);

    frm.submit();
}

function jsSearchBrand(makerid) {
	if(makerid != "") {
		document.frm.makerid.value = makerid;
	}
	document.frm.submit();
}

function jsSetRadioValue(name, value) {
	var obj = document.getElementById(name + "_" + value);
	if (obj != undefined) {
		obj.checked = true;
	}
}

function jsSetSellY(makerid) {
	var frm = document.frmAct;
	var i, obj, itemidArr;

	if (makerid == '') {
		alert('���� �귣�带 �˻��ϼ���.');
		return;
	}

	itemidArr = '';
	for (i = 0; ; i++) {
		obj = document.getElementById('itemid_' + i);
		if (obj == undefined) { break; }
		if (obj.disabled == true) { continue; }
		if (obj.checked != true) { continue; }
		itemidArr = (itemidArr == '') ? obj.value : itemidArr + ',' + obj.value;
	}

	if (itemidArr == '') {
		alert('���õ� ��ǰ�� �����ϴ�.');
		return;
	}

	if (confirm('���û�ǰ�� �Ǹ���ȯ�Ͻðڽ��ϱ�?') == true) {
		frm.mode.value = 'modisellY';
		frm.itemidArr.value = itemidArr;
		frm.submit();
	}
}

function jsChkAll(o) {
	var chk = o.checked;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('itemid_' + i);
		if (obj == undefined) { break; }
		if (obj.disabled == true) { continue; }
		obj.checked = chk;
	}
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="research" value="<%= research %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣��ID : <% drawSelectBoxDesignerWithName "makerid",makerid %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		&nbsp;
		����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		&nbsp;
		��ǰ��ȣ : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="6" maxlength="9">

	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="radio" id="ipgoyn_BY" name="ipgoyn" value="BY" <% if ipgoyn="BY" then response.write "checked" %>>�԰�Ϸ��ǰ(�귣�庰)
		<input type="radio" id="ipgoyn_BN" name="ipgoyn" value="BN" <% if ipgoyn="BN" then response.write "checked" %>>�԰����ǰ(�귣�庰)
		&nbsp;&nbsp;
		<input type="radio" id="ipgoyn_Y" name="ipgoyn" value="Y" <% if ipgoyn="Y" then response.write "checked" %>>�԰�Ϸ��ǰ
		<input type="radio" id="ipgoyn_N" name="ipgoyn" value="N" <% if ipgoyn="N" then response.write "checked" %>>�԰����ǰ
		<input type="radio" name="ipgoyn" value="U" <% if ipgoyn="U" then response.write "checked" %>>��ü��ۻ�ǰ(����� ���԰��ǰ)
		&nbsp;&nbsp;
		<input type="radio" name="ipgoyn" value="GIY" <% if ipgoyn="GIY" then response.write "checked" %>>����ǰ�԰�Ϸ��ǰ
		<input type="radio" name="ipgoyn" value="GIN" <% if ipgoyn="GIN" then response.write "checked" %>>����ǰ�԰����ǰ
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br />

<!-- ����Ʈ ���� -->
<% if (ipgoyn="BY") or (ipgoyn="BN") then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><% = oitem.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="300">�귣��ID</td>
		<td >��������</td>
		<td width="100">�Ǽ�</td>
	</tr>
	<% if oitem.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
	<% else %>
	<% for ix=0 to oitem.FresultCount-1 %>
		<% ttlCount=ttlCount+oitem.FItemList(ix).FCount %>
		<tr class="a" bgcolor="#FFFFFF">
			<td align="center"><a href="javascript:fnResearchSellWait('<%= oitem.FItemList(ix).Fmakerid %>','<%=right(ipgoyn,1)%>');"><%= oitem.FItemList(ix).Fmakerid %></a></td>
			<td align="center"><a href="javascript:fnResearchSellWait('<%= oitem.FItemList(ix).Fmakerid %>','<%=right(ipgoyn,1)%>');"><%= oitem.FItemList(ix).fpurchasetypename %></a></td>
			<td align="center"><a href="javascript:fnResearchSellWait('<%= oitem.FItemList(ix).Fmakerid %>','<%=right(ipgoyn,1)%>');"><%= FormatNumber(oitem.FItemList(ix).FCount,0) %></a></td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF">
			<td ></td>
			<td align="center"></td>
			<td align="center"><%= FormatNumber(ttlCount,0) %></td>
		</tr>
	<% end if %>
</table>
<% else %>
	<input type="button" class="button" value="���û�ǰ �Ǹ���ȯ" onClick="jsSetSellY('<%= makerid %>')">

	<br />

	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			�˻���� : <b><% = oitem.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= oitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="chkAll" onClick="jsChkAll(this);"></td>
		<td width="80">�귣��ID</td>
		<td width="40">��ǰID</td>
		<td width="50">�̹���</td>
		<td >��ǰ��</td>
		<td >�ɼ�</td>
		<td width="40">�ǸŰ�</td>
		<td width="40">���԰�</td>
		<td width="50">����</td>
		<td width="70">�����</td>
		<td width="50">��������</td>
		<td width="30">�Ǹ�<br>��</td>
		<td width="30">���<br>��</td>
		<td width="30">�ֹ�<br>��</td>
		<td width="30">��<br>�԰�<br>��</td>
		<td width="30">����<br>���</td>
	</tr>
	<% if oitem.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="16" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
	<% else %>
	<% for ix=0 to oitem.FresultCount-1 %>

	<% if (oitem.FItemList(ix).FSellyn="Y") then %>
		<tr class="a" bgcolor="#EEEEEE">
			<td align="center"><input id="itemid_<%= ix %>" type="checkbox" name="chkitem" value="" disabled>
	<% else %>
		<tr class="a" bgcolor="#FFFFFF">
			<td align="center"><input id="itemid_<%= ix %>" type="checkbox" name="chkitem" value="<%= oitem.FItemList(ix).Fitemid %>">
	<% end if %>
			<td align="center"><a href="javascript:jsSearchBrand('<%= oitem.FItemList(ix).Fmakerid %>')"><%= oitem.FItemList(ix).Fmakerid %></a></td>
			<td align="center"><a href="javascript:PopItemSellEdit('<%= oitem.FItemList(ix).Fitemid %>')"><%= oitem.FItemList(ix).Fitemid %></a></td>
			<td align="center"><img src="<%= oitem.FItemList(ix).FImageSmall %>" width="50" height="50" border="0" alt=""></td>
			<td align="center"><a target=_blank href="/admin/itemmaster/itemmodify.asp?itemid=<% =oitem.FItemList(ix).Fitemid %>&makerid=<%= oitem.FItemList(ix).Fmakerid %>"><% =oitem.FItemList(ix).Fitemname %></a><br><a target=_blank href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<% =oitem.FItemList(ix).Fitemid %>"><font color="blue">(�����κ���)</font></a></td>
			<td align="center"><%= oitem.FItemList(ix).FItemOptionname %></td>
			<td align="right"><%= FormatNumber(oitem.FItemList(ix).FSellPrice,0) %></td>
			<td align="right"><%= FormatNumber(oitem.FItemList(ix).FBuyPrice,0) %></td>
			<td align="center">
				<font color="<%= oitem.FItemList(ix).getMwDivColor %>"><%= oitem.FItemList(ix).getMwDivName %></font>
				<% if oitem.FItemList(ix).FSellPrice<>0 then %>
				<%= 100-CLng(oitem.FItemList(ix).FBuyPrice/oitem.FItemList(ix).FSellPrice*100*100)/100 %>
				<% end if %>
			</td>
			<td align="center"><%= FormatDateTime(oitem.FItemList(ix).Fregdate,2) %></td>
			<td align="center">
			<% if oitem.FItemList(ix).FSellyn="N" then %>
			�Ǹ�<font color=red>x</font><br>
			<% end if %>
			<% if oitem.FItemList(ix).FLimityn="Y" then %>
			<font color=red>����</font><%= oitem.FItemList(ix).FLimitNo-oitem.FItemList(ix).FLimitSold %>
			<% end if %>
			</td>
			<td align="center"><%= FormatNumber(oitem.FItemList(ix).Fsellno,0) %></td>
			<td align="center"><%= FormatNumber(oitem.FItemList(ix).Fchulno,0) %></td>
			<td align="center"><%= FormatNumber(oitem.FItemList(ix).Fpreorderno,0) %></td>
			<td align="center"><%= FormatNumber(oitem.FItemList(ix).Fipgono,0) %></td>
			<td align="center"><%= FormatNumber(oitem.FItemList(ix).Fcurrno,0) %></td>
		</tr>
		<% next %>
	<% end if %>
	<tr bgcolor="#FFFFFF">
		<td colspan="16" align="center">
			<% if oitem.HasPreScroll then %>
		<a href="javascript:goPage(<%= oitem.StartScrollPage-1 %>)">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
		<% if lp>oitem.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage(<%= lp %>)">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oitem.HasNextScroll then %>
		<a href="javascript:goPage(<%= lp %>)">[next]</a>
	<% else %>
		[next]
	<% end if %>

		</td>
	</tr>
	</table>
<% end if %>

<form name="frmAct" method="post" action="item_new_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemidArr" value="">
</form>

<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
