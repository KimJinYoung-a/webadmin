<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : [����]�����ǸŰ���
' History	:  �̻� ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim designerid, itemid, dispyn, sellyn, isusing, diffdiv, mwdiv, i, vPurchasetype, tplgubun, isSellStart
dim dispCate, StockMwDiv
	tplgubun	= requestCheckvar(request("tplgubun"),32)
	designerid = requestCheckvar(request("designerid"),32)
	itemid = requestCheckvar(request("itemid"),10)
	dispyn = requestCheckvar(request("dispyn"),2)
	sellyn = requestCheckvar(request("sellyn"),2)
	isusing = requestCheckvar(request("isusing"),1)
	isSellStart = requestCheckvar(request("isSellStart"),1)
	diffdiv = requestCheckvar(trim(request("diffdiv")),32)
	mwdiv = requestCheckvar(request("mwdiv"),2)
	vPurchasetype 	= requestCheckvar(request("purchasetype"),3)
	dispCate = requestCheckvar(request("disp"),16)
	StockMwDiv  	= RequestCheckVar(request("StockMwDiv"),1)

if (diffdiv = "") then diffdiv = "sellN"
if ((request("research") = "") and (isusing = "")) then isusing = "Y"
if (request("research") = "") and tplgubun="" then tplgubun="3X"
if (request("research") = "") and isSellStart="" then isSellStart="Y"

dim osummarystock
set osummarystock = new CSummaryItemStock
	osummarystock.FRectMakerid = designerid
	osummarystock.FRectItemID = itemid
	osummarystock.FRectOnlyIsUsing = isusing
	osummarystock.FRectdiffdiv = diffdiv
	osummarystock.FRectMWDiv = mwdiv
	osummarystock.FRectTplGubun = tplgubun
	osummarystock.FRectIsSellStart = isSellStart
	osummarystock.FRectPurchasetype = vPurchasetype
	osummarystock.FRectDispCate		= dispCate
	osummarystock.FRectStockMwDiv = StockMwDiv
	osummarystock.GetCurrentStockByOnlineBrandDispSell

%>


<!-- obuyprice.GetDispYNSet ���� ���ο� ��� Ŭ������ ��ü�� -->

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function Research(page){
	frm.page.value = page;
	frm.submit();
}

function CheckNSellDispYN(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� ��ǰ�� �����ϴ�.');
		return;
	}

	var ret = confirm('���� ��ǰ�� �����Ͻðڽ��ϱ�?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;

					if (frm.sellyn[0].checked){
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "Y";
					}else if (frm.sellyn[1].checked){
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "S";
					}else{
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "N";
					}

				}
			}
		}
		frm.submit();
	}
}
</script>


<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �귣��: <% drawSelectBoxDesignerwithName "designerid",designerid %>&nbsp;
		&nbsp;
		* ��ǰ�ڵ�: <input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">&nbsp;
		&nbsp;
		* �ŷ�����: <% drawSelectBoxMWU "mwdiv",mwdiv %>&nbsp;
		&nbsp;
		* �������� : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
	</td>	
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!--<input type="radio" name="diffdiv" value="sellSlimit1" <% if (diffdiv = "sellSlimit1") then %>checked<% end if %>>�Ͻ�ǰ��/����1�̻�&nbsp;//-->
		<!--<input type="radio" name="diffdiv" value="sellY0" <% if (diffdiv = "sellY0") then %>checked<% end if %>>�Ǹ�/���� 1�̸�//-->
		<input type="radio" name="diffdiv" value="sellN" <% if (diffdiv = "sellN") then %>checked<% end if %>>ǰ��/��������� 1�̻�
		<input type="radio" name="diffdiv" value="sellSlimit2" <% if (diffdiv = "sellSlimit2") then %>checked<% end if %>>�Ͻ�ǰ��/��������� 1�̻�
		&nbsp;&nbsp;
		* ��뿩��: <% drawSelectBoxUsingYN "isusing", isusing %>
		&nbsp;
		* 3PL���� : <% Call drawSelectBoxTPLGubun("tplgubun", tplgubun) %>
		&nbsp;
		* �Ǹ���ȯ���� : <% Call drawSelectBoxUsingYN("isSellStart", isSellStart) %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �����Ա��� :
		<select class="select" name="StockMwDiv">
			<option value="">����</option>
			<option value="M" <% if (StockMwDiv = "M") then %>selected<% end if %> >M</option>
			<option value="W" <% if (StockMwDiv = "W") then %>selected<% end if %> >W</option>
			<option value="X" <% if (StockMwDiv = "X") then %>selected<% end if %> >��Ÿ</option>
		</select>
		&nbsp;
		* ����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<Br>

<!-- �׼� ���� -->
<form name="frmttl" onsubmit="return false;" style="margin:0px;">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
        <!--
			<input type="button" value="��ü����" onClick="AnSelectAllFrame(true)">&nbsp;<input type="button" value="���û�ǰ����" onClick="CheckNSellDispYN()">        </td>
        -->
	</td>
	<td align="right">	
	</td>
</tr>
<tr>
	<td align="left">
	</td>
</tr>
</table>
</form>
<!-- �׼� �� -->
<style>
th {
  background: #E6E6E6;
  position: sticky;
  top: 0;
  box-shadow: 0 0 1px 0 rgba(0, 0, 0, 0.4);
}
</style>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<thead>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="26">
		�˻���� : <b><%= osummarystock.FresultCount %></b>
	</td>
</tr>
<% if osummarystock.FresultCount > 0 then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<th width="30"></th>
	<th width="50">�̹���</th>
	<th width="120">�귣��</th>
	<th width="50">��ǰ����</th>
	<th width="60">��ǰ�ڵ�</th>
	<th width="50">�ɼ��ڵ�</th>
	<th>��ǰ��<br>(�ɼǸ�)</th>
	<th width="70">�Ǹ���ȯ��</th>
	<th width="70">���ڵ�</th>
	<th width="70">�������ڵ�</th>
	<th width="35">���<br>����</th>
	<th width="35">��ü<br>�԰�<br>��ǰ</th>
	<th width="35">��ü<br>�Ǹ�<br>��ǰ</th>
	<th width="35">��ü<br>���<br>��ǰ</th>
	<th width="35">��Ÿ<br>���<br>��ǰ</th>

	<th width="35">��<br>�ǻ�<br>����</th>
	<th width="35">�ǻ�<br>���</th>
	<th width="35">��<br>�ҷ�</th>
	<th width="35">��ȿ<br>���</th>

	<th width="35">��<br>��ǰ<br>�غ�</th>
	<th width="35">���<br>�ľ�<br>���</th>
	<th width="35">ON<br>����<br>�Ϸ�</th>
	<th width="35">ON<br>�ֹ�<br>����</th>
	<th width="35">����<br>��<br>���</th>
<!--    <th width="35">����<br>����<br>���</th>	-->
	<th width="40">�Ǹ�<br>����</th>
	<th width="50">����<br>����</th>
<!--	<th width="35">ǰ��<br>����</th>	-->
</tr>
</thead>
<tbody>
<% for i=0 to osummarystock.FresultCount-1 %>
<form name="frmBuyPrc_<%= i %>" method="post" onSubmit="return false;" action="/admin/shopmaster/dolimitsoldset.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="<%= osummarystock.FItemList(i).FItemID %>">
<% if osummarystock.FItemList(i).Fisusing="Y" then %>
	<tr bgcolor="#FFFFFF" align="center">
<% else %>
	<tr bgcolor="#EEEEEE" align="center">
<% end if %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
	<td align="left">
		<%= osummarystock.FItemList(i).FMakerID %>
	</td>
	<td><%= osummarystock.FItemList(i).Fitemgubun %></td>
	<td>
		<a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemID %>');"><%= osummarystock.FItemList(i).FItemID %></a>
	</td>
	<td><%= osummarystock.FItemList(i).Fitemoption %></td>
	<td align="left">
		<a href="javascript:PopItemDetail('<%= osummarystock.FItemList(i).FItemID %>','<%= osummarystock.FItemList(i).FItemOption %>')"><%= osummarystock.FItemList(i).FItemName %></a>
		<% if (osummarystock.FItemList(i).FItemOptionName <> "") then %>
			<br>(<%= osummarystock.FItemList(i).FItemOptionName %>)
		<% end if %>
	</td>
	<td><%= chkIIF(isNull(osummarystock.FItemList(i).FsellStdate),"",left(osummarystock.FItemList(i).FsellStdate,10)) %></td>
	<td><%= osummarystock.FItemList(i).FItemRackCode %></td>
	<td><%= osummarystock.FItemList(i).FItemsubrackcode %></td>
	<td><font color="<%= mwdivColor(osummarystock.FItemList(i).Fmwdiv) %>"><%= mwdivName(osummarystock.FItemList(i).Fmwdiv) %></font></td>
	<td><%= osummarystock.FItemList(i).Ftotipgono %></td>
	<td><%= -1*osummarystock.FItemList(i).Ftotsellno %></td>
	<td><%= osummarystock.FItemList(i).Foffchulgono + osummarystock.FItemList(i).Foffrechulgono %></td>
	<td><%= osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono %></td>

	<td align="right"><b><%= FormatNumber(osummarystock.FItemList(i).Ferrrealcheckno, 0) %></b>&nbsp;</td>
	<td align="right"><%= FormatNumber(osummarystock.FItemList(i).getErrAssignStock, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Ferrbaditemno, 0) %>&nbsp;</td>
	<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Frealstock, 0) %>&nbsp;</td>

	<td><%= osummarystock.FItemList(i).Fipkumdiv5 + osummarystock.FItemList(i).Foffconfirmno %></td>
	<td><b><%= osummarystock.FItemList(i).GetCheckStockNo %></b></td>
	<td><%= osummarystock.FItemList(i).Fipkumdiv4 %></td>
	<td><%= osummarystock.FItemList(i).Fipkumdiv2 %></td>
	<td><b><%= osummarystock.FItemList(i).GetLimitStockNo %></b></td>
<!--    <td><b><%= round(osummarystock.FItemList(i).GetLimitStockNo * 0.95,0) %></b></td>	-->
<!--    <td><b><font color="red"><%= osummarystock.FItemList(i).GetLimitStockNo - osummarystock.FItemList(i).GetLimitStr %></font></b></td>	-->
	<td>
		<%= osummarystock.FItemList(i).Fsellyn %>
		<!--
		<input type="radio" name="sellyn" value="Y" <% if osummarystock.FItemList(i).Fsellyn="Y" then response.write "checked" %>>Y
		<input type="radio" name="sellyn" value="N" <% if osummarystock.FItemList(i).Fsellyn="N" then response.write "checked ><font color=red>N</font>" else response.write ">N" %>
		-->
	</td>

	<td>
	<% if (osummarystock.FItemList(i).Flimityn = "Y") then %>
		����(<%= osummarystock.FItemList(i).GetLimitStr %>)
		<% if (osummarystock.FItemList(i).Foptlimityn = "Y") then %>
		<br>(<%= osummarystock.FItemList(i).Foptlimitno %>/<%= osummarystock.FItemList(i).Foptlimitsold %>)
		<% else %>
		<br>(<%= osummarystock.FItemList(i).FLimitNo %>/<%= osummarystock.FItemList(i).FLimitSold %>)
		<% end if %>
	<% end if %>
	</td>
<!--    <td><% if osummarystock.FItemList(i).IsSoldOut  then %><font color="red">ǰ��</font><% end if %></td>	-->
</tr>
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="26" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</tbody>
</table>

<form name="frmArrupdate" method="post" action="/admin/shopmaster/dolimitsoldset.asp" style="margin:0px;">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="sellyn" value="">
</form>
<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
