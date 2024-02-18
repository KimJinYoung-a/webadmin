<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->

<%

dim designerid, itemid,  sellyn, isusing
dim SearchMode

designerid  = request("designerid")
itemid      = request("itemid")
sellyn      = request("sellyn")
isusing     = request("isusing")
SearchMode  = request("SearchMode")

if ((request("research") = "") and (isusing = "")) then
        isusing = "on"
        SearchMode = "S1"
end if


dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FPageSize=300
osummarystock.FRectMakerid = designerid
osummarystock.FRectItemID = itemid
osummarystock.FRectOnlyIsUsing = isusing
osummarystock.FRectSearchMode = SearchMode
osummarystock.GetCurrentStockByOnlineBrandLimitSoldout

dim i

%>


<script language='javascript'>
function CheckThisRow(comp){
    var frm = comp.form;
    frm.cksel.checked = true;
    AnCheckClick(frm.cksel);
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function changecontent(){
	// nothing
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
    upfrm.sellyn.value = "";
    upfrm.itemid.value = "";
	var ret = confirm('���� ��ǰ�� �����Ͻðڽ��ϱ�?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;

					if (frm.sellyn[0].checked){
					    alert('���� ǰ���� ��ǰ�� �Ǹ��� �� �����ϴ�.');
					    frm.sellyn[0].focus();
					    return;
						//upfrm.sellyn.value = upfrm.sellyn.value + "|" + "Y";
					}else if (frm.sellyn[1].checked){
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "S";
					}else{
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "N";
					}
                    /*
					if (frm.dispyn[0].checked){
						upfrm.dispyn.value = upfrm.dispyn.value + "|" + "Y";
					}else{
						upfrm.dispyn.value = upfrm.dispyn.value + "|" + "N";
					}
					*/
				}
			}
		}
		frm.submit();
	}
}
</script>

<!-- ����� ���� ���� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("menubar") %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>�����Ǹ� ǰ�� ����</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			�����Ǹ� ��ǰ�� ǰ���� ��ǰ�� ���� �����Դϴ�.
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>
<!-- ����� ���� ���� �� -->

<p>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	�귣�� : <% drawSelectBoxDesignerwithName "designerid",designerid %>&nbsp;
        	��ǰ�ڵ� : <input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">&nbsp;
        	<input type=checkbox name="isusing" value="on" <% if isusing="on" then response.write "checked" %> >����ǰ��


        	<input type="radio" name="SearchMode" value="Y0" <%= ChkIIF(SearchMode="Y0","checked","") %> > �Ǹ�Y, ����Y, ����0&nbsp;&nbsp;
          	<input type="radio" name="SearchMode" value="S1" <%= ChkIIF(SearchMode="S1","checked","") %> > �Ǹ�S, ����Y, ����1�̻�&nbsp;&nbsp;

        </td>
        <td align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a><br>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frmttl" onsubmit="return false;">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	�˻���� : <b><%= FormatNumber(osummarystock.FresultCount,0) %></b> (�ִ� : <%= osummarystock.FPageSize %>)
        </td>
        <td align="right">
        	<input type="button" value="��ü����" onClick="AnSelectAllFrame(true)">&nbsp;<input type="button" value="���û�ǰ����" onClick="CheckNSellDispYN()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>����</td>
		<td width="50">�̹���</td>
		<td width="70">�귣��</td>
		<td width="40">��ǰ<br>�ڵ�</td>
		<td>��ǰ��<br>(�ɼǸ�)</td>
		<td width="35">���<br>����</td>
        <td width="35">��ü<br>�԰�<br>��ǰ</td>
        <td width="35">��ü<br>�Ǹ�<br>��ǰ</td>
        <td width="35">��ü<br>���<br>��ǰ</td>
        <td width="35">��Ÿ<br>���<br>��ǰ</td>
<!--    <td width="35">�ý���<br>���</td>	-->
		<td width="35">��<br>�ҷ�</td>
<!--    <td width="35">��ȿ<br>���</td>	-->
        <td width="35">��<br>�ǻ�<br>����</td>
        <td width="35">�ǻ�<br>���</td>
        <td width="35">��<br>��ǰ<br>�غ�</td>
        <td width="35">���<br>�ľ�<br>���</td>
        <td width="35">ON<br>����<br>�Ϸ�</td>
        <td width="35">ON<br>�ֹ�<br>����</td>
        <td width="35">����<br>��<br>���</td>
<!--    <td width="35">����<br>����<br>���</td>	-->
        <td width="50">����<br>����</td>
		<td width="50">�Ǹ�<br>����</td>
		<td width="60">����<br>����</td>
		<td width="35">ǰ��<br>����</td>
		<td width="35">����<br>����</td>
    </tr>
<% for i=0 to osummarystock.FresultCount-1 %>
	<form name="frmBuyPrc_<%= osummarystock.FItemList(i).FItemID %>" method="post" onSubmit="return false;" action="dolimitsoldset.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemid" value="<%= osummarystock.FItemList(i).FItemID %>">
	<% if osummarystock.FItemList(i).Fisusing="Y" then %>
    <tr bgcolor="#FFFFFF" align="center">
    <% else %>
    <tr bgcolor="#EEEEEE" align="center">
    <% end if %>
    	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <%= ChkIIf (osummarystock.FItemList(i).FItemOptionName <> "","disabled","") %> ></td>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
		<td align="left">
          <%= osummarystock.FItemList(i).FMakerID %>
        </td>
		<td>
          <a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemID %>');"><%= osummarystock.FItemList(i).FItemID %></a>
        </td>
		<td align="left">
          <a href="javascript:PopItemDetail('<%= osummarystock.FItemList(i).FItemID %>','<%= osummarystock.FItemList(i).FItemOption %>')"><%= osummarystock.FItemList(i).FItemName %></a>
        <% if (osummarystock.FItemList(i).FItemOptionName <> "") then %>
          <br><font color="blue">(<%= osummarystock.FItemList(i).FItemOptionName %>)</font>
        <% end if %>
        </td>
        <td><%= osummarystock.FItemList(i).GetMwDivName %></td>
		<td><%= osummarystock.FItemList(i).Ftotipgono %></td>
		<td><%= -1*osummarystock.FItemList(i).Ftotsellno %></td>
		<td><%= osummarystock.FItemList(i).Foffchulgono + osummarystock.FItemList(i).Foffrechulgono %></td>
        <td><%= osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono %></td>
<!--    <td><%= osummarystock.FItemList(i).Ftotsysstock %></td>	-->
        <td><%= osummarystock.FItemList(i).Ferrbaditemno %></td>
<!--    <td><%= osummarystock.FItemList(i).Favailsysstock %></td>	-->
        <td><%= osummarystock.FItemList(i).Ferrrealcheckno %></td>
        <td><b><%= osummarystock.FItemList(i).Frealstock %></b></td>
        <td><%= osummarystock.FItemList(i).Fipkumdiv5 + osummarystock.FItemList(i).Foffconfirmno %></td>
        <td><b><%= osummarystock.FItemList(i).GetCheckStockNo %></b></td>
        <td><%= osummarystock.FItemList(i).Fipkumdiv4 %></td>
        <td><%= osummarystock.FItemList(i).Fipkumdiv2 %></td>
        <td><b><%= osummarystock.FItemList(i).GetLimitStockNo %></b></td>
<!--        <td><b><%= round(osummarystock.FItemList(i).GetLimitStockNo * 0.95,0) %></b></td>	-->

<!--        <td><b><font color="red"><%= osummarystock.FItemList(i).GetLimitStockNo - osummarystock.FItemList(i).GetLimitStr %></font></b></td>	-->
        <td>

        </td>
        <td>
			<input type="radio" name="sellyn" value="Y" onClick="CheckThisRow(this);" <% if osummarystock.FItemList(i).Fsellyn="Y" then response.write "checked" %> >Y
			<input type="radio" name="sellyn" value="S" onClick="CheckThisRow(this);" <% if osummarystock.FItemList(i).Fsellyn="S" then response.write "checked" %> >S
			<input type="radio" name="sellyn" value="N" onClick="CheckThisRow(this);" <% if osummarystock.FItemList(i).Fsellyn="N" then response.write "checked ><font color=red>N</font>" else response.write ">N" %>
        </td>

        <td>
          	����(<%= osummarystock.FItemList(i).GetLimitStr %>)
            <% if (osummarystock.FItemList(i).Foptlimityn = "Y") then %>
            <br>(<%= osummarystock.FItemList(i).Foptlimitno %>/<%= osummarystock.FItemList(i).Foptlimitsold %>)
            <% else %>
            <br>(<%= osummarystock.FItemList(i).FLimitNo %>/<%= osummarystock.FItemList(i).FLimitSold %>)
          	<% end if %>
        </td>
        <td><% if osummarystock.FItemList(i).IsSoldOut  then %><font color="red">ǰ��</font><% end if %></td>
        <td>
            <% if osummarystock.FItemList(i).FDanjongyn="Y" then %>
            <font color="red">����</font>
            <% elseif osummarystock.FItemList(i).FDanjongyn="S" then %>
            <font color="blue">�Ͻ�<br>ǰ��</font>
            <% else %>
            <% end if %>
        </td>
	</tr>
	</form>
<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<form name="frmArrupdate" method="post" action="dolimitsoldset.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="dispyn" value="">
<input type="hidden" name="sellyn" value="">
</form>

<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->