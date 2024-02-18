<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/baditemcls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim itemid, makerid, mode, actType

itemid  = requestCheckVar(request("itemid"),32)     '' length > 9
makerid = requestCheckVar(request("makerid"),32)
mode    = requestCheckVar(request("mode"),9)
actType = requestCheckVar(request("actType"),9)     '' actType="actloss" �ν�ó�� actType<>"actloss" ��ǰó��

dim osummarystock
set osummarystock = new CSummaryItemStock
if (Len(itemid) = 12) then
        osummarystock.FRectItemID =  Mid(itemid,3,6)
end if
osummarystock.FRectmakerid = makerid
osummarystock.FRectSearchType = "err"

if (makerid<>"") then
osummarystock.GetDailyErrItemListByBrand
end if

if (osummarystock.FResultCount > 0) then
        makerid = osummarystock.FItemList(0).Fmakerid
end if

dim i

dim LorRText
if (actType="actloss") then
    LorRText = "�ν�"
else
    LorRText = "��ǰ"
end if

%>
<script language='javascript'>


function getOnLoad(){
	document.frm.itemid.focus();
	document.frm.itemid.select();
}

window.onload=getOnLoad;

function checkhL(e){
    if (e.value*1 != 0){
        hL(e);
    }else{
        dL(e);
    }
}

function SubmitSearchItem() {
        if (document.frm.itemid.value.length != 12) {
                if (document.frm.makerid.selectedIndex == 0) {
                        alert("�귣�� �Ǵ� ��ǰ�ڵ带 �Է��ϼ���.");
                        return;
                }
                document.frm.itemid.value = "";
                document.frm.submit();
        } else {
                document.frm.makerid.selectedIndex = 0;
                document.frm.submit();
        }
}

function SubmitInsert(){
    <% if (osummarystock.FResultCount < 1) then %>
        alert("�˻��� ��ǰ�� �����ϴ�.");
        return;
    <% else %>
        if (document.frm.itemid.value.length != 12) {
			alert("��ǰ�ڵ带 ��Ȯ�� �Է��ϼ���.");
			return;
        }

		var frm = document.frm;
		var itembarcode = frm.itemid.value;

		for (var i = 0; ; i++) {
			var itemgubun = document.getElementById("itemgubun_" + i);
			var itemid = document.getElementById("itemid_" + i);
			var itemoption = document.getElementById("itemoption_" + i);

			var itemno = document.getElementById("itemno_" + i);
			var itemmaxno = document.getElementById("itemmaxno_" + i);

			if (itemgubun == undefined) {
				alert("��ǰ�� ��Ͽ� �����ϴ�. �ٸ� �귣���̰ų�, ��������� �Ǿ� ���� �ʽ��ϴ�.");
				break;
			}

			if ((itemgubun.value == itembarcode.substring(0,2)) && (itemid.value*1 == (1 * itembarcode.substring(2,8))) && (itemoption.value == itembarcode.substring(8))) {
				itemno.value = (1 * itemno.value) + 1;

				/*
				if ((1 * itemno.value) > (itemmaxno.value * -1)) {
					itemno.value = (itemmaxno.value * -1);
					alert("������ϵ� �������� ������ Ů�ϴ�. ���� ��������� �ϼ���.");
				}
				*/

				hL(itemno);
				break;
			}
		}

        frm.itemid.select();
        frm.itemid.focus();
    <% end if %>
}

function SubmitCheckInsert(v) {
	var curridx = v.value;
	var itemno = document.getElementById("itemno_" + curridx);
	var itemmaxno = document.getElementById("itemmaxno_" + curridx);

	if (v.checked == true) {
		itemno.value = itemmaxno.value*-1;
	} else {
		itemno.value = 0;
	}
	checkhL(itemno);
}

function SubmitCheckInsertAll(v) {
	for (var i = 0;; i++) {
		var chk = document.getElementById("chk_" + i);
		if (chk == undefined) {
			return;
		}
		chk.checked = v.checked;
		SubmitCheckInsert(chk);
	}
}

function CheckInsert(f, maxvalue){
        alert(f.value);
        if (f.value = "") {
                return;
        }

        if (f.value < 0) {
                alert("����<%= LorRText %>������ ���̳ʽ��� �ɼ� �����ϴ�.");
                f.value = 0;
                return;
        }

        if (f.value > (maxvalue * -1)) {
                alert("������ϵ� �������� ������ Ů�ϴ�. ���� ��������� �ϼ���.");
                f.value = (maxvalue * -1);
                return;
        }


}

function SubmitList(){
	window.open('/common/pop_item_search.asp','pop_item_search','width=900,height=600');
}


function ReActItems(itemgubunarr,
                    itemarr,
                    itemoptionarr,
                    sellcasharr,
                    suplycasharr,
                    buycasharr,
                    itemnoarr,
                    itemnamearr,
                    itemoptionnamearr,
                    designerarr,
                    mwdivarr)
{
        document.frm.itemgubunarr.value = itemgubunarr;
        document.frm.itemidarr.value = itemarr;
        document.frm.itemoptionarr.value = itemoptionarr;
        document.frm.itemnoarr.value = itemnoarr;

        document.frm.method = "post";
        document.frm.mode.value = "arrinsert";
        document.frm.action = "do_bad_item_input.asp";
        document.frm.submit();

        return true;
}





function SubmitUpdateAll(){
    var pmwdiv = "";

    var frm = document.frm;

    frm.itemgubunarr.value = "";
    frm.itemidarr.value = "";
    frm.itemoptionarr.value = "";
    frm.itemnoarr.value = "";

	for (var i = 0; ; i++) {
		var itemgubun = document.getElementById("itemgubun_" + i);
		var itemid = document.getElementById("itemid_" + i);
		var itemoption = document.getElementById("itemoption_" + i);

		var itemno = document.getElementById("itemno_" + i);
		var mwdiv = document.getElementById("mwdiv_" + i);

		if (itemgubun == undefined) {
			break;
		}

		if (itemno.value*1 != 0) {
			if (pmwdiv == "") {
				pmwdiv = mwdiv.value;
			} else {
				// ��ǰ�� ���
				/*
				if (pmwdiv != mwdiv.value) {
					alert("���� �Ӽ��� �ٸ���ǰ�� ���� ó�� �� �� �����ϴ�.");
					return;
				}
				*/
			}

			frm.itemgubunarr.value = frm.itemgubunarr.value + itemgubun.value + "|";
			frm.itemidarr.value = frm.itemidarr.value + itemid.value + "|";
			frm.itemoptionarr.value = frm.itemoptionarr.value + itemoption.value + "|";
			frm.itemnoarr.value = frm.itemnoarr.value + itemno.value + "|";
		}
	}

	if (frm.itemgubunarr.value == "") {
        alert("<%= LorRText %>ó���� ��ǰ�� �����ϴ�.");
        return;
    }

    if (confirm('<%= LorRText %> �������� �ۼ��Ͻðڽ��ϱ�?')){
        document.frm.method = "post";
        <% if (actType="actloss") then %>
        document.frm.mode.value = "lossarr";
        <% else %>
        document.frm.mode.value = "notused";
        <% end if %>
        document.frm.pmwdiv.value = pmwdiv;
        document.frm.action = "/common/erritem_re_input_process.asp";
        document.frm.submit();
    }
}
</script>

<!-- �˻� ���� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" >
<tr bgcolor="#FFFFFF">
    <td>** ������� ��ǰ <strong>�ν�</strong> ó��</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="" onsubmit="return false;">
	<input type="hidden" name="actType" value="<%= actType %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="brandid" value="<%= makerid %>">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="pmwdiv" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣��� : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
			<input type="button" class="button_s" value="�귣�������ϻ�ǰ��ϰ˻�" onClick="SubmitSearchItem()">
		</td>
	</tr>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td>
			��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="12" size="14" onKeyPress="if (event.keyCode == 13) { SubmitInsert(); return false; }">
			<input type="button" class="button" value="�ν��߰�" onclick="SubmitInsert()">
        	&nbsp;
			* �������� �귣�常 �ϰ�ó�� �����մϴ�.
		</td>
		<td align="right">
			<input type="button" class="button" value="��ü����" onclick="SubmitUpdateAll()">
		</td>
	</tr>
	</form>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="20"><input type="checkbox" name="chkall" value="" onClick="SubmitCheckInsertAll(this);"></td>
      <td width="100">�귣��ID</td>
      <td width="40">����<br>����</td>
      <td width="25">����</td>
      <td width="40">��ǰ<br>�ڵ�</td>
      <td width="30">�ɼ�</td>
      <td>�����۸�</td>
      <td>�ɼǸ�</td>
      <td width="50">�Һ��ڰ�</td>
      <td width="40">����<br>����</td>
      <td width="40">���<br>����</td>
    </tr>
    <form name="frmlist" method=get action="" onsubmit="return false;">
<% for i=0 to osummarystock.FResultCount - 1 %>
    <tr align="center" bgcolor="<%= chkIIF(osummarystock.FItemList(i).FItemgubun="10","#FFFFFF","#EEEEEE") %>">
      <td><input type="checkbox" id="chk_<%= i %>" name="chk" value="<%= i %>" onClick="SubmitCheckInsert(this);"></td>
      <td><%= osummarystock.FItemList(i).Fmakerid %></td>
      <td>
        <% if osummarystock.FItemList(i).FItemgubun="10" then %>
        <font color="<%= mwdivColor(osummarystock.FItemList(i).FMwdiv) %>"><%= osummarystock.FItemList(i).GetMwDivName %></font>
        <% end if %>
      </td>
      <td><%= osummarystock.FItemList(i).FItemgubun %></td>
      <td><%= osummarystock.FItemList(i).FItemid %></td>
      <td><%= osummarystock.FItemList(i).FItemoption %></td>
      <td align="left"><%= osummarystock.FItemList(i).FItemname %></td>
      <td align="left"><%= osummarystock.FItemList(i).FItemOptionName %></td>
      <td align="right"><%= formatnumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
      <td>
        <%= osummarystock.FItemList(i).Ferrrealcheckno %>
      </td>
      <input type="hidden" id="itemgubun_<%= i %>" name="itemgubun" value="<%= osummarystock.FItemList(i).FItemgubun %>">
      <input type="hidden" id="itemid_<%= i %>" name="itemid" value="<%= osummarystock.FItemList(i).FItemid %>">
      <input type="hidden" id="itemoption_<%= i %>" name="itemoption" value="<%= osummarystock.FItemList(i).FItemOption %>">
      <td>
        <input type="text" class="text" id="itemno_<%= i %>" name="itemno" value="0" size="3" onKeyUP="checkhL(this);" <%= chkIIF(osummarystock.FItemList(i).FItemgubun="10","","disabled") %> >
      </td>
      <input type="hidden" id="itemmaxno_<%= i %>" name="itemmaxno" value="<%= osummarystock.FItemList(i).Ferrrealcheckno %>" >
      <input type="hidden" id="mwdiv_<%= i %>" name="mwdiv" value="<%= osummarystock.FItemList(i).FMwdiv %>">
    </tr>
   	<% next %>
<% if osummarystock.FResultCount = 0 then %>
    <tr align="center" bgcolor="#FFFFFF">
      <td colspan="11" align="center">�˻��� ��ǰ�� �����ϴ�.</td>
    </tr>
<% end if %>
    </form>
</table>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->


<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->