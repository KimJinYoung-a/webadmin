<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ν����
' History : �̻� ����
'           2021.04.06 �ѿ�� ����(�˻����Ǽ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/baditemcls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim itemid, makerid, mode, actType, searchtype, purchasetype, mwdiv, sellyn, onlyisusing, makeruseyn, itemgubun
dim datetype, centermwdiv, monthlymwdiv, yyyy, mm, i
	itemid  	= requestCheckVar(request("itemid"),32)     '' length > 9
	makerid 	= requestCheckVar(request("makerid"),32)
	mode    	= requestCheckVar(request("mode"),9)
	searchtype 	= requestCheckVar(request("searchtype"),9)     	'' searchtype="bad" �ҷ� actType<>"err" ����
	purchasetype 	= requestcheckvar(request("purchasetype"),1)
	actType 	= requestCheckVar(request("actType"),32)     	'' actType="actloss" �ν�ó�� actType="actshopchulgo" (actType<>"actloss" and actType<>"actshopchulgo") ��ǰó��
	mwdiv 			= requestcheckvar(request("mwdiv"),1)
	sellyn 			= requestcheckvar(request("sellyn"),1)
	onlyisusing 	= requestcheckvar(request("onlyisusing"),1)
	makeruseyn	 	= requestcheckvar(request("makeruseyn"),1)
	itemgubun 		= requestcheckvar(request("itemgubun"),3)
	datetype 		= requestcheckvar(request("datetype"),8)
	yyyy 			= requestcheckvar(request("yyyy1"),4)
	mm 				= requestcheckvar(request("mm1"),2)
	centermwdiv		= requestcheckvar(request("centermwdiv"),1)
	monthlymwdiv	= requestcheckvar(request("monthlymwdiv"),1)

'if (makerid <> "" and itemgubun = "") then
'	itemgubun = "10"
'end if
datetype = "yyyymm"
' ������ϰ��
if yyyy = Left(now(),4) and mm = mid(now(),6,2) then
	datetype = "curr"
end if
dim osummarystock
set osummarystock = new CSummaryItemStock

if (Len(itemid) = 12) then
	osummarystock.FRectItemID =  Mid(itemid,3,6)
end if
if (Len(itemid) = 14) then
        osummarystock.FRectItemID =  Mid(itemid,3,8)
end if

osummarystock.FPageSize=100                 ''�߰� 2016/08/04
osummarystock.FRectmakerid = makerid
osummarystock.FRectSearchType = searchtype
osummarystock.FRectMWDiv = mwdiv
osummarystock.FRectlastmwdiv = monthlymwdiv
osummarystock.FRectSellYN = sellyn
osummarystock.FRectOnlyIsUsing = onlyisusing
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectPurchaseType = purchasetype
osummarystock.FRectMakerUseYN = makeruseyn
osummarystock.FRectCenterMWDiv = centermwdiv
osummarystock.FRectDatetype   = datetype
osummarystock.FRectYYYYMM = yyyy+"-"+mm

if (makerid<>"") or (Len(itemid) = 12) or (Len(itemid) = 14) then
    osummarystock.FPageSize=500
	osummarystock.GetBadOrErrItemListByBrand
end if

if (osummarystock.FResultCount > 0) and (makerid = "") then
    makerid = osummarystock.FItemList(0).Fmakerid
end if

dim IsReturnOK : IsReturnOK = True
dim IscheckReturnOK : IscheckReturnOK = True
dim opartner, ogroup
if (searchtype="bad") and (actType<>"actloss") and (actType<>"actshopchulgo") and (makerid <> "") then
	'// �ҷ� ��ǰ ��ǰ�� ��� üũ
	set opartner = new CPartnerUser
	opartner.FRectDesignerID = makerid
	opartner.GetOnePartnerNUser

	set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = opartner.FOneItem.FGroupid
	ogroup.GetOneGroupInfo

	if (ogroup.FOneItem.Fcompany_no = "211-87-00620") then
		' 8��1�ϱ��� �ӽ÷� ǰ
		if date()>="2022-08-01" then
		IsReturnOK = False
		end if
	end if
end if

if searchtype="bad" then
	' 8��1�ϱ��� �ӽ÷� ǰ
	if date()>="2022-08-01" then
	IscheckReturnOK = False
	end if
end if

dim BadOrErrText
if (searchtype="bad") then
    BadOrErrText = "�ҷ�"
else
    BadOrErrText = "�������"
end if

dim LorRText
if (actType="actloss") then
    LorRText = "�ν�"
elseif (actType="actshopchulgo") then
    LorRText = "�������"
else
    LorRText = "��ǰ"
end if

dim BasicMonth
BasicMonth = CStr(DateSerial(Year(now()),Month(now())-1,1))

%>
<script type='text/javascript'>

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
        if ((document.frm.itemid.value.length != 12) && (document.frm.itemid.value.length != 14)) {
                if (document.frm.makerid.value == "") {
                        alert("�귣�� �Ǵ� ��ǰ�ڵ带 �Է��ϼ���.");
                        return;
                }
                document.frm.itemid.value = "";
                document.frm.submit();
        } else {
                document.frm.makerid.value = "";
                document.frm.submit();
        }
}

function SubmitInsert(){
    <% if (osummarystock.FResultCount < 1) then %>
        alert("�˻��� ��ǰ�� �����ϴ�.");
        return;
    <% else %>
        if ((document.frm.itemid.value.length != 12) && (document.frm.itemid.value.length != 14)) {
			alert("��ǰ�ڵ带 ��Ȯ�� �Է��ϼ���.");
			return;
        }

		var frm = document.frm;
		var itembarcode = frm.itemid.value;

		var t_itemgubun = "";
		var t_itemid = "";
		var t_itemoption = "";

		for (var i = 0; ; i++) {
			var itemgubun = document.getElementById("itemgubun_" + i);
			var itemid = document.getElementById("itemid_" + i);
			var itemoption = document.getElementById("itemoption_" + i);

			var itemno = document.getElementById("itemno_" + i);
			var itemmaxno = document.getElementById("itemmaxno_" + i);

			if (itemgubun == undefined) {
				alert("��ǰ�� ��Ͽ� �����ϴ�. �ٸ� �귣���̰ų�, <%= BadOrErrText %>��ǰ ����� �Ǿ� ���� �ʽ��ϴ�.");
				break;
			}

			if (itembarcode.length == 12) {
				t_itemgubun = itembarcode.substring(0,2);
				t_itemid = (1 * itembarcode.substring(2,8));
				t_itemoption = itembarcode.substring(8);
			} else if (itembarcode.length == 14) {
				t_itemgubun = itembarcode.substring(0,2);
				t_itemid = (1 * itembarcode.substring(2,10));
				t_itemoption = itembarcode.substring(10);
			} else {
				alert("ERROR");
				return;
			}

			if ((itemgubun.value == t_itemgubun) && (itemid.value*1 == (1 * t_itemid)) && (itemoption.value == t_itemoption)) {
				itemno.value = (1 * itemno.value) + 1;

				<% if (searchtype = "bad") then %>
				if ((1 * itemno.value) > (itemmaxno.value * -1)) {
					itemno.value = (itemmaxno.value * -1);
					alert("�ҷ���ϵ� �������� ������ Ů�ϴ�. ���� �ҷ������ �ϼ���.");
				}
				<% end if %>

				hL(itemno);
				break;
			}
		}

		CalcTotalSelectedBuyPrice();

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
			break;
		}
		chk.checked = v.checked;
		SubmitCheckInsert(chk);
	}

	CalcTotalSelectedBuyPrice();
}

function CheckInsert(f, maxvalue){
        alert(f.value);
        if (f.value = "") {
                return;
        }

		<% if (searchtype = "bad") then %>
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
		<% end if %>

		CalcTotalSelectedBuyPrice();
}

function CalcTotalSelectedBuyPrice() {
    var frm = document.frm;
	var tot = 0;

	for (var i = 0; ; i++) {
		var buycash = document.getElementById("buycash_" + i);
		var itemno = document.getElementById("itemno_" + i);

		if (buycash == undefined) {
			break;
		}

		if (itemno.value*0 != 0) {
			return;
		}

		if (itemno.value*1 != 0) {
			tot = tot + buycash.value*1 * itemno.value*1;
		}
	}

	frm.totbuyprice.value = tot;
}

// �ſ� 5�ϱ��� �������� ��������
function checkAvail3(modiexecutedt) {
	var thisDate = "<%= Left(Now, 7) %>-01";
	var availDate = "<%= Left(Now, 7) %>-05";
	var nowdate = "<%= Left(now(),10) %>";
	var BasicMonth = "<%= BasicMonth %>";

	if (modiexecutedt < BasicMonth) {
		<% if Not C_ADMIN_AUTH then %>
			alert('����Ұ�\n\n������� �δ� ������¥�Դϴ�.');
			return false;
		<% else %>
			alert('[�����ڱ���]\n\n������� �δ� ������¥�Դϴ�.');
		<% end if %>
	}

	if (modiexecutedt < thisDate) {
		if (nowdate > availDate) {
			<% if Not C_ADMIN_AUTH then %>
				alert("����Ұ�\n\n�ſ� 5�ϱ����� �������� ��ϰ����մϴ�.");
				return false;
			<% else %>
				alert('[�����ڱ���]\n\n�ſ� 5�ϱ����� �������� ��ϰ����մϴ�.');
			<% end if %>
		}
	}

	return true;
}

function SubmitUpdateAll(){
    var pmwdiv = "";
    var onoffgubun = "";

    var frm = document.frm;

	if (checkAvail3(frm.yyyymmdd.value) != true) {
		return;
	}

    frm.itemgubunarr.value = "";
    frm.itemidarr.value = "";
    frm.itemoptionarr.value = "";
    frm.itemnoarr.value = "";

	//if ((frm.datetype[1].checked == true) && (frm.yyyymmdd.value.substring(0, 7) != (frm.yyyy1.value + "-" + frm.mm1.value))) {
	<% if datetype <> "curr" then %>
		if ((frm.yyyymmdd.value.substring(0, 7) != (frm.yyyy1.value + "-" + frm.mm1.value))) {
			alert("���ؿ��� ������ �ٸ��ϴ�.");
			return;
		}
	<% end if %>

	for (var i = 0; ; i++) {
		var itemgubun = document.getElementById("itemgubun_" + i);
		var itemid = document.getElementById("itemid_" + i);
		var itemoption = document.getElementById("itemoption_" + i);

		var itemno = document.getElementById("itemno_" + i);
		var mwdiv = document.getElementById("mwdiv_" + i);
		var centermwdiv = document.getElementById("centermwdiv_" + i);

		if (itemgubun == undefined) {
			break;
		}

		if ((mwdiv.value != "M") && (mwdiv.value != "W")) {
			mwdiv.value = centermwdiv.value;

			if ((mwdiv.value != "M") && (mwdiv.value != "W")) {
				alert("���Ա����� �����ʵǾ� �ִ� ��ǰ�� �ֽ��ϴ�.");
				return;
			}
		}

		if (itemno.value*1 != 0) {
			if (pmwdiv == "") {
				pmwdiv = mwdiv.value;
			} else {
				<% if (actType = "actreturn") then %>
				if (pmwdiv != mwdiv.value) {
					alert("��ǰ�� ���, ���� �Ӽ��� �ٸ���ǰ�� ���� ó�� �� �� �����ϴ�.");
					return;
				}
				<% end if %>
			}

			if (onoffgubun == "") {
				if (itemgubun.value == "10") {
					onoffgubun = "on";
				} else {
					onoffgubun = "off";
				}
			} else {
				if (((itemgubun.value == "10") && (onoffgubun != "on")) || ((itemgubun.value != "10") && (onoffgubun == "on"))) {
					alert("�¶��� �������� ��ǰ�� ���� ó�� �� �� �����ϴ�.");
					return;
				}
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

	<% if (actType="actloss") then %>
	if (frm.divcode.value == "") {
        alert("������� �����ϼ���.");
        frm.divcode.focus();
        return;
    }
	<% end if %>

	<% if (actType="actshopchulgo") then %>
	if (frm.chulgotargetid.value != "streetshop999") {
        alert("streetshop999 �� ���� �����մϴ�.");
        frm.chulgotargetid.focus();
        return;
    }
	<% end if %>

    if (confirm('<%= LorRText %> �������� �ۼ��Ͻðڽ��ϱ�?')){
        document.frm.method = "post";
        document.frm.pmwdiv.value = pmwdiv;
        document.frm.action = "/common/BadOrErrItem_re_input_process.asp";
        document.frm.submit();
    }
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'PopItemSellEdit','width=500,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopChulgoDate(comp) {
	calendarOpen(comp);
}

</script>

<!-- �˻� ���� -->
<form name="frm" method=get action="" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="brandid" value="<%= makerid %>">
<input type="hidden" name="searchtype" value="<%= searchtype %>">
<input type="hidden" name="actType" value="<%= actType %>">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="pmwdiv" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" >
<tr bgcolor="#FFFFFF">
    <td>** <%= BadOrErrText %>��ǰ <strong><%= LorRText %></strong> ó��</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" height="25" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			���ؿ� :
			<% ' �̹��� �̻�� ��û���� ���ܽ�Ŵ. Ư������ �������� ������ ���� ��� ���ذ� ������ �ƴ�?	' 2021.04.15 �ѿ�� %>
			<!--<input type="radio" name="datetype" value="curr" <% if (datetype = "curr") then %>checked<% end if %>> �������
			<input type="radio" name="datetype" value="yyyymm" <% if (datetype = "yyyymm") then %>checked<% end if %>> Ư���������� -->
			<% Call DrawYMBox(yyyy, mm) %>
			&nbsp;
			�귣��� : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="SubmitSearchItem()">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="25">
			<b>�귣�� ����</b>
			&nbsp;
			&nbsp;
			��뿩�� :
			<select class="select" name="makeruseyn">
				<option value="">-����-</option>n
				<option value="Y" <% if (makeruseyn = "Y") then %>selected<% end if %> >�����</option>
				<option value="N" <% if (makeruseyn = "N") then %>selected<% end if %> >������</option>
			</select>
			&nbsp;
			�������� : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="25">
			<b>��ǰ ����</b>
			&nbsp;
			&nbsp;
			��ǰ���� :
			<select class="select" name="itemgubun">
				<option value="">-����-</option><% '�̹��� �̻�� ��û���� ��ü �߰� %>
				<option value="10" <% if (itemgubun = "10") then %>selected<% end if %> >�»�ǰ(10)</option>
				<option value="OFF" <% if (itemgubun = "OFF") then %>selected<% end if %> >������ü</option>
				<option value="70" <% if (itemgubun = "70") then %>selected<% end if %> >����(70)</option>
				<option value="80" <% if (itemgubun = "80") then %>selected<% end if %> >����(80)</option>
				<option value="85" <% if (itemgubun = "85") then %>selected<% end if %> >����(85)</option>
				<option value="90" <% if (itemgubun = "90") then %>selected<% end if %> >����(90)</option>
			</select>
			&nbsp;
			&nbsp;
            <%'= CHKIIF(datetype<>"yyyymm", "ON���Ա���(����)", "<del>ON���Ա���(����)</del>") %>ON���Ա���(����) :
			<select class="select" name="mwdiv">
				<option value="">-����-</option>
				<option value="M" <% if (mwdiv = "M") then %>selected<% end if %> >����</option>
				<option value="W" <% if (mwdiv = "W") then %>selected<% end if %> >Ư��</option>
				<option value="U" <% if (mwdiv = "U") then %>selected<% end if %> >��ü</option>
				<option value="Z" <% if (mwdiv = "Z") then %>selected<% end if %> >������</option>
			</select>
			&nbsp;
            ���͸��Ա���(����) :
     		<select class="select" name="centermwdiv">
				<option value="">����</option>
				<option value="M" <%= CHKIIF(centermwdiv="M","selected","")%> >����</option>
				<option value="W" <%= CHKIIF(centermwdiv="W","selected","")%> >��Ź</option>
				<option value="X" <%= CHKIIF(centermwdiv="X","selected","")%> >������</option>
			</select>
     		&nbsp;
            <%'= CHKIIF(datetype="yyyymm", "���Ա���(����)", "<del>���Ա���(����)</del>") %>���Ա���(���) :
     		<select class="select" name="monthlymwdiv">
				<option value="">����</option>
				<option value="M" <%= CHKIIF(monthlymwdiv="M","selected","")%> >����</option>
				<option value="W" <%= CHKIIF(monthlymwdiv="W","selected","")%> >��Ź</option>
				<option value="X" <%= CHKIIF(monthlymwdiv="X","selected","")%> >������</option>
			</select>
			&nbsp;
			�Ǹſ��� :
			<select class="select" name="sellyn">
				<option value="">-����-</option>
				<option value="Y" <% if (sellyn = "Y") then %>selected<% end if %> >�Ǹ���</option>
				<option value="N" <% if (sellyn = "N") then %>selected<% end if %> >�Ǹž���</option>
			</select>
			&nbsp;
			��뿩�� :
			<select class="select" name="onlyisusing">
				<option value="">-����-</option>
				<option value="Y" <% if (onlyisusing = "Y") then %>selected<% end if %> >�����</option>
				<option value="N" <% if (onlyisusing = "N") then %>selected<% end if %> >������</option>
			</select>
		</td>
	</tr>
</table>

<br>

	<% if (IsReturnOK = False) then %><center><font color="red">��ǰ �Ұ� �귣��</font></center><% end if %>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td>
			��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="14" size="14" onKeyPress="if (event.keyCode == 13) { SubmitInsert(); return false; }">
			<input type="button" class="button" value="<%= LorRText %>�߰�" onclick="SubmitInsert()">
		</td>
		<td align="right">
			<% if (C_ADMIN_AUTH) and (FALSE) then %>
				<font color="red">[�����ں�]</font>
			<% end if %>
			����� : <input type="text" class="text_ro" name="yyyymmdd" value="<%= Left(Now(), 10) %>" size=11 readonly >
			<% if (C_ADMIN_AUTH) or (TRUE) then %>
				<a href="javascript:PopChulgoDate(frm.yyyymmdd);"><img src="/images/calicon.gif" align="absmiddle" border="0"></a>
			<% end if %>
			&nbsp;
			<% if (actType="actloss") or (actType="actshopchulgo") then %>
				���� :
				<% if (C_ADMIN_AUTH) or (TRUE) then %>
					<% Call drawSelectBoxIpChulDivcode("etclosschulgo", "divcode", "") %>
				<% else %>
					��Ÿ
					<input type="hidden" name="divcode" value="007">
				<% end if %>
				&nbsp;
				���ó :
				<% if (C_ADMIN_AUTH) or (TRUE) then %>
					<select class="select" name="chulgotargetid">
						<% if (searchtype="bad") and (actType="actloss") and (makerid <> "ithinkso") then %>
						<option value="itemdisuse" selected >itemdisuse</option>
						<% elseif (searchtype="err") and (actType="actloss") then %>
						<option value="itemloss">itemloss</option>
						<% end if %>
						<% if (searchtype="bad") and (actType="actshopchulgo") then %>
						<option value="streetshop999">streetshop999</option>
						<% end if %>
						<!--
						<option value="itemoutlet">itemoutlet</option>
						-->
						<option value="3pl_its_loss" <%=CHKIIF(makerid="ithinkso","selected","")%> >3pl_its_loss</option>
						<!-- <option value="itemstockmodify">itemstockmodify</option> -->
					</select>
				<% else %>
					itemloss
					<input type="hidden" name="chulgotargetid" value="itemloss">
				<% end if %>
				&nbsp;
			<% end if %>
			���԰��� : <input type="text" class="text_ro"  name="totbuyprice" size="10" value="0" readonly>
			&nbsp;
			<input type="button" class="button" value="��ü����" onclick="SubmitUpdateAll()" <% if (IsReturnOK = False) then %>disabled<% end if %> >
		</td>
	</tr>
</table>
</form>
<!-- �׼� �� -->

<% if (osummarystock.FResultCount>=osummarystock.FPageSize) then %>
�ִ� <%=osummarystock.FPageSize%> �� ǥ��
<% end if %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="20">
		<input type="checkbox" name="chkall" value="" onClick="SubmitCheckInsertAll(this);" <% if (IscheckReturnOK = False) then %>disabled<% end if %> >
	  </td>
      <td width="100">�귣��ID</td>
		<td width="40">���<br>����<br>����</td>
      <td width="40">ON<br>����<br>����</td>
	  <td width="40">����<br>����<br>����</td>
      <td width="25">����</td>
      <td width="40">��ǰ<br>�ڵ�</td>
      <td width="30">�ɼ�</td>
      <td>��ǰ��<br><font color="blue">[�ɼǸ�]</font></td>
      <td width="50">�Һ��ڰ�</td>
      <td width="50">���԰�</td>
      <td width="30">�Ǹ�<br>����</td>
      <td width="30">���<br>����</td>
      <td width="60"><%= BadOrErrText %><br>����</td>
      <td width="40"><%= LorRText %><br>����</td>
      <td width="60">�ǻ����</td>
    </tr>
    <form name="frmlist" method=get action="" onsubmit="return false;">
<% for i=0 to osummarystock.FResultCount - 1 %>

	<% if (osummarystock.FItemList(i).Fisusing = "Y") then %>
		<% if (osummarystock.FItemList(i).FItemgubun = "10") then %>
			<tr align="center" bgcolor="#FFFFFF">
		<% else %>
			<tr align="center" bgcolor="#EEEEEE">
		<% end if %>
	<% else %>
		<tr align="center" bgcolor="#BBBBBB">
	<% end if %>

		<td><input type="checkbox" id="chk_<%= i %>" name="chk" value="<%= i %>" onClick="SubmitCheckInsert(this); CalcTotalSelectedBuyPrice();" <% if (IscheckReturnOK = False) then %>disabled<% end if %> ></td>
		<td><%= osummarystock.FItemList(i).Fmakerid %></td>
		<td align="center" style="color:<%=GetMwDivColorCd(osummarystock.FItemList(i).flastmwdiv)%>;"><%= osummarystock.FItemList(i).flastmwdiv %></td>
		<td align="center" style="color:<%=osummarystock.FItemList(i).GetMwDivColor%>;"><%= osummarystock.FItemList(i).Fmwdiv %></td>
		<td align="center" style="color:<%=GetMwDivColorCd(osummarystock.FItemList(i).Fcentermwdiv)%>;"><%= osummarystock.FItemList(i).Fcentermwdiv %></td>
		<td><%= osummarystock.FItemList(i).FItemgubun %></td>
		<td><a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemid %>');"><%= osummarystock.FItemList(i).FItemid %></a></td>
		<td><%= osummarystock.FItemList(i).FItemoption %></td>
		<td align="left"><a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= osummarystock.FItemList(i).FItemgubun %>&itemid=<%= osummarystock.FItemList(i).FItemid %>&itemoption=<%= osummarystock.FItemList(i).FItemoption %>" target=_blank ><%= osummarystock.FItemList(i).FItemname %></a><br><font color="blue">[<%= osummarystock.FItemList(i).FItemOptionName %>]</font></td>
		<td align="right"><%= formatnumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
		<td align="right"><%= formatnumber(osummarystock.FItemList(i).Fbuycash,0) %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fsellyn %></td>
		<td align="center"><%= osummarystock.FItemList(i).Fisusing %></td>
		<td>
		<%= osummarystock.FItemList(i).Fregitemno %>
		</td>
      <input type="hidden" id="itemgubun_<%= i %>" name="itemgubun" value="<%= osummarystock.FItemList(i).FItemgubun %>">
      <input type="hidden" id="itemid_<%= i %>" name="itemid" value="<%= osummarystock.FItemList(i).FItemid %>">
      <input type="hidden" id="itemoption_<%= i %>" name="itemoption" value="<%= osummarystock.FItemList(i).FItemOption %>">
      <td>
        <input type="text" class="text" id="itemno_<%= i %>" name="itemno" value="0" size="3" onKeyUP="checkhL(this);  CalcTotalSelectedBuyPrice();">
      </td>
      <td align="center"><%= FormatNumber(osummarystock.FItemList(i).Frealstock, 0) %></td>
      <input type="hidden" id="itemmaxno_<%= i %>" name="itemmaxno" value="<%= osummarystock.FItemList(i).Fregitemno %>" >
      <input type="hidden" id="mwdiv_<%= i %>" name="mwdiv" value="<%= osummarystock.FItemList(i).FMwdiv %>">
	  <input type="hidden" id="centermwdiv_<%= i %>" name="centermwdiv" value="<%= osummarystock.FItemList(i).Fcentermwdiv %>">
      <input type="hidden" id="buycash_<%= i %>" name="buycash" value="<%= osummarystock.FItemList(i).Fbuycash %>">
    </tr>
   	<% next %>
<% if osummarystock.FResultCount = 0 then %>
    <tr align="center" bgcolor="#FFFFFF">
      <td colspan="20" align="center">�˻��� ��ǰ�� �����ϴ�.</td>
    </tr>
<% end if %>
    </form>
</table>

<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
