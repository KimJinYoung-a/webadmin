<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� �����������Ʈ
' Hieditor : 2009.04.07 ������ ����
'			 2011.04.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False
dim EditEnabled , yyyymmdd,yyyy1,mm1,dd1 ,i ,oipchulmaster, oipchul
dim PriceEditEnable ,idx ,DispReqNo ,edityn
	idx = requestCheckVar(request("idx"),10)

edityn = FALSE

if idx = "" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('idx ���� �����ϴ�');"
	response.write "	history.back();"
	response.write "</script>"
	dbget.close() : response.end
end if

'if Not (C_IS_SHOP) and Not (C_IS_Maker_Upche) then PriceEditEnable = true

set oipchulmaster = new CShopIpChul
	oipchulmaster.FRectIdx = idx
	oipchulmaster.GetIpChulMasterList

set oipchul = new CShopIpChul
	oipchul.FRectIdx = idx
	oipchul.GetIpChulDetail

if oipchulmaster.ftotalcount < 1 then
	response.write "<script type='text/javascript'>"
	response.write "	alert('�ش�Ǵ� ��������� �����ϴ�');"
	response.write "	history.back();"
	response.write "</script>"
	dbget.close() : response.end
end if

if C_ADMIN_USER or C_IS_OWN_SHOP then
	edityn = TRUE

'//�����ϰ�� ���������� ���θ��常
elseif (C_IS_SHOP) then
	IS_HIDE_BUYCASH = True

	if C_STREETSHOPID = oipchulmaster.FItemList(0).FShopid then
		edityn = TRUE
	else
		edityn = FALSE
	end if
else
	edityn = TRUE
end if

yyyymmdd = Left(CStr(oipchulmaster.FItemList(0).FScheduleDt),10)
yyyy1 = left(yyyymmdd,4)
mm1 = mid(yyyymmdd,6,2)
dd1 = mid(yyyymmdd,9,2)

''�԰��û ������ ��� Ȯ������ ����.
if (C_IS_Maker_Upche) and (oipchulmaster.FItemList(0).IsRequireConfirm) then
    oipchulmaster.FItemList(0).UpcheConfirmProcess
    response.write "<script type='text/javascript'>alert('�԰��û Ȯ�εǾ����ϴ�. ������ Ȯ���� �߼� ó�� �� �ּ���.');</script>"
end if

''�԰��� ���� && �ڱⰡ ����ѳ����� ���� ����
EditEnabled = oipchulmaster.FItemList(0).IsEditEnabled
PriceEditEnable = oipchulmaster.FItemList(0).IsPriceEditEnabled
DispReqNo   = oipchulmaster.FItemList(0).IsDispReqNo

'/����, ������ ��� ���� ���� ����
if C_ADMIN_USER or C_IS_SHOP then EditEnabled = true

%>

<script type='text/javascript'>

<% if (EditEnabled) then %>
	var ipgowait = true;
<% else %>
	var ipgowait = false;
<% end if %>

function ReActItems(igubun,iitemid,iitemoption,isellcash,isuplycash,ishopbuyprice,iitemno,iitemname,iitemoptionname,iitemdesigner){
	frmArrupdate.itemgubunarr.value = igubun;
	frmArrupdate.itemarr.value = iitemid;
	frmArrupdate.itemoptionarr.value = iitemoption;
	frmArrupdate.sellcasharr.value = isellcash;
	frmArrupdate.suplycasharr.value = isuplycash;
	frmArrupdate.shopbuypricearr.value = ishopbuyprice
	frmArrupdate.itemnoarr.value = iitemno;
	frmArrupdate.designerarr.value = iitemdesigner;
	frmArrupdate.submit();
}

function ReAct(){
	location.reload();
}

function UpcheChulgoProc(frm){
    if (frm.songjangdiv.value.length<1){
		alert('�ù�縦 ���� �ϼ���');
		frm.songjangdiv.focus();
		return;
	}


	if (frm.songjangno.value.length<1){
		alert('���� ��ȣ�� �Է� �ϼ���');
		frm.songjangno.focus();
		return;
	}

    var ret= confirm('�߼� ó�� �Ͻðڽ��ϱ�?');
	if (ret){
	    frm.mode.value = "upchechulgoproc";
		frm.submit();
	}
}


function ModiMaster(frm,scd){
	if (!ipgowait){
		alert('�԰��� ���°� �ƴϸ� ������ �� �����ϴ�.');
		return;
	}

	if (frm.chargeid.value.length<1){
		alert('����ó�� �����ϼ���.');
		return;
	}

	if (frm.shopid.value.length<1){
		alert('��ID�� �����ϼ���.');
		return;
	}

	var ret= confirm('���� �Ͻðڽ��ϱ�?');
	if (ret){
		if(scd != "")
		{
			frm.statecd.value = scd;
		}
		frm.submit();
	}
}

function AddItems(chargeid, idx){
	if (!ipgowait){
		alert('�԰��� ���°� �ƴϸ� ������ �� �����ϴ�.');
		return;
	}

	var popwin;
	popwin = window.open('popshopitem2.asp?shopid=' + frmMaster.shopid.value + '&chargeid=' + chargeid + '&idx=' + idx,'shopitem','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ModiDetail(frm){
	if (!ipgowait){
		<% if Not(C_ADMIN_AUTH) then %>
			alert('�԰��� ���°� �ƴϸ� ������ �� �����ϴ�.');
			return;
		<% else %>
			alert("������ ����!!");
		<% end if %>
	}

	if (!IsDigit(frm.sellcash.value)){
		alert('�ǸŰ��� ���ڷ� �Է��ϼ���.');
		frm.sellcash.focus();
		return;
	}

	if (frm.suplycash.value*0 != 0) {
		alert('���ް��� ���ڷ� �Է��ϼ���.');
		frm.suplycash.focus();
		return;
	}

	if (!IsInteger(frm.itemno.value)){
		alert('������ ������ �Է��ϼ���.');
		frm.itemno.focus();
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="detailmodi";
		frm.submit();
	}
}

<% if (idx <> "") and (edityn = True) or (C_ADMIN_AUTH) then %>
	function ModiDetailArr() {
		var frm;

		var mode = "detailmodiarr";
		var midx = <%= idx %>;

		var didxarr = "";
		var sellcasharr = "";
		var suplycasharr = "";
		var shopbuypricearr = "";
		var itemnoarr = "";

		for (var i = 0;i < document.forms.length; i++) {
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (!IsDigit(frm.sellcash.value)) {
					alert('�ǸŰ��� ���ڷ� �Է��ϼ���.');
					frm.sellcash.focus();
					return;
				}

				if (frm.suplycash.value*0 != 0) {
					alert('���ް��� ���ڷ� �Է��ϼ���.');
					frm.suplycash.focus();
					return;
				}

				if (!IsInteger(frm.itemno.value)) {
					alert('������ ������ �Է��ϼ���.');
					frm.itemno.focus();
					return;
				}

				didxarr = didxarr + "|" + frm.idx.value;
				sellcasharr = sellcasharr + "|" + frm.sellcash.value;
				suplycasharr = suplycasharr + "|" + frm.suplycash.value;
				shopbuypricearr = shopbuypricearr + "|" + frm.shopbuyprice.value;
				itemnoarr = itemnoarr + "|" + frm.itemno.value;
			}
		}

		if (confirm('���� �Ͻðڽ��ϱ�?')) {
			frm = document.frmArrupdate;

			frm.mode.value = mode;
			frm.midx.value = midx;
			frm.didxarr.value = didxarr;
			frm.sellcasharr.value = sellcasharr;
			frm.suplycasharr.value = suplycasharr;
			frm.shopbuypricearr.value = shopbuypricearr;
			frm.itemnoarr.value = itemnoarr;

			frm.action = "do_shopipchul.asp";

			frm.submit();
		}
	}
<% end if %>

function DelDetail(frm){
	if (!ipgowait){
		alert('�԰��� ���°� �ƴϸ� ������ �� �����ϴ�.');
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="detaildel";
		frm.submit();
	}
}
function AddItemsBarCode(frm, digitflag){
	if (frm.shopid.value.length<1){
		alert('�������� ���� �����ϼ���');
		frm.shopid.focus();
		return;
	}

	var popwin;
	popwin = window.open('popshopitemBybarcode.asp?shopid=' + frmMaster.shopid.value + '&chargeid=' + frmMaster.chargeid.value + '&digitflag=' + digitflag,'popshopitemBybarcode','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>

<!-- ���Ŀ� �޴�����κп� �־�� �մϴ�. -->
<table width="100%" border="0" valign="top" cellpadding="0" cellspacing="0" class="a">
<tr bgcolor="#FFFFFF">
	<td style="padding:5; border:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>���� ���� ����� ����</strong></font><br>
		* �԰� Ȯ���� ���� ���� 1�ÿ� ��� �ݿ��˴ϴ�.<br>
		* ��ǰ�� ������ ���̳ʽ��� ����ּ���
	</td>
</tr>
</table>
<!-- ���Ŀ� �޴�����κп� �־�� �մϴ�. -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="do_shopipchul.asp">
<input type="hidden" name="mode" value="modimaster">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="divcode" value="006">
<input type="hidden" name="vatcode" value="008">
<input type="hidden" name="statecd" value="">
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">����ó</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="chargeid" value="<%= oipchulmaster.FItemList(0).FChargeid %>">
		<%= oipchulmaster.FItemList(0).FChargeid %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">���� </td>
	<td bgcolor="#FFFFFF">
	<input type="hidden" name="shopid" value="<%= oipchulmaster.FItemList(0).FShopid %>">
		<%= oipchulmaster.FItemList(0).FShopid %> (<%= oipchulmaster.FItemList(0).FShopname %>)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">���ǸŰ�</td>
	<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalSellCash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�Ѱ��ް�</td>
	<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalSuplyCash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�԰�����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="scheduledt" value="<%= oipchulmaster.FItemList(0).FScheduleDt %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.scheduledt);">
		<img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>

		&nbsp;
		�ù��:<% drawSelectBoxDeliverCompany "songjangdiv", oipchulmaster.FItemList(0).Fsongjangdiv %>
		&nbsp;
		�����ȣ:<input type="text" class="text" name="songjangno" size=14 maxlength=16 value="<%= oipchulmaster.FItemList(0).Fsongjangno %>" >

		<% IF (C_IS_Maker_Upche) and (oipchulmaster.FItemList(0).FisbaljuExists="Y") and (oipchulmaster.FItemList(0).Fstatecd=-1) then %>
		    <input type="button" class="button" value="�߼�ó��" onClick="UpcheChulgoProc(frmMaster);">
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�԰���</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).FexecDt %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">����Ȯ����</td>
	<td bgcolor="#FFFFFF">
		<%= oipchulmaster.FItemList(0).Fshopconfirmdate %>
		<% if Not IsNULL(oipchulmaster.FItemList(0).Fshopconfirmuserid) then %>
			(Ȯ��ID : <%= oipchulmaster.FItemList(0).Fshopconfirmuserid %>)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">��üȮ����</td>
	<td bgcolor="#FFFFFF">
		<%= oipchulmaster.FItemList(0).Fupcheconfirmdate %>
		<% if Not IsNULL(oipchulmaster.FItemList(0).Fupcheconfirmuserid) then %>
			(Ȯ��ID : <%= oipchulmaster.FItemList(0).Fupcheconfirmuserid %>)
		<% end if %>
	</td>
</tr>
<% if Not IsNULL(oipchulmaster.FItemList(0).Fbaljuconfirmdate) then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�԰��ûȮ����</td>
	<td bgcolor="#FFFFFF">
		<%= oipchulmaster.FItemList(0).Fbaljuconfirmdate %>
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
	<td bgcolor="#FFFFFF">
		<%= oipchulmaster.FItemList(0).FRegDate %>
		<% if Not IsNULL(oipchulmaster.FItemList(0).Freguserid) then %>
			(���ID : <%= oipchulmaster.FItemList(0).Freguserid %>)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF"><font color="<%= oipchulmaster.FItemList(0).getStateColor %>"><%= oipchulmaster.FItemList(0).getStateName %></font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ��û����</td>
	<td>
		<textarea name="comment" class="textarea" cols="80" rows="6"><%= oipchulmaster.FItemList(0).fcomment %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" colspan="2" align="center">
	<% if (C_IS_Maker_Upche) and (oipchulmaster.FItemList(0).FStatecd<0) then %>

	<% else %>
		<% if not(edityn) or not(EditEnabled) then %>

	    <% else %>
	    	<input type="button" value="����" onClick="ModiMaster(frmMaster,'')" class="button">
	    <% end if %>
	<% end if %>
	<%
	'//���忡�� �Է��� ��ü�� ���ֿ�û
	if oipchulmaster.FItemList(i).FisbaljuExists="Y" then
	%>
		<% if oipchulmaster.FItemList(0).FStatecd = -5 then %>
			&nbsp;<input type="button" value="�԰��û���κ���" onClick="ModiMaster(frmMaster,'-2')" class="button">
		<% end if %>
	<%
	else
	%>
		<% if oipchulmaster.FItemList(0).FStatecd = -5 then %>
			&nbsp;<input type="button" value="�԰���κ���" onClick="ModiMaster(frmMaster,'0')" class="button">
		<% end if %>
	<% end if %>

	</td>
</tr>
</form>
</table>

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="��ǰ�߰�" onclick="AddItems('<%= oipchulmaster.FItemList(0).FChargeid %>','<%= oipchulmaster.FItemList(0).FIdx %>')" <% if not EditEnabled then response.write "disabled" %>>

		<%' If C_IS_SHOP or C_ADMIN_AUTH or C_OFF_AUTH or C_logics_Part then %>
			<input type="button" class="button" value="����(���ڵ�)" onclick="AddItemsBarCode(frmMaster,'P')">
			<input type="button" class="button" value="��ǰ(���ڵ�)" onclick="AddItemsBarCode(frmMaster,'M')">
		<%' End If %>
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oipchul.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oipchul.FTotalCount %></b>
	</td>
</tr>
<% end if %>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">���ڵ�</td>
	<td width="80">�귣��ID</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td width="50">�ǸŰ�</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	    <td width="60">�ٹ�����<br>���԰�</td>
	    <td width="60">����<br>���ް�</td>
	<% elseif (C_IS_Maker_Upche) then %>
		<td width="60">�ٹ�����<br>���ް�</td>
	<% else %>
		<td width="60">����<br>���ް�</td>
	<% end if %>

	<td width="50">����</td>

	<% if (DispReqNo) then %>
		<td width="50">��û<br>����</td>
	<% end if %>

	<td width="60">�ǸŰ��հ�</td>
	<td width="40">����</td>
	<td width="40">����</td>
</tr>
<% for i=0 to oipchul.FResultCount-1 %>
<form name="frmBuyPrc_<%= i %>" method="post" action="do_shopipchul.asp">
<input type="hidden" name="chargeid" value="<%= oipchulmaster.FItemList(0).FChargeid %>">
<input type="hidden" name="shopid" value="<%= oipchulmaster.FItemList(0).FShopid %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="midx" value="<%= idx %>">
<input type="hidden" name="idx" value="<%= oipchul.FItemList(i).FIdx %>">

<% if Not PriceEditEnable then %>
	<input type="hidden" name="sellcash" value="<%= oipchul.FItemList(i).FSellCash %>">
	<% if IS_HIDE_BUYCASH = True then %>
	<input type="hidden" name="suplycash" value="-1">
	<% else %>
	<input type="hidden" name="suplycash" value="<%= oipchul.FItemList(i).FSuplyCash %>">
	<% end if %>
	<input type="hidden" name="shopbuyprice" value="<%= oipchul.FItemList(i).Fshopbuyprice %>">
<% end if %>

<tr align="center" bgcolor="#FFFFFF">
	<td><%= oipchul.FItemList(i).GetBarCode %></td>
	<td>
		<%= oipchul.FItemList(i).Fdesignerid %>
		<% if (C_ADMIN_AUTH) then %>
		    <% if (oipchul.FItemList(i).Fdesignerid<>oipchul.FItemList(i).FCurrMakerid) then %>
		    <br>(<%=oipchul.FItemList(i).FCurrMakerid%>)
		    <% end if %>
	    <% end if %>
	</td>
	<td align="left"><%= oipchul.FItemList(i).FItemName %></td>
	<td><%= oipchul.FItemList(i).FItemOptionName %></td>

	<% if Not (PriceEditEnable) then %>
		<td align="right"><%= FormatNumber(oipchul.FItemList(i).FSellCash,0) %></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><%= FormatNumber(oipchul.FItemList(i).FSuplyCash,0) %></td><!--�ٹ����� ���԰�-->
			<td align="right"><%= FormatNumber(oipchul.FItemList(i).Fshopbuyprice,0) %></td><!--���� ���ް�-->
		<% elseif (C_IS_Maker_Upche) then %>
			<td align="right"><%= FormatNumber(oipchul.FItemList(i).FSuplyCash,0) %></td><!--�ٹ����� ���ް�-->
		<% else %>
			<td align="right"><%= FormatNumber(oipchul.FItemList(i).Fshopbuyprice,0) %></td><!--���� ���ް�-->
		<% end if %>
	<% else %>
		<td align="right"><input type="text" name="sellcash" value="<%= oipchul.FItemList(i).FSellCash %>" size="7" maxlength="9"></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right">
				<input type="text" name="suplycash" value="<%= oipchul.FItemList(i).FSuplyCash %>" size="7" maxlength="9"><!--�ٹ����� ���԰�-->
				<% if oipchul.FItemList(i).FSellCash<>0 then response.write FormatNumber(100-CLNG(oipchul.FItemList(i).FSuplyCash/oipchul.FItemList(i).FSellCash*100*100)/100,2) end if %>
			</td>
			<td align="right">
				<input type="text" name="shopbuyprice" value="<%= oipchul.FItemList(i).Fshopbuyprice %>" size="7" maxlength="9"><!--���� ���ް�-->
				<% if oipchul.FItemList(i).FSellCash<>0 then response.write FormatNumber(100-CLNG(oipchul.FItemList(i).Fshopbuyprice/oipchul.FItemList(i).FSellCash*100*100)/100,2) end if %>
			</td>
		<% elseif (C_IS_Maker_Upche) then %>
			<td align="right">
				<input type="text" name="suplycash" value="<%= oipchul.FItemList(i).FSuplyCash %>" size="7" maxlength="9"><!--�ٹ����� ���ް�-->
				<% if oipchul.FItemList(i).FSellCash<>0 then response.write FormatNumber(100-CLNG(oipchul.FItemList(i).FSuplyCash/oipchul.FItemList(i).FSellCash*100*100)/100,2) end if %>
			</td>
		<% else %>
			<td align="right">
				<input type="text" name="shopbuyprice" value="<%= oipchul.FItemList(i).Fshopbuyprice %>" size="7" maxlength="9"><!--���� ���ް�-->
				<% if oipchul.FItemList(i).FSellCash<>0 then response.write FormatNumber(100-CLNG(oipchul.FItemList(i).Fshopbuyprice/oipchul.FItemList(i).FSellCash*100*100)/100,2) end if %>
			</td>
		<% end if %>
	<% end if %>

	<td><input type="text" class="text" name="itemno" value="<%= oipchul.FItemList(i).Fitemno %>" size="3" maxlength="4"></td>

	<% if (DispReqNo) then %>
		<td><%= oipchul.FItemList(i).Freqno %></td>
	<% end if %>

	<td align="right">
		<%= ForMatNumber(oipchul.FItemList(i).Fitemno*oipchul.FItemList(i).FSellCash,0) %>
	</td>
	<td>
		<input type="button" class="button" value="����" <% if not(edityn) and Not(C_ADMIN_AUTH) then response.write " disabled" %> onclick="ModiDetail(frmBuyPrc_<%= i %>)" <% if not EditEnabled and Not(C_ADMIN_AUTH) then response.write "disabled" %>>
	</td>
	<td>
		<input type="button" class="button" value="����" <% if not(edityn) then response.write " disabled" %> onclick="DelDetail(frmBuyPrc_<%= i %>)" <% if not EditEnabled then response.write "disabled" %>>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="12" align=center>
		<% if (idx <> "") and (edityn = True) or (C_ADMIN_AUTH) then %>
			<input type="button" class="button" value=" ��ü���� " onclick="ModiDetailArr()" >
		<% end if %>
	</td>
</tr>
</table>

<form name="frmArrupdate" method="post" action="shopipchulitem_process.asp">
	<input type="hidden" name="mode" value="arrins">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="midx" value="">
	<input type="hidden" name="didxarr" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="shopbuypricearr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="designerarr" value="">
	<input type="hidden" name="chargeid" value="<%= oipchulmaster.FItemList(0).FChargeid %>">
	<input type="hidden" name="shopid" value="<%= oipchulmaster.FItemList(0).FShopid %>">
</form>

<%
set oipchulmaster = Nothing
set oipchul = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
