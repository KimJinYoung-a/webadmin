<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
''��ǥ �� ���̵� : ��������ǥ or ��������ǥ or �ؿ�����ǥ�� ����.
dim ProtoShopid
ProtoShopid = "streetshop000"


dim idx, isfixed, opage, ourl,oshopid,ostatecd,odesinger

idx     = RequestCheckVar(request("idx"),9)
opage   = RequestCheckVar(request("opage"),9)
ourl    = RequestCheckVar(request("ourl"),128)
oshopid = RequestCheckVar(request("oshopid"),32)
ostatecd    = RequestCheckVar(request("ostatecd"),32)
odesinger   = RequestCheckVar(request("odesinger"),32)

if idx="" then idx=0

dim ojumunmaster, ojumundetail, oupchemwinfo

set ojumunmaster = new COrderSheet
ojumunmaster.FRectIdx = idx
ojumunmaster.GetOneOrderSheetMaster


set ojumundetail= new COrderSheet
ojumundetail.FRectIdx = idx
ojumundetail.FRectShopid = ProtoShopid '''ojumunmaster.FoneItem.FBaljuid
ojumundetail.GetOrderSheetDetail


set oupchemwinfo = new CUpcheMwInfo
oupchemwinfo.FRectdesignerId = ojumunmaster.FOneItem.Ftargetid
oupchemwinfo.GetDesignerMWInfo


dim yyyymmdd
yyyymmdd = Left(ojumunmaster.FOneItem.Fscheduledate,10)


if (ojumunmaster.FOneItem.FStatecd>7) then
	isfixed = true
else
	isfixed = false
end if


''�⺻ ���� ����
dim DIFFCenterMWDivExists
DIFFCenterMWDivExists = False

dim DefaultItemMwDiv
DefaultItemMwDiv = GetDefaultItemMwdivByBrand(odesinger)


''��ǥ �� ���̵� ���� - ���� ������������.
dim sqlStr
sqlStr = " select top 1 s.shopid, s.defaultmargin"
sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_designer s,"
sqlStr = sqlStr & " db_shop.dbo.tbl_shop_user u"
sqlStr = sqlStr & " where s.shopid=u.userid"
sqlStr = sqlStr & " and s.makerid='best_ever'"
sqlStr = sqlStr & " and u.shopdiv in ('2','4','6')"
sqlStr = sqlStr & " order by s.defaultmargin desc, u.shopdiv"

rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
    ProtoShopid = rsget("shopid")
else
    response.write "<script>alert('������ ���� �Ǿ� ���� �ʽ��ϴ�. ������ ���� ���');</script>"
end if
rsget.Close

dim tmpcolor

%>
<script language='javascript'>
function popOffItemEdit(ibarcode){
	var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}

<% if ojumunmaster.FOneItem.FStatecd="0" then %>
var jumunwait = true;
<% else %>
var jumunwait = false;
<% end if %>

function DelArr(){
	var upfrm = document.frmadd;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� ������ �����ϴ�.');
		return;
	}

	upfrm.detailidxarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				upfrm.detailidxarr.value = upfrm.detailidxarr.value + frm.detailidx.value + ",";
			}
		}
	}

	if (confirm('���� ������ ���� �Ͻðڽ��ϱ�?')){
		upfrm.mode.value = "delshopjumunarr";
		upfrm.submit();
	}
}

function SaveArr(){
	var upfrm = document.frmadd;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.baljuitemnoarr.value = "";
	upfrm.realitemnoarr.value = "";
	upfrm.commentarr.value = "";
	upfrm.detailidxarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsInteger(frm.baljuitemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.baljuitemno.focus();
					return;
				}

				if (frm.baljuitemno.value.length<1){
					alert('������ �Է��ϼ���.');
					frm.baljuitemno.focus();
					return;
				}

				if (!IsInteger(frm.realitemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.realitemno.focus();
					return;
				}

				if (frm.realitemno.value.length<1){
					alert('������ �Է��ϼ���.');
					frm.realitemno.focus();
					return;
				}

				upfrm.detailidxarr.value = upfrm.detailidxarr.value + frm.detailidx.value + "|";
				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
				upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
				upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
				upfrm.baljuitemnoarr.value = upfrm.baljuitemnoarr.value + frm.baljuitemno.value + "|";
				upfrm.realitemnoarr.value = upfrm.realitemnoarr.value + frm.realitemno.value + "|";
				upfrm.commentarr.value = upfrm.commentarr.value + frm.comment.value + "|";
			}
		}
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		upfrm.mode.value = "modeshopjumunarr";
		upfrm.submit();
	}
}

function SaveALL(){
	var masterfrm = document.frmMaster;
	var upfrm = document.frmadd;
	var frm;
	var pass = false;



	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.baljuitemnoarr.value = "";
	upfrm.realitemnoarr.value = "";
	upfrm.commentarr.value = "";
	upfrm.detailidxarr.value = "";

	upfrm.ipgoflagarr.value = "";
	upfrm.defaultmaginflagarr.value = "";
	upfrm.buymaginflagarr.value = "";
	upfrm.suplymaginflagarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

				if (!IsInteger(frm.baljuitemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.baljuitemno.focus();
					return;
				}

				if (frm.baljuitemno.value.length<1){
					alert('������ �Է��ϼ���.');
					frm.baljuitemno.focus();
					return;
				}

				if (!IsInteger(frm.realitemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.realitemno.focus();
					return;
				}

				if (frm.realitemno.value.length<1){
					alert('������ �Է��ϼ���.');
					frm.realitemno.focus();
					return;
				}

				upfrm.detailidxarr.value = upfrm.detailidxarr.value + frm.detailidx.value + "|";
				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
				upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
				upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
				upfrm.baljuitemnoarr.value = upfrm.baljuitemnoarr.value + frm.baljuitemno.value + "|";
				upfrm.realitemnoarr.value = upfrm.realitemnoarr.value + frm.realitemno.value + "|";
				upfrm.commentarr.value = upfrm.commentarr.value + frm.comment.value + "|";

				//if (frm.ipgoflag.checked){
					upfrm.ipgoflagarr.value = upfrm.ipgoflagarr.value + frm.ipgoflag.value + "|";
				//}else{
				//	upfrm.ipgoflagarr.value = upfrm.ipgoflagarr.value + "|";
				//}

				upfrm.defaultmaginflagarr.value = upfrm.defaultmaginflagarr.value + frm.defaultmaginflag.value + "|";
				upfrm.buymaginflagarr.value = upfrm.buymaginflagarr.value + frm.buymaginflag.value + "|";
				upfrm.suplymaginflagarr.value = upfrm.suplymaginflagarr.value + frm.suplymaginflag.value + "|";
		}
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		if (masterfrm.beasongdate!=undefined){
			upfrm.songjangname.value = masterfrm.songjangdiv.options[masterfrm.songjangdiv.selectedIndex].text;
			upfrm.beasongdate.value = masterfrm.beasongdate.value;
			upfrm.songjangdiv.value = masterfrm.songjangdiv.value;
			upfrm.songjangno.value = masterfrm.songjangno.value;
		}
		upfrm.yyyymmdd.value = masterfrm.yyyymmdd.value;
		upfrm.comment.value = masterfrm.comment.value;

		upfrm.statecd.value = getCheckboxValue(masterfrm,'statecd');
		upfrm.divcode.value = getCheckboxValue(masterfrm,'divcode');
		upfrm.mode.value = "modeshopjumunmasterdetail";
		upfrm.submit();
	}
}

function getCheckboxValue(f,compname){
    for(var i=0;i<f.elements.length;i++){
      if(f.elements[i].name==compname && f.elements[i].checked){
        return f.elements[i].value;
      }
    }
    return false;
}


function AddItems(frm){
	//if (jumunwait!=true){
	//	alert('�ֹ����� ���°� �ƴϸ� �����Ͻ� �� �����ϴ�.');
	//	return;
	//}

	var popwin;
	var suplyer, shopid;

	if (frm.suplyer.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid  = frm.shopid.value;
	popwin = window.open('/common/offshop/popshopjumunitem.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=' + frm.masteridx.value ,'upchejumuninputadd','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ModiThis(frm){
	//if (jumunwait==true){
	//	alert('�ֹ����� ���¿��� �����Ͻ� �� �����ϴ�.');
	//	return;
	//}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}

function DelThis(frm){
	//if (jumunwait!=true){
	//	alert('�ֹ����� ���°� �ƴϸ� �����Ͻ� �� �����ϴ�.');
	//	return;
	//}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="deldetail";
		frm.submit();
	}
}

function DelMaster(frm){
	//if (jumunwait!=true){
	//	alert('�ֹ����� ���°� �ƴϸ� �����Ͻ� �� �����ϴ�.');
	//	return;
	//}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="delmaster";
		frm.submit();
	}
}

function ModiMaster(frm){
	if (frm.beasongdate!=undefined){
		frm.songjangname.value = frm.songjangdiv.options[frm.songjangdiv.selectedIndex].text;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="modimaster";
		frm.submit();
	}
}

function ReActItems(iidx,igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner){
	if (iidx!='<%= idx %>'){
		alert('�ֹ����� ��ġ���� �ʽ��ϴ�. �ٽýõ��� �ּ���.');
		return;
	}

	frmadd.itemgubunarr.value = igubun;
	frmadd.itemarr.value = iitemid;
	frmadd.itemoptionarr.value = iitemoption;
	frmadd.sellcasharr.value = isellcash;
	frmadd.suplycasharr.value = isuplycash;
	frmadd.buycasharr.value = ibuycash;
	frmadd.itemnoarr.value = iitemno;

	frmadd.submit();
}

function ChulgoProc(frm){
	if (frm.ipgodate.value.length<1){
		alert('������� �Է��� �ּ���.');
		frm.ipgodate.focus();
		if (!calendarOpen2(frm.ipgodate)) { return };
	}

	if (frm.beasongdate!=undefined){
		frm.songjangname.value = frm.songjangdiv.options[frm.songjangdiv.selectedIndex].text;
	}

	var ret = confirm('���ó�� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="chulgoproc";
		frm.submit();
	}
}

function showSpecialInput(objTarget){
	if(objTarget[objTarget.selectedIndex].id=='special'){
	 	output = window.showModalDialog("/lib/inputpop.html" , null, "dialogwidth:250px;dialogheight:120px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	 	if(output!=''){
	 		objTarget[objTarget.selectedIndex].text=output;
	  		objTarget[objTarget.selectedIndex].value=output;
	 	}else{

	 	}
	 }
}

function IpgoFinish(){
	var imsg = "";

	if (frmMaster.ipgodate.value.length<1){
		var ret1 = calendarOpen2(frmMaster.ipgodate);
		if (!ret1) return;
	}

	var ret2 = confirm('�԰��� : ' + frmMaster.ipgodate.value + ' OK?');
	if (!ret2) return;

	var idivcode = getCheckboxValue(frmMaster,'divcode');

	if (idivcode=="121"){
		imsg = "[�¶�����Ź���->����������Ź] �ΰ�� \r\n�¶��� ������ ���� ������ \r\n���������� ��Ź�԰�˴ϴ�. \r\n�԰� Ȯ������ ���� �Ͻðڽ��ϱ�?";
	}else if(idivcode=="131"){
		imsg = "[�¶�����Ź���->�����������] �ΰ�� \r\n�¶��� ������ ���� ������ \r\n���������� �����԰�˴ϴ�. \r\n�԰� Ȯ������ ���� �Ͻðڽ��ϱ�?";
	}else if(idivcode=="201"){
		imsg = "[�¶��θ������->�����������] �ΰ�� \r\n�¶��� ������ ���� ������ \r\n���������� �����԰�˴ϴ�. \r\n�԰� Ȯ������ ���� �Ͻðڽ��ϱ�?";
	}else{
		imsg = " �԰� Ȯ������ ���� �Ͻðڽ��ϱ�?";
	}

	var ret = confirm(imsg);

	if (ret){

		frmMaster.mode.value= "franupcheipgofinish";
		frmMaster.targetid.value= frmMaster.suplyer.value;
		frmMaster.submit();
	}
}

function DelAlink(frm,alinkcode){
	if (confirm('���õ� ����� ������ ���� �Ͻðڽ��ϱ�?')){
		frmMaster.mode.value = "delalinkipchul";
		frmMaster.alinkcode.value = alinkcode;
		frmMaster.submit();
	}
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmMaster" method="post" action="shopjumun_process.asp">
	<input type=hidden name="mode" value="">
	<input type=hidden name="opage" value="<%= opage %>">
	<input type=hidden name="ourl" value="<%= ourl %>">
	<input type=hidden name="oshopid" value="<%= oshopid %>">
	<input type=hidden name="ostatecd" value="<%= ostatecd %>">
	<input type=hidden name="odesinger" value="<%= odesinger %>">
	<input type=hidden name="masteridx" value="<%= idx %>">
	<!-- <input type=hidden name="shopid" value="<%= ojumunmaster.FOneItem.Fbaljuid %>"> -->
	<input type=hidden name="shopid" value="<%= ProtoShopid %>">
	<input type=hidden name="baljuname" value="<%= ojumunmaster.FOneItem.Fbaljuname %>">
	<input type=hidden name="reguser" value="<%= session("ssBctid") %>">
	<input type=hidden name="regname" value="<%= session("ssBctCname") %>">
	<input type=hidden name="orgbaljucode" value="<%= ojumunmaster.FOneItem.FBaljuCode %>">

	<input type=hidden name="targetid" value="">
	<input type=hidden name="alinkcode" value="">

	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red"><strong>�ֹ�����</strong></font>
				        &nbsp;
				        <b>[ <%= ojumunmaster.FOneItem.Fbaljucode %> ]</b>
				        &nbsp;
				        <% if (Not IsNULL(ojumunmaster.FOneItem.FALinkCode)) and (ojumunmaster.FOneItem.FALinkCode<>"") then %>
							��������:<%= ojumunmaster.FOneItem.FALinkCode %>
							<% if (not IsNULL(ojumunmaster.FOneItem.Fipchuldeldt)) then %>
								<font color="red">������</font>
							<% end if %>
							&nbsp;�ѼҺ�:<%= FormatNumber(ojumunmaster.FOneItem.Fipchulsellcash,0) %>
							<!-- &nbsp;�Ѱ��ް�:<%= FormatNumber(ojumunmaster.FOneItem.Fipchulsuplycash,0) %> -->
							&nbsp;�Ѹ��԰�:<%= FormatNumber(ojumunmaster.FOneItem.Fipchulbuycash,0) %>
							<!-- ���� ����� ������� ����. (2011-11-28 eastone)
							<input type="button" class="button" value="���� ����� ����" onClick="DelAlink(frmMaster,'<%= ojumunmaster.FOneItem.FALinkCode %>');">
							-->
						<% end if %>

				    </td>
				    <td align="right">
						<input type="button" class="button" value="������� �̵�" onclick="document.location='upchejumunlist.asp?page=<%= opage %>&designer=<%= odesinger %>&statecd=<%= ostatecd %>'">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- ��ܹ� �� -->

	<tr bgcolor="#FFFFFF">
		<td width="110" bgcolor="<%= adminColor("tabletop") %>" >����ó(�귣��)</td>
		<td width="400">
			<input type=hidden name="suplyer" value="<%= ojumunmaster.FOneItem.Ftargetid %>">
			<%= ojumunmaster.FOneItem.Ftargetid %>&nbsp;(<%= ojumunmaster.FOneItem.Ftargetname %>)
		</td>
		<td width="110" bgcolor="<%= adminColor("tabletop") %>" >����ó(�ֹ���)</td>
		<td><%= ojumunmaster.FOneItem.Fbaljuid %>&nbsp;(<%= ojumunmaster.FOneItem.Fbaljuname %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >�ֹ��Ͻ�</td>
		<td><%= ojumunmaster.FOneItem.Fregdate %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" >�԰��û��</td>
		<td>
			<input type="text" class="text" name="yyyymmdd" value="<%= yyyymmdd %>" size=12 readonly ><a href="javascript:calendarOpen(frmMaster.yyyymmdd);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<td colspan="3">
			<input type=radio name="statecd" value="0" <% if ojumunmaster.FOneItem.FStatecd="0" then response.write "checked" %> >�ֹ�����
			<input type=radio name="statecd" value="1" <% if ojumunmaster.FOneItem.FStatecd="1" then response.write "checked" %> >��ü�ֹ�Ȯ��
			<input type=radio name="statecd" value="5" <% if ojumunmaster.FOneItem.FStatecd="5" then response.write "checked" %> >��ü����غ�
			<input type=radio name="statecd" value="7" <% if ojumunmaster.FOneItem.FStatecd="7" then response.write "checked" %> >��ü���Ϸ�
			<input type=radio name="statecd" value="8" <% if ojumunmaster.FOneItem.FStatecd="8" then response.write "checked" %> >�԰���(�����Ϸ�)
			<% if ojumunmaster.FOneItem.FStatecd="9" then %>
			<input type=radio name="statecd" value="9" <% if ojumunmaster.FOneItem.FStatecd="9" then response.write "checked" %> >�԰�Ϸ�
				<% if (not IsNULL(ojumunmaster.FOneItem.Fipchuldeldt)) or (IsNULL(ojumunmaster.FOneItem.Falinkcode))  then %>
				&nbsp;<input type="button" class="button" value="���º���" onClick="ModiMaster(frmMaster)">
				<% else %>
				&nbsp;<input type="button" class="button" value="���º���" onClick="alert('���� ����� ������ ��밡���մϴ�.')">
				<% end if %>
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >������Է�</td>
		<td>
			�ù�� : <% drawSelectBoxDeliverCompany "songjangdiv", ojumunmaster.FOneItem.Fsongjangdiv %>
			������ȣ: <input type="text" class="text" name="songjangno" size=14 maxlength=16 value="<%= ojumunmaster.FOneItem.Fsongjangno %>" >
			<input type="hidden" name="songjangname" value="">
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" >�����</td>
		<td>
			<input type="text" class="text" name="beasongdate" value="<%= ojumunmaster.FOneItem.Fbeasongdate %>" size=12 readonly ><a href="javascript:calendarOpen(frmMaster.beasongdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
		</td>
	</tr>


	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >���Ա���</td>
		<td colspan="3">
			<input type=radio name="divcode" value="101" <% if (ojumunmaster.FOneItem.Fdivcode="101") or (IsNULL(ojumunmaster.FOneItem.Fdivcode) and (DefaultItemMwDiv="M")) then response.write "checked" %> >
			<% if (DefaultItemMwDiv="M") then %>
			<b>�������� ��������</b>
			<% else %>
			�������� ��������
			<% end if %>
			&nbsp;
			<input type=radio name="divcode" value="111" <% if (ojumunmaster.FOneItem.Fdivcode="111") or (IsNULL(ojumunmaster.FOneItem.Fdivcode) and (DefaultItemMwDiv="W"))  then response.write "checked" %> >
			�������� ������Ź
			&nbsp;&nbsp;
			<% if ojumunmaster.FOneItem.FStatecd="8" then %>
			<input type="button" class="button" value="�԰�ó��" onclick="IpgoFinish()">
			<% end if %>
			&nbsp;&nbsp;
			�¶��� : <%= oupchemwinfo.FOneItem.GetOnlineMwDivName %>&nbsp;<%= oupchemwinfo.FOneItem.GetOnlineDefaultmargine %>%
			&nbsp;������: <%= oupchemwinfo.FOneItem.GetfranChargeDivName %>&nbsp;<%= oupchemwinfo.FOneItem.GefranDefaultmargine %>%
		</td>
	</tr>


	<% if (ojumunmaster.FOneItem.FStatecd="6") then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td colspan="3"><input type=text name="ipgodate" value="<%= ojumunmaster.FOneItem.Fipgodate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.ipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		(��� ����� ��������)

		</td>
	</tr>
	<% elseif (ojumunmaster.FOneItem.FStatecd>6) then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�԰���</td>
		<td colspan="3"><input type="text" class="text" name="ipgodate" value="<%= ojumunmaster.FOneItem.Fipgodate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.ipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		(��� ����� ��������)
		</td>
	</tr>
	<% end if %>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�Һ��ڰ��հ�(�ֹ�)</td>
		<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsellcash,0) %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">���԰��ް��հ�(�ֹ�)</td>
		<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunbuycash,0) %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�Һ��ڰ��հ�(Ȯ��)</td>
		<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsellcash,0) %></b></td>
		<td bgcolor="<%= adminColor("tabletop") %>">���԰��ް��հ�(Ȯ��)</td>
		<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalbuycash,0) %></b></td>
	</tr>

	<!-- ���� ����� �ٸ��� �ִµ�...� ����Ÿ�� ǥ���Ѱ���
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#DDDDFF" width=100>�� ���ް�</td>
		<td colsapn="3"><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsuplycash,0) %> / <%= FormatNumber(ojumunmaster.FOneItem.Fjumunsuplycash,0) %></td>
	</tr>
	-->

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�ֹ��귣��</td>
		<td colspan="3"><textarea class="textarea_ro" cols=80 rows=3 readonly><%= ojumunmaster.FOneItem.FBrandList %></textarea></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ��û����</td>
		<td colspan="3"><textarea class="textarea" name=comment cols=80 rows=6><%= ojumunmaster.FOneItem.FComment %></textarea></td>
	</tr>

	</form>
</table>

<p>

<%

dim i,selltotal, suplytotal, buytotal
dim selltotalfix, suplytotalfix, buytotalfix
selltotal =0
suplytotal =0
buytotal =0
selltotalfix =0
suplytotalfix =0
buytotalfix =0
%>


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red"><strong>�󼼳���</strong></font>
				        &nbsp;
			        	<font color="#FF0000">�ٹ�</font>&nbsp;
			        	<font color="#000000">����</font>&nbsp;
			        	<font color="#0000FF">��������</font>
				    </td>
				    <td align="right">
						�ѰǼ�:  <%= ojumundetail.FResultCount %>
			        	&nbsp;
			        	<% if not isfixed then %>
							<input type="button" class="button" value="���ó�������" onClick="DelArr()">
						<% end if %>
						<% if not isfixed then %>
							<input type="button" class="button" value="��ǰ�߰�" onclick="AddItems(frmMaster)">
						<% end if %>

					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- ��ܹ� �� -->

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="20"><!-- <input type="checkbox"  name="ckall" onClick="AnSelectAllFrame(this.checked)"> --></td>
	    <td width="50">�̹���</td>
		<td width="80">���ڵ�</td>
		<td>�귣��ID</td>
		<td>��ǰ��</td>
		<td>�ɼǸ�</td>
		<td width="50">�ֹ���<br>�ǸŰ�</td>
		<td width="50">���԰�</td>
		<td width="30">����</td>
		<td width="60">�ֹ�<br>����</td>
		<td width="60">Ȯ��<br>����</td>
		<td width="60">Ȯ��<br>����</td>
		<td width="30">����<br>����<br>����</td>
		<% if isfixed then %>
		<td >���</td>
		<!-- td width="30">����<br>�԰�</td -->
		<% else %>
		<td width="90">���</td>
		<!-- td width="30">����<br>�԰�</td -->
		<% end if %>
	</tr>
	<% for i=0 to ojumundetail.FResultCount-1 %>
	<%
    if ((ojumunmaster.FOneItem.Fdivcode="101") and (ojumundetail.FItemList(i).Fcentermwdiv="W")) or ((ojumunmaster.FOneItem.Fdivcode="111") and (ojumundetail.FItemList(i).Fcentermwdiv="M")) then
        DIFFCenterMWDivExists = true
    end if

	selltotal  = selltotal + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno
	suplytotal = suplytotal + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno
	buytotal   = buytotal + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Fbaljuitemno

	selltotalfix  = selltotalfix + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Frealitemno
	suplytotalfix = suplytotalfix + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Frealitemno
	buytotalfix   = buytotalfix + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Frealitemno
	%>

	<%

	if (Not ojumundetail.FItemList(i).IsOnLineItem) then
		tmpcolor = "#0000FF"
	else
		if (ojumundetail.FItemList(i).IsUpchebeasong = True) then
			tmpcolor = "#000000"
		else
			tmpcolor = "#FF0000"
		end if
	end if

	%>
	<form name="frmBuyPrc_<%= i %>" method="post" action="shopjumun_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="masteridx" value="<%= idx %>">
	<input type="hidden" name="detailidx" value="<%= ojumundetail.FItemList(i).Fidx %>">
	<input type="hidden" name="itemgubun" value="<%= ojumundetail.FItemList(i).FItemGubun %>">
	<input type="hidden" name="itemid" value="<%= ojumundetail.FItemList(i).FItemID %>">
	<input type="hidden" name="itemoption" value="<%= ojumundetail.FItemList(i).Fitemoption %>">
	<input type="hidden" name="desingerid" value="<%= ojumundetail.FItemList(i).Fmakerid %>">
	<input type="hidden" name="sellcash" value="<%= ojumundetail.FItemList(i).FSellcash %>">

	<tr align="center" bgcolor="#FFFFFF">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><img src="<%= ojumundetail.FItemList(i).GetImageSmall %>" border="0" width="50" height="50" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
		<td>
			<a href="javascript:popOffItemEdit('<%= ojumundetail.FItemList(i).FItemGubun %><%= CHKIIF(ojumundetail.FItemList(i).FItemID>=1000000,format00(8,ojumundetail.FItemList(i).FItemID),format00(6,ojumundetail.FItemList(i).FItemID)) %><%= ojumundetail.FItemList(i).Fitemoption %>');">
			<font color="<%= tmpcolor %>">
			<%= ojumundetail.FItemList(i).FItemGubun %><%= CHKIIF(ojumundetail.FItemList(i).FItemID>=1000000,format00(8,ojumundetail.FItemList(i).FItemID),format00(6,ojumundetail.FItemList(i).FItemID)) %><%= ojumundetail.FItemList(i).Fitemoption %>
			</font>
			</a>
		</td>
		<td><%= ojumundetail.FItemList(i).Fmakerid %></td>
		<td align="left"><%= ojumundetail.FItemList(i).Fitemname %></td>
		<td><%= DdotFormat(ojumundetail.FItemList(i).Fitemoptionname,10) %></td>

		<td align=right>
		<% if   (ojumundetail.FItemList(i).FItemDefaultMwDiv<>"W") and (ojumundetail.FItemList(i).Fbuycash>ojumundetail.FItemList(i).Fsuplycash) then %>
		<b><font color=red><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></font></b>
		<% else %>
		<%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %>
		<% end if %>

		<% if (ojumundetail.FItemList(i).IsOnLineItem) and (ojumundetail.FItemList(i).Fsellcash<>ojumundetail.FItemList(i).Fonlinesellcash) then %>
		<br>
		<div ><font color=red>��:<%= FormatNumber(ojumundetail.FItemList(i).Fonlinesellcash,0) %></font></div>
		<% end if %>
	    </td>

		<input type="hidden" name="suplycash" value="<%= ojumundetail.FItemList(i).Fsuplycash %>">
		<td align=right>
			<input type="text" class="text" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>" size="7" maxlength="9" style="text-align:right">

			<% if (ojumundetail.FItemList(i).Fbuycash<>ojumundetail.FItemList(i).Fonlinebuycash) and ((ojumundetail.FItemList(i).FItemDefaultMwDiv="W") and (ojumundetail.FItemList(i).FoffChargeDiv="4")) then %>
			<div ><font color=red>��:<%= ojumundetail.FItemList(i).Fonlinebuycash %></font></div>
			<% end if %>
		</td>
		<td align="center">
	        <% if ojumundetail.FItemList(i).Fsellcash<>0 then %>
	            <%= CLng((ojumundetail.FItemList(i).Fsellcash-ojumundetail.FItemList(i).Fbuycash)/ojumundetail.FItemList(i).Fsellcash*100*100)/100 %>
	        <% end if %>
	    </td>
		<td align=center><input type="text" class="text" name="baljuitemno" value="<%= ojumundetail.FItemList(i).Fbaljuitemno %>"  size="3" maxlength="4" style="text-align:right"></td>
		<td align=center><input type="text" class="text" name="realitemno" value="<%= ojumundetail.FItemList(i).Frealitemno %>"  size="3" maxlength="4" style="text-align:right"></td>
		<td align=center>
			<% if Not IsNull(ojumundetail.FItemList(i).Fcheckitemno) then %>
			<% if (ojumundetail.FItemList(i).Fbaljuitemno <> ojumundetail.FItemList(i).Fcheckitemno) then %>
			<font color="red"><b><%= ojumundetail.FItemList(i).Fcheckitemno %></b></font>
			<% else %>
			<%= ojumundetail.FItemList(i).Fcheckitemno %>
			<% end if %>
			<% end if %>
		</td>
		<td align=center><%= ojumundetail.FItemList(i).Fcentermwdiv %></td>
		<% if isfixed then %>
			<td >

				<%= ojumundetail.FItemList(i).Fcomment %>
				<input type="hidden" name="comment" value="<%= ojumundetail.FItemList(i).Fcomment %>">
				<input type="hidden" name="ipgoflag" value="<%= ojumundetail.FItemList(i).Fipgoflag %>">
				<div align=center><%= ojumundetail.FItemList(i).GetOn2Off2DivName %></div>
			</td>
		<% else %>
			<td align=center >
				<input type="hidden" name="comment" value="<%= ojumundetail.FItemList(i).Fcomment %>">
				<div align=center><%= ojumundetail.FItemList(i).GetOn2Off2DivName %></div>
			</td>
			<input type=hidden name="ipgoflag" value="F">
		<% end if %>

		<input type=hidden name="defaultmaginflag" value="<%= ojumundetail.FItemList(i).GetNoinputDefaultmaginflag %>">
		<input type=hidden name="buymaginflag" value="<%= ojumundetail.FItemList(i).GetNoinputBuymaginflag %>">
		<input type=hidden name="suplymaginflag" value="<%= ojumundetail.FItemList(i).GetNoinputSuplymaginflag %>">


	</tr>
	</form>
	<% next %>

	<% if (ojumundetail.FResultCount>0) then %>
	<tr bgcolor="#FFFFFF">
		<td ></td>
		<td align="center">�Ѱ�</td>
		<td colspan="4" align="center">
		<td align=right>
			<%= formatNumber(selltotal,0) %><br>
			<b><%= formatNumber(selltotalfix,0) %></b>
		</td>
		<!--
		<td align=right>
			<%= formatNumber(suplytotal,0) %><br>
			<b><%= formatNumber(suplytotalfix,0) %></b>
		</td>
		-->
		<td align=right>
			<%= formatNumber(buytotal,0) %><br>
			<b><%= formatNumber(buytotalfix,0) %></b>
		</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<% end if %>
	<tr bgcolor="#FFFFFF">
		<td colspan="16" align=center>
		<% if (ojumunmaster.FOneItem.FStatecd="9") and  (not C_ADMIN_AUTH) then %>
			<b>�԰� �Ϸ�� ������ ���� �Ͻ� �� �����ϴ�.</b>
		<% else %>
			<input type="button" class="button" value="��ü����" onclick="SaveALL()">
			&nbsp;
			<input type="button" class="button" value="��ü����" onclick="DelMaster(frmMaster)">
		<% end if %>
		</td>
	</tr>
</table>
<%
set oupchemwinfo = Nothing
set ojumunmaster = Nothing
set ojumundetail = Nothing
%>
<form name="frmadd" method=post action="shopjumun_process.asp">
<input type=hidden name="mode" value="shopjumunitemaddarr">
<input type=hidden name="masteridx" value="<%= idx %>">
<input type=hidden name="detailidxarr" value="">
<input type=hidden name="itemgubunarr" value="">
<input type=hidden name="itemarr" value="">
<input type=hidden name="itemoptionarr" value="">
<input type=hidden name="sellcasharr" value="">
<input type=hidden name="suplycasharr" value="">
<input type=hidden name="buycasharr" value="">
<input type=hidden name="itemnoarr" value="">

<input type=hidden name="baljuitemnoarr" value="">
<input type=hidden name="realitemnoarr" value="">
<input type=hidden name="commentarr" value="">
<input type=hidden name="ipgoflagarr" value="">

<input type=hidden name="defaultmaginflagarr" value="">
<input type=hidden name="buymaginflagarr" value="">
<input type=hidden name="suplymaginflagarr" value="">


<input type=hidden name="yyyymmdd" value="">
<input type=hidden name="comment" value="">
<input type=hidden name="statecd" value="">
<input type=hidden name="beasongdate" value="">
<input type=hidden name="songjangdiv" value="">
<input type=hidden name="songjangno" value="">
<input type=hidden name="songjangname" value="">
<input type=hidden name="divcode" value="">



</form>
<% if (DIFFCenterMWDivExists) then %>
<script language='javascript'>
    alert('���� ���Ա����� ��ġ���� �ʽ��ϴ�. - ������ ���� ��� ');
</script>
<% end if  %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
