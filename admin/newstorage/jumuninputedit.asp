<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [LOG]��������>>�ֹ�������
' History : ������ ����
'			2022.09.15 �ѿ�� ����(ǰ�ǹ�ȣ �������� ���� ���� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim idx, isfixed, opage, ourl, odesigner, ostatecd, ojumunmaster, ojumundetail, purchasetype, companyno, yyyymmdd
    idx = RequestCheckVar(getNumeric(trim(request("idx"))),10)
	opage = RequestCheckVar(getNumeric(trim(request("opage"))),10)
ourl = request("ourl")
odesigner = request("odesigner")
ostatecd = request("ostatecd")

if idx="" then idx=0

set ojumunmaster = new COrderSheet
ojumunmaster.FRectIdx = idx
ojumunmaster.GetOneOrderSheetMaster

if ojumunmaster.FtotalCount<1 then
	response.write "�߸��� �ֹ���ȣ �Դϴ�."
	dbget.close() : response.end
end if

isfixed = ojumunmaster.FOneItem.FStatecd="9"

ojumunmaster.FRectMakerid = ojumunmaster.FOneItem.Ftargetid
ojumunmaster.fnGetBrandInfo
companyno =	ojumunmaster.Fcompanyno
purchasetype 		=	ojumunmaster.Fpurchasetype

set ojumundetail= new COrderSheet
ojumundetail.FRectIdx = idx
ojumundetail.GetOrderSheetDetail

yyyymmdd = Left(ojumunmaster.FOneItem.Fscheduledate,10)

dim IsTenOrder, HasReport, HasReportOld
IsTenOrder = (companyno = "211-87-00620")
HasReport = (ojumunmaster.FOneItem.FppMasterIdx <> "")
HasReportOld = (ojumunmaster.FOneItem.Freportidx <> "")

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

<% if ojumunmaster.FOneItem.FStatecd="0" then %>
var jumunwait = true;
<% else %>
var jumunwait = false;
<% end if %>

function publicbarreg(barcode){
	//var popwin = window.open('/common/popbarcode_input.asp?itembarcode=' + barcode,'popbarcode_input','width=500,height=400,resizable=yes,scrollbars=yes');
	var popwin = window.open('/admin/stock/popBarcodeManage.asp?itemcode=' + barcode,'popbarcode_input','width=550,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function upcheBarReg(barcode){
	var popwin = window.open('/admin/stock/popUpcheManageCode.asp?itemcode=' + barcode,'upcheBarReg_input','width=550,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
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
	popwin = window.open('popjumunitem.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=' + frm.masteridx.value,'shopjumunitemedit','width=800,height=600,scrollbars=yes,status=no');
	popwin.focus();
}


function AddItemsNew(frm){
	var popwin;
	var suplyer, shopid;

	if (frm.shopid.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.shopid.focus();
		return;
	}

	if (frm.suplyer.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	popwin = window.open('popjumunitemNew.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=' + frm.masteridx.value ,'upcheorderinputedit','width=960,height=600,scrollbars=yes,resizable=yes');
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
	//if (jumunwait!=true){
	//	alert('�ֹ����� ���°� �ƴϸ� �����Ͻ� �� �����ϴ�.');
	//	return;
	//}
	//alert(frm.beasongdate);
	if (frm.beasongdate!=undefined){
		//if (frm.beasongdate.value.length<1){
		//	alert('������� �Է��ϼ���.');
		//	frm.beasongdate.focus();
		//	return;
		//}

		//if (frm.songjangdiv.value.length<1){
		//	alert('�ù�縦 �����ϼ���.');
		//	frm.songjangdiv.focus();
		//	return;
		//}

		//if (frm.songjangno.value.length<1){
		//	alert('�����ȣ�� �Է��ϼ���.');
		//	frm.songjangno.focus();
		//	return;
		//}
		frm.songjangname.value = frm.songjangdiv.options[frm.songjangdiv.selectedIndex].text;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="modimaster";
		frm.submit();
	}
}

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner,immwdiv){
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
	frmadd.mwdivarr.value = immwdiv;
	frmadd.submit();
}

function DelDetail(detailfrm){


	if (confirm("���õ� ��ǰ�� �����մϴ�.") == true) {
		detailfrm.mode.value = "deldetail2";
		detailfrm.action = "orderinput_process.asp";
		detailfrm.submit();
	}

}

function ModiDetail(detailfrm){
	var frm;
	var found = false;


	if (((detailfrm.buycash.value*0) != 0) || ((detailfrm.baljuitemno.value*0) != 0) || ((detailfrm.realitemno.value*0) != 0)) {
		alert("�Է°��� Ȯ���ϼ���.");
		return false;
	}

	var compno = eval(detailfrm.baljuitemno.value)>eval(detailfrm.realitemno.value)?true:false

	if(false){
		if(detailfrm.dtstat.value=='ipt'){//�����Է�
			if (detailfrm.comment.value==''){
				alert('Ȯ�������� �����ϴ�.\������ �Է����ּ���');
				detailfrm.comment.focus();
				return false;
			}
		}else if ((detailfrm.dtstat.value=='optsso') || (detailfrm.dtstat.value=='sso')) { // �Ͻ�ǰ��
			if(detailfrm.comment.value==''){
				alert('Ȯ�������� �����ϴ�.\���԰� �������� �Է����ּ���');
				detailfrm.comment.focus();
				return false ;
			}
		}else if(detailfrm.dtstat==''){

		}

	}else{
		//frm.comdiv.style.display='none';
		//frm.comdiv.value='';
	}

	if (confirm("���õ� ��ǰ�� �����մϴ�.") == true) {
		detailfrm.mode.value = "modidetail2";
		detailfrm.action = "orderinput_process.asp";
		detailfrm.submit();
	}

}

function ModiDetailArr() {
	var frm;
	var frmAct = document.frmAct;

	var detailidxarr, suplycasharr, buycasharr, baljuitemnoarr, realitemnoarr, dtstatarr, commentarr;

	detailidxarr = "";
	suplycasharr = "";
	buycasharr = "";
	baljuitemnoarr = "";
	realitemnoarr = "";
	dtstatarr = "";
	commentarr = "";

	for (var i = 0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.chk.checked == true) {
				if (((frm.buycash.value*0) != 0) || ((frm.baljuitemno.value*0) != 0) || ((frm.realitemno.value*0) != 0)) {
					alert("�Է°��� Ȯ���ϼ���.");
					return;
				}

				detailidxarr = detailidxarr + "|" + frm.detailidx.value;
				suplycasharr = suplycasharr + "|" + frm.suplycash.value;
				buycasharr = buycasharr + "|" + frm.buycash.value;
				baljuitemnoarr = baljuitemnoarr + "|" + frm.baljuitemno.value;
				realitemnoarr = realitemnoarr + "|" + frm.realitemno.value;

				if (frm.dtstat) {
					dtstatarr = dtstatarr + "|" + frm.dtstat.value;
				} else {
					dtstatarr = dtstatarr + "|";
				}

				if (frm.comment) {
					commentarr = commentarr + "|" + frm.comment.value;
				} else {
					commentarr = commentarr + "|";
				}
			}
		}
	}

	if (detailidxarr == "") {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm("���û�ǰ �����Ͻðڽ��ϱ�?") == true) {
		frmAct.detailidxarr.value = detailidxarr;
		//frmAct.suplycasharr.value = suplycasharr;
		frmAct.buycasharr.value = buycasharr;
		frmAct.baljuitemnoarr.value = baljuitemnoarr;
		frmAct.realitemnoarr.value = realitemnoarr;
		frmAct.dtstatarr.value = dtstatarr;
		frmAct.commentarr.value = commentarr;

		frmAct.submit();
	}
}

function regAGVArr() {
	var frm;
	var frmagvreg = document.frmagvreg;

	var itemgubunarr, itemidarr, itemoptionarr, suplycasharr, buycasharr, baljuitemnoarr, realitemnoarr, agvitemnoarr, commentarr;
	itemgubunarr = "";
	itemidarr = "";
	itemoptionarr = "";
	suplycasharr = "";
	buycasharr = "";
	baljuitemnoarr = "";
	realitemnoarr = "";
	agvitemnoarr = "";
	commentarr = "";

	for (var i = 0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.chk.checked == true) {
				if (((frm.buycash.value*0) != 0) || ((frm.baljuitemno.value*0) != 0) || ((frm.realitemno.value*0) != 0)) {
					alert("�Է°��� Ȯ���ϼ���.");
					return;
				}

				itemgubunarr = itemgubunarr + "|" + frm.itemgubun.value;
				itemidarr = itemidarr + "|" + frm.itemid.value;
				itemoptionarr = itemoptionarr + "|" + frm.itemoption.value;
				suplycasharr = suplycasharr + "|" + frm.suplycash.value;
				buycasharr = buycasharr + "|" + frm.buycash.value;
				baljuitemnoarr = baljuitemnoarr + "|" + frm.baljuitemno.value;
				realitemnoarr = realitemnoarr + "|" + frm.realitemno.value;
				agvitemnoarr = agvitemnoarr + "|" + frm.agvitemno.value;

				if (frm.comment) {
					commentarr = commentarr + "|" + frm.comment.value;
				} else {
					commentarr = commentarr + "|";
				}
			}
		}
	}

	if (itemgubunarr == "") {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm("���û�ǰ�� AGV�������̽��� ���� �Ͻðڽ��ϱ�?") == true) {
		frmagvreg.itemgubunarr.value = itemgubunarr;
		frmagvreg.itemidarr.value = itemidarr;
		frmagvreg.itemoptionarr.value = itemoptionarr;
		//frmagvreg.suplycasharr.value = suplycasharr;
		frmagvreg.buycasharr.value = buycasharr;
		frmagvreg.baljuitemnoarr.value = baljuitemnoarr;
		frmagvreg.realitemnoarr.value = realitemnoarr;
		frmagvreg.agvitemnoarr.value = agvitemnoarr;
		frmagvreg.commentarr.value = commentarr;
		frmagvreg.submit();
	}
}

//Ȯ������&������ �׼�
function chkRealItemNo(tn){
	var frm = eval("frmBuyPrc_"+ tn);
	var t = frm.baljuitemno;
	var v= frm.realitemno;

	if (isNaN(v.value)||v.value.length<1){
		v.value=0;
	}else{
		v.value=parseInt(v.value);
	}

	//var seldiv = eval("seldiv" + tn);

	if(parseInt(t.value) > v.value){
		if (frm.dtstat!=''){
			$('#seldiv'+tn).html('<select class="select" name="dtstat" onchange="fnselcom(this.value,' + tn +');"><option value="ipt">�����Է�</option><option value="so">����</option><option value="sso">�Ͻ�ǰ��</option><option value="optso">�ɼ� ����</option><option value="optsso">�ɼ� �Ͻ�ǰ��</option></select><br>');
			//seldiv.innerHTML='<select name="dtstat" onchange="fnselcom(this.value,' + tn +');"><option value="ipt">�����Է�</option><option value="so">����</option><option value="sso">�Ͻ�ǰ��</option></select><br>';
			//fnselcom('ipt',tn);
		}else{
			$('#seldiv'+tn).empty();
			//seldiv.innerHTML='';
			//fnselcom('',tn);
		}
	}else{
		$('#seldiv'+tn).empty();
		//seldiv.innerHTML='<input type="text" name="comment" value=""  size="8" maxlength="10">';
		//fnselcom('',tn);
	}

}
//������ ǥ��
function fnselcom(val,tn){
	//var comdiv = eval("comdiv" + tn);
	if(val=='ipt'){
		$('#comdiv'+tn).show();
		$('#calspan'+tn).hide();
		//comdiv.innerHTML='<input type="text" name="comment" value=""  size="10" maxlength="10">';
	}else if ((val=='sso') || (val=='optsso')) {
		$('#comdiv'+tn).show();
		$('#calspan'+tn).show();
		//comdiv.innerHTML ='<input type="text" name="comment" value="" size="10" readonly ><a href="javascript:calendarOpen(document.all.comment['+eval(tn+1)+']);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>';
	}else{
		$('#comdiv'+tn).hide();
		$('#calspan'+tn).hide();
		//comdiv.innerHTML ='';
	}
	jsCheckBox(tn);
}

function IpgoProc(iidx){
	var frmdetail;

	<% if Not IsNull(ojumunmaster.FOneItem.Fcheckusername) then %>
		for (var i = 0;i < document.forms.length; i++) {
			frmdetail = document.forms[i];
			if (frmdetail.name.substr(0,9)=="frmBuyPrc") {
				if (frmdetail.realitemno.value*1 != frmdetail.checkitemno.value*1) {
					alert("Ȯ�������� ��ǰ������ �ٸ��ϴ�.");
					return;
				}
			}
		}
	<% end if %>

	var popwin = window.open("popipgoproc.asp?idx=" + iidx ,"popipgoproc","width=800,height=550,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function jsCheckBox(i){
    eval("document.frmBuyPrc_"+i+".chk").checked =  true ;
}
function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=400, height=300');
	winCal.focus();
}

//���ڰ��� ǰ�Ǽ� ���
function jsRegEapp(){
 var frm = document.frmMaster;


	var winEapp = window.open("","popE","width=1000,height=600,scrollbars=yes,resizable=yes");
	document.frmEapp.target = "popE";
	document.frmEapp.submit();
	winEapp.focus();
}

//���ڰ��� ǰ�Ǽ� ���뺸��
function jsViewEapp(reportidx,reportstate){
	var winEapp = window.open("/admin/approval/eapp/modeapp.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}

function jsViewEappNew(ppMasterIdx) {
    var pop = window.open("/admin/newstorage/PurchasedProductModify.asp?menupos=9175&idx=" + ppMasterIdx, "jsViewEappNew", "");
	pop.focus();
}

function jsRegEappNew() {
    var pop = window.open("/admin/newstorage/PurchasedProductList.asp?menupos=9175","jsRegEappNew","width=1600,height=800,scrollbars=yes,resizable=yes");
	pop.focus();
}

function DivisionOrder() {
	var frm;
	var frmDivAct = document.frmDivAct;

	var detailidxarr, baljuitemnoarr, realitemnoarr;

	detailidxarr = "";
	baljuitemnoarr = "";
	realitemnoarr = "";

	for (var i = 0; i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.chk.checked == true) {
				if (((frm.buycash.value*0) != 0) || ((frm.baljuitemno.value*0) != 0) || ((frm.realitemno.value*0) != 0)) {
					alert("�Է°��� Ȯ���ϼ���.");
					return;
				}
				detailidxarr = detailidxarr + "|" + frm.detailidx.value;
				baljuitemnoarr = baljuitemnoarr + "|" + frm.baljuitemno.value;
				realitemnoarr = realitemnoarr + "|" + frm.realitemno.value;
			}
		}
	}

	if (detailidxarr == "") {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm("���û�ǰ�� �ֹ����� �и� �Ͻðڽ��ϱ�?") == true) {
		frmDivAct.detailidxarr.value = detailidxarr;
		frmDivAct.baljuitemnoarr.value = baljuitemnoarr;
		frmDivAct.realitemnoarr.value = realitemnoarr;
		frmDivAct.submit();
	}
}

function checkboxAll(){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.chk.disabled!=true){
				frm.chk.checked = true;
				AnCheckClick(frm.chk);
			}
		}
	}
}

function AGVIpgoProc(){
	var frm;
	var frmAGVAct = document.frmAGVAct;
	var agvitemnoarr, itemgubunarr, itemarr, itemoptionarr;

	agvitemnoarr = "";
	itemgubunarr = "";
	itemarr = "";
	itemoptionarr = "";
	checkboxAll();

	for (var i = 0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.chk.checked == true) {
				if (((frm.agvitemno.value*0) != 0)) {
					alert("�Է°��� Ȯ���ϼ���.");
					return;
				}
				if(agvitemnoarr==""){

					agvitemnoarr = frm.agvitemno.value;
					itemgubunarr = frm.itemgubun.value;
					itemarr = frm.itemid.value;
					itemoptionarr = frm.itemoption.value;
				}
				else{
					agvitemnoarr = agvitemnoarr + "|" + frm.agvitemno.value;
					itemgubunarr = itemgubunarr + "|" + frm.itemgubun.value;
					itemarr = itemarr + "|" + frm.itemid.value;
					itemoptionarr = itemoptionarr + "|" + frm.itemoption.value;
				}
			}
		}
	}
	if (agvitemnoarr == "") {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}
	if (confirm("���û�ǰ�� AGV �԰� ���� �Ͻðڽ��ϱ�?") == true) {
		frmAGVAct.agvitemnoarr.value = agvitemnoarr;
		frmAGVAct.itemgubunarr.value = itemgubunarr;
		frmAGVAct.itemarr.value = itemarr;
		frmAGVAct.itemoptionarr.value = itemoptionarr;
		frmAGVAct.submit();
	}
}
function AGVIpgoDelProc(){
	if (confirm("AGV �԰� ���� �Ͻðڽ��ϱ�?") == true) {
		frmAGVAct.mode.value = "agvjumunitemdivisionorderdelete";
		frmAGVAct.submit();
	}
}
</script>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
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
				    </td>
				    <td align="right">
						<input type="button" class="button" value="������� �̵�" onclick="document.location='orderlist.asp?page=<%= opage %>&designer=<%= odesigner %>&statecd=<%= ostatecd %>'">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- ��ܹ� �� -->

	<form name="frmEapp" method="post" action="/admin/newstorage/jumun_regeapp.asp">
	<input type="hidden" name="iSL" value="<%=ojumunmaster.FOneItem.Fidx %>">
	<input type="hidden" name="purchasetype" value="<%=purchasetype%>">
	</form>
	<form name="frmMaster" method="post" action="orderinput_process.asp">
	<input type=hidden name="mode" value="">
	<input type=hidden name="opage" value="<%= opage %>">
	<input type=hidden name="ourl" value="<%= ourl %>">
	<input type=hidden name="masteridx" value="<%= idx %>">
	<input type=hidden name="shopid" value="<%= ojumunmaster.FOneItem.Fbaljuid %>">

	<tr bgcolor="#FFFFFF">
		<td width="110" bgcolor="<%= adminColor("tabletop") %>" >����ó(�귣��)</td>
		<td width="500">
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
		<td bgcolor="<%= adminColor("tabletop") %>" >�������</td>
		<td>
			<input type=radio name="statecd" value="0" <% if ojumunmaster.FOneItem.FStatecd="0" then response.write "checked" %> >�ֹ�����
			<input type=radio name="statecd" value="1" <% if ojumunmaster.FOneItem.FStatecd="1" then response.write "checked" %> >��ü�ֹ�Ȯ��
			<!--
			<input type=radio name="statecd" value="5" <% if ojumunmaster.FOneItem.FStatecd="5" then response.write "checked" %> >��ü����غ�
			-->
			<input type=radio name="statecd" value="7" <% if ojumunmaster.FOneItem.FStatecd="7" then response.write "checked" %> >��ü���Ϸ�
			<input type=radio name="statecd" value="8" <% if ojumunmaster.FOneItem.FStatecd="8" then response.write "checked" %> >��ǰ�Ϸ�(�԰���)
			<% if ojumunmaster.FOneItem.FStatecd="8" then %>
			<input type="button" class="button" value="�԰�ó��" onClick="IpgoProc('<%= ojumunmaster.FOneItem.Fidx %>')">
			<% end if %>
			<% if ojumunmaster.FOneItem.FStatecd="9" then %>
			<input type=radio name="statecd" value="9" <% if ojumunmaster.FOneItem.FStatecd="9" then response.write "checked" %> >�԰�Ϸ�
			<% if fnGetAGVCheckBalju(ojumunmaster.FOneItem.Fbaljucode) then %>
			<input type="button" class="button" value="AGV�԰����" onClick="AGVIpgoDelProc();">
			<% else %>
			<input type="button" class="button" value="AGV�԰�" onClick="AGVIpgoProc();">
			<% end if %>
			<% end if %>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" >�԰�����</td>
		<td>
			<%= Left(ojumunmaster.FOneItem.getScheduleIpgodate,10) %>
		</td>
	</tr>
	<% if ojumunmaster.FOneItem.FStatecd>="4" then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >������Է�</td>
		<td>
			�ù�� : <% drawSelectBoxDeliverCompany "songjangdiv", ojumunmaster.FOneItem.Fsongjangdiv %>
			������ȣ: <input type="text" class="text" name="songjangno" size=12 maxlength=16 value="<%= ojumunmaster.FOneItem.Fsongjangno %>" >
			<input type="hidden" name="songjangname" value="">
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" >�����</td>
		<td>
			<input type="text" class="text" name="beasongdate" value="<%= ojumunmaster.FOneItem.Fbeasongdate %>" size=12 readonly ><a href="javascript:calendarOpen(frmMaster.beasongdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
		</td>
	</tr>
	<% end if %>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="110" bgcolor="<%= adminColor("tabletop") %>" >ǰ�ǻ���</td>
		<td>
            <% if IsTenOrder then %>
                <% if HasReport then %>
					<%if ojumunmaster.FOneItem.FppReportstate="7" or ojumunmaster.FOneItem.FppReportstate="8" or ojumunmaster.FOneItem.FppReportstate="9" then %>
						ǰ�ǿϷ�
					<%elseif ojumunmaster.FOneItem.FppReportstate="5" then %>
						ǰ�ǹݷ�
					<%else%>
						������
					<%end if%>
                    (<a href="javascript:jsViewEappNew('<%= ojumunmaster.FOneItem.FppMasterIdx %>');">ǰ���ڷ� ��ȣ: <%= ojumunmaster.FOneItem.FppMasterIdx %></a>)
                <% elseif HasReportOld then %>
					<%if ojumunmaster.FOneItem.Freportstate="7" or ojumunmaster.FOneItem.Freportstate="8" or ojumunmaster.FOneItem.Freportstate="9" then %>
						ǰ�ǿϷ�
					<%elseif ojumunmaster.FOneItem.Freportstate="5" then %>
						ǰ�ǹݷ�
					<%else%>
						������
					<%end if%>
                    (<a href="javascript:jsViewEapp('<%= ojumunmaster.FOneItem.Freportidx %>','<%= ojumunmaster.FOneItem.Freportstate %>');">ǰ�ǹ�ȣ: <%= ojumunmaster.FOneItem.Freportidx %></a>)
                <% else %>
                    <strong>�ۼ���</strong>
                <% end if %>
            <% else %>
                <strong>ǰ�Ǵ�� �ƴ�</strong>
            <% end if %>
		</td>
		<td width="110" bgcolor="<%= adminColor("tabletop") %>" >���Ա���</td>
		<td >
			<input type=radio name="divcode" value="301" <% if ojumunmaster.FOneItem.Fdivcode="301" then response.write "checked" %> >����
			<input type=radio name="divcode" value="302" <% if ojumunmaster.FOneItem.Fdivcode="302" then response.write "checked" %> >��Ź
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�Һ��ڰ��հ�(�ֹ�)</td>
		<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsellcash,0) %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">���ް��հ�(�ֹ�)</td>
		<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunbuycash,0) %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�Һ��ڰ��հ�(Ȯ��)</td>
		<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsellcash,0) %></b></td>
		<td bgcolor="<%= adminColor("tabletop") %>">���ް��հ�(Ȯ��)</td>
		<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalbuycash,0) %></b></td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
		<td><%= ojumunmaster.FOneItem.Fcheckusername %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�԰�ó����</td>
		<td><%= ojumunmaster.FOneItem.Ffinishname %></td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�ֹ��귣��</td>
		<td colspan="3"><textarea class="textarea_ro" cols=80 rows=3 readonly><%= ojumunmaster.FOneItem.FBrandList %></textarea></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ��û����</td>
		<td colspan="3"><textarea class="textarea" name=comment cols=80 rows=6><%= ojumunmaster.FOneItem.FComment %></textarea></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ�亯</td>
		<td colspan="3"><%= nl2br(ojumunmaster.FOneItem.FReplyComment) %></td>
	</tr>
	</form>

	<!-- �ϴܹ� ����-->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if ojumunmaster.FOneItem.FStatecd="9" then %>
				<b>�԰� �Ϸ�� ������ ���� �Ͻ� �� �����ϴ�.</b>
				<% if C_ADMIN_AUTH then %>
					<br>
					<input type="button" class="button" value="[������]����" onclick="ModiMaster(frmMaster)">
					&nbsp;
					<input type="button" class="button" value="[������]��ü����" onclick="DelMaster(frmMaster)">
				<% else %>
				<!-- <input type="button" class="button" value="����" onclick="ModiMaster(frmMaster)">
				&nbsp;
				-->
				<!-- <input type="button" class="button" value="��ü����" onclick="DelMaster(frmMaster)"> -->
				<% end if %>
			<% else %>
				<input type="button" class="button" value="����" onclick="ModiMaster(frmMaster)">
				&nbsp;
				<input type="button" class="button" value="��ü����" onclick="DelMaster(frmMaster)">
				&nbsp;
			<% end if %>
            <% if IsTenOrder then %>
                <% if HasReport then %>
                    <input type="button" class="button"  value="ǰ���ڷ� ����" onClick="jsViewEappNew('<%= ojumunmaster.FOneItem.FppMasterIdx %>');" >
                <% elseif HasReportOld then %>
                    <input type="button" class="button"  value="ǰ�Ǽ� ���� OLD" onClick="jsViewEapp('<%=ojumunmaster.FOneItem.Freportidx%>','<%= ojumunmaster.FOneItem.Freportstate %>');" >
                <% else %>
                	<% if C_ADMIN_AUTH then %>
                    <input type="button" class="button" value="ǰ���ڷ� �ۼ�" onclick="jsRegEappNew()">
                    &nbsp;
                    &nbsp;
                    &nbsp;
                    <% end if %>
                    <input type="button" class="button" value="ǰ�Ǽ� �ۼ� OLD" onclick="jsRegEapp()">
                <% end if %>
            <% end if %>
		</td>
	</tr>
	<!-- �ϴܹ� ��-->
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
		<td colspan="17">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red"><strong>�󼼳���</strong></font>
			        	&nbsp;&nbsp;
			        	<font color="<%= mwdivColor("M") %>">����</font>&nbsp;
			        	<font color="<%= mwdivColor("W") %>">��Ź</font>&nbsp;
			        	<font color="<%= mwdivColor("U") %>">��ü���</font>
				    </td>
				    <td align="right">
						�ѰǼ�:  <%= ojumundetail.FResultCount %>
			        	&nbsp;
			        	<% if ojumunmaster.FOneItem.FStatecd="9" then %>

						<% else %>
			<!--			<input type="button" class="button" value="��ǰ�߰�_old" onclick="AddItems(frmMaster)">	-->
							<input type="button" class="button" value="��ǰ�߰�" onclick="AddItemsNew(frmMaster)">
						<% end if %>
							<!--
							&nbsp;&nbsp;&nbsp;
							<input type=button value="��ü����" onclick="ModiArr(frmMaster)">
							-->
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- ��ܹ� �� -->



	<form name="frmDetail" method="post" action="">
	<input type="hidden" name="masteridx" value="<%= idx %>">
	<input type="hidden" name="mode" value="">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20">-</td>
		<td width="50">�̹���</td>
		<td width="100">��ǰ�ڵ�</td>
		<td width="100">������ڵ�</td>
		<td width="100">��ü�����ڵ�</td>
		<td>��ǰ��</td>
		<td>�ɼǸ�</td>
		<td width="70">�ǸŰ�</td>
		<td width="70">���԰�</td>
		<td width="30">����<br>����</td>
		<td width="30">����<br>����<br>����</td>
		<td width="60">�ֹ�����</td>
		<td width="60">Ȯ������</td>
		<td width="60">��ǰ����</td>
		<td width="60">AGV����</td>
		<% if isfixed then %>
		<td width="100">���</td>
		<% else %>
		<td width="100">���</td>
		<% end if %>
		<td width="70"></td>
	</tr>
	</form>
	<% for i=0 to ojumundetail.FResultCount-1 %>
	<%
	selltotal  = selltotal + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno
	suplytotal = suplytotal + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno
	buytotal   = buytotal + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Fbaljuitemno

	selltotalfix  = selltotalfix + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Frealitemno
	suplytotalfix = suplytotalfix + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Frealitemno
	buytotalfix   = buytotalfix + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Frealitemno
	%>
	<form name="frmBuyPrc_<%=i %>" method="post" action="">
	<input type="hidden" name="masteridx" value="<%= idx %>">
	<input type="hidden" name="detailidx" value="<%= ojumundetail.FItemList(i).Fidx %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemgubun" value="<%= ojumundetail.FItemList(i).FItemGubun %>">
	<input type="hidden" name="itemid" value="<%= ojumundetail.FItemList(i).FItemID %>">
	<input type="hidden" name="itemoption" value="<%= ojumundetail.FItemList(i).Fitemoption %>">
	<input type="hidden" name="desingerid" value="<%= ojumundetail.FItemList(i).Fmakerid %>">
	<input type="hidden" name="sellcash" value="<%= ojumundetail.FItemList(i).FSellcash %>">
	<input type="hidden" name="suplycash" value="<%= ojumundetail.FItemList(i).FSuplycash %>">
	<input type="hidden" name="mwdiv" value="<%= ojumundetail.FItemList(i).FItemDefaultMwDiv %>">

	<!-- <input type="hidden" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>"> -->
	<tr bgcolor="#FFFFFF">
		<td><input type="checkbox" name="chk" value="<%= i %>" onClick="AnCheckClick(this);"></td>
		<td><img src="<%= CHKIIF((ojumundetail.FItemList(i).FItemGubun="10"), ojumundetail.FItemList(i).Fsmallimage, ojumundetail.FItemList(i).Foffimgsmall) %>" width=50 height=50 /></td>
		<td>
			<font color="<%= mwdivColor(ojumundetail.FItemList(i).FItemDefaultMwDiv) %>"><%= ojumundetail.FItemList(i).FItemGubun %>-<%= BF_GetFormattedItemId(ojumundetail.FItemList(i).FItemID) %>-<%= ojumundetail.FItemList(i).Fitemoption %></font>
		</td>
		<td>
			<a href="javascript:publicbarreg('<%= ojumundetail.FItemList(i).FItemGubun %><%= BF_GetFormattedItemId(ojumundetail.FItemList(i).FItemID) %><%= ojumundetail.FItemList(i).Fitemoption %>');">
			<% if ojumundetail.FItemList(i).FPublicBarcode<>"" then %>
				<font color="#AAAAAA"><b><%= ojumundetail.FItemList(i).FPublicBarcode %></b></font>
			<% else %>
				���>>
			<% end if %>
			</a>
		</td>
		<td>
			<a href="javascript:upcheBarReg('<%= ojumundetail.FItemList(i).FItemGubun %><%= BF_GetFormattedItemId(ojumundetail.FItemList(i).FItemID) %><%= ojumundetail.FItemList(i).Fitemoption %>');">
				<% if ojumundetail.FItemList(i).FUpcheManageCode<>"" then %>
				<font color="#AAAAAA"><b><%= ojumundetail.FItemList(i).FUpcheManageCode %></b></font>
				<% else %>
				���>>
				<% end if %>
			</a>
		</td>
		<td><%= ojumundetail.FItemList(i).Fitemname %></td>
		<td><%= ojumundetail.FItemList(i).Fitemoptionname %></td>

		<td align=right><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></td>
		<td align=right><input type="text" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>" size="7" maxlength="12" onKeyup="jsCheckBox(<%=i%>);"></td>
		<td align=center><%= ojumundetail.FItemList(i).FItemDefaultMwDiv %></td>
		<td align=center><%= ojumundetail.FItemList(i).Fcentermwdiv %></td>
		<td width="50" align=center><input type="text" class="text" name="baljuitemno" value="<%= ojumundetail.FItemList(i).Fbaljuitemno %>"  size="4" maxlength="8" onKeyup="jsCheckBox(<%=i%>);"></td>
		<td width="50" align=center><input type="text" class="text" name="realitemno" value="<%= ojumundetail.FItemList(i).Frealitemno %>"  size="4" maxlength="8" onKeyup="jsCheckBox(<%=i%>);chkRealItemNo(<%= i %>);" onfocus="this.selected"></td>
		<input type="hidden" name="checkitemno" value="<%= ojumundetail.FItemList(i).Fcheckitemno %>">
		<td width="50" align=center>
			<% if Not IsNull(ojumunmaster.FOneItem.Fcheckusersn) then %>
				<% if (ojumundetail.FItemList(i).Frealitemno <> ojumundetail.FItemList(i).Fcheckitemno) and Not IsNull(ojumunmaster.FOneItem.Fcheckusersn) then %><b><font color=red>&lt;=&nbsp;&nbsp;<% end if %>
				<%= ojumundetail.FItemList(i).Fcheckitemno %>
			<% end if %>
		</td>
		<td width="60" align=center><input type="text" class="text" name="agvitemno" size="4" maxlength="8" value="<%= ojumundetail.FItemList(i).Frealitemno %>"></td>
		<% if isfixed then %>
		<td><%= ojumundetail.FItemList(i).FDetail_status %><br><%= ojumundetail.FItemList(i).Fcomment %></td>
		<td></td>
		<% else %>
		<!-- ����
		<td align=center><input type="text" name="comment" value="<%= ojumundetail.FItemList(i).Fcomment %>"  size="10" maxlength="10"></td>
		-->
		<td>
			<span id="seldiv<%=i%>">
				<% if ojumundetail.FItemList(i).FDetail_status<>"" then %>
					<select class="select" name="dtstat" onchange="fnselcom(this.value,<%= i %>);">
						<option value="ipt" <% if ojumundetail.FItemList(i).FDetail_status="�����Է�" then response.write "selected" %>>�����Է�</option>
						<option value="so" <% if ojumundetail.FItemList(i).FDetail_status="����" then response.write "selected" %>>����</option>
						<option value="sso" <% if ojumundetail.FItemList(i).FDetail_status="�Ͻ�ǰ��" then response.write "selected" %>>�Ͻ�ǰ��</option>
                        <option value="optso" <% if ojumundetail.FItemList(i).FDetail_status="�ɼǴ���" then response.write "selected" %>>�ɼ� ����</option>
                        <option value="optsso" <% if ojumundetail.FItemList(i).FDetail_status="�ɼ��Ͻ�ǰ��" then response.write "selected" %>>�ɼ� �Ͻ�ǰ��</option>
					</select><br>
				<% end if %>
			</span>
			<span id="comdiv<%=i%>">
				<% if (ojumundetail.FItemList(i).FDetail_status="����") or (ojumundetail.FItemList(i).FDetail_status="�ɼǴ���") then %>

				<% else %>
					<input type="text" class="text" name="comment" value="<%= ojumundetail.FItemList(i).Fdetail_description %>"  size="10" maxlength="10"  onKeyup="jsCheckBox(<%=i%>);">
				<% end if %>
			</span>
			<span id="calspan<%=i%>" style="display:none;">
				<a href="javascript:jsPopCal('comment['+eval(<%=i%>+1)+']');"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
			</span>
		</td>
		<td align="center">
		<input type="button" class="button" class="button" value="����" onclick="ModiDetail(frmBuyPrc_<%=i %>)">
    	<input type="button" class="button" class="button" value="����" onclick="DelDetail(frmBuyPrc_<%=i %>)">
    	</td>
		<% end if %>

	</tr>
	</form>
	<% next %>

</table>

<% if (ojumundetail.FResultCount>0) then %>
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	�ֹ� �Һ��ڰ��� : <%= formatNumber(selltotal,0) %>&nbsp;&nbsp;
        	�ֹ� ���ް��� : <%= formatNumber(buytotal,0) %>
        </td>
        <td align="right">
        	Ȯ�� �Һ��ڰ��� : <b><%= formatNumber(selltotalfix,0) %></b>&nbsp;&nbsp;
        	Ȯ�� ���ް��� : <b><%= formatNumber(buytotalfix,0) %></b>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
</table>
<!-- ǥ �߰��� ��-->
<% end if %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;
        	<input type="button" class="button" value=" ���û�ǰ���� " onclick="ModiDetailArr(frmDetail)">
    		<input type="button" class="button" value=" ���û�ǰ�ֹ����и� " onclick="DivisionOrder(frmDetail)">
			<input type="button" class="button" value=" ���û�ǰAGV�������̽�����" onclick="regAGVArr(frmDetail);">
	    </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>

<form name="frmadd" method=post action="orderinput_process.asp">
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
<input type=hidden name="mwdivarr" value="">
</form>

<form name="frmAct" method=post action="orderinput_process.asp">
<input type=hidden name="mode" value="shopjumunitemmodifyarr">
<input type=hidden name="masteridx" value="<%= idx %>">
<input type=hidden name="detailidxarr" value="">
<input type=hidden name="suplycasharr" value="">
<input type=hidden name="buycasharr" value="">
<input type=hidden name="baljuitemnoarr" value="">
<input type=hidden name="realitemnoarr" value="">
<input type=hidden name="dtstatarr" value="">
<input type=hidden name="commentarr" value="">
</form>

<form name="frmDivAct" method=post action="orderinput_process.asp">
<input type=hidden name="mode" value="shopjumunitemdivisionorder">
<input type=hidden name="masteridx" value="<%= idx %>">
<input type=hidden name="detailidxarr" value="">
<input type=hidden name="baljuitemnoarr" value="">
<input type=hidden name="realitemnoarr" value="">
</form>

<form name="frmAGVAct" method=post action="orderinput_process.asp">
<input type=hidden name="mode" value="agvjumunitemdivisionorder">
<input type=hidden name="masteridx" value="<%= idx %>">
<input type=hidden name="baljucode" value="<%= ojumunmaster.FOneItem.Fbaljucode %>">
<input type=hidden name="agvitemnoarr" value="">
<input type=hidden name="itemgubunarr" value="">
<input type=hidden name="itemarr" value="">
<input type=hidden name="itemoptionarr" value="">
</form>
<form name="frmagvreg" method="post" action="/admin/logics/logics_agv_pickup_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="agvregarr">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="baljuitemnoarr" value="">
<input type="hidden" name="realitemnoarr" value="">
<input type="hidden" name="agvitemnoarr" value="">
<input type="hidden" name="commentarr" value="">
<input type="hidden" name="code" value="<%= ojumunmaster.FOneItem.Fbaljucode %>">
<input type="hidden" name="refergubun" value="A">
</form>
<%
set ojumunmaster = Nothing
set ojumundetail = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
