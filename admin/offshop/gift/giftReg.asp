<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ǰ ����
' History : 2010.03.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/gift/gift_Cls.asp"-->

<%
Dim clsGift , evt_code, cEvent, gift_status, sStateDesc , gift_code ,gift_scope
Dim gift_name, gift_startdate, gift_enddate, opendate, gift_type , gift_itemname
dim makerid , gift_range1 , gift_range2 , giftkind_code , giftkind_type ,giftkind_cnt
dim giftkind_limit, gift_using ,regdate , adminid , giftkind_name , closedate
dim gift_img, gift_imgurl, gift_img_50X50_url

dim cEvtCont
dim chkdisp , evt_using , evt_kind , evt_name , evt_startdate ,evt_enddate
dim evt_state , evt_prizedate , evt_opendate , evt_closedate , brand , partMDid ,evt_forward ,issale
dim evt_comment , evt_regdate , shopid , isgift ,israck ,isprize , isracknum ,racknum  ,img_basic
dim shopname
dim evt_sStateDesc

dim itemgubun, shopitemid, itemoption, shopitemname
dim gift_itemgubun, gift_shopitemid, gift_itemoption
dim receiptstring

dim gift_scope_add, giftkind_limit_sold
'
''// ����ǰ ������ ǥ�ó���
'Function fnComGetEventConditionStr(ByVal Fgiftkind_type, ByVal Fgift_scope,ByVal Fgift_type,ByVal Fgift_range1, ByVal Fgift_range2,ByVal FGiftName,ByVal Fgiftkind_cnt, ByVal Fgiftkind_orgcnt, ByVal Fgiftkind_limit, ByVal Fgiftkind_givecnt,ByVal FMakerid)
'Dim reStr
'dim remainEa
'
'        reStr = ""
'        if (FMakerid<> "") then
'        	reStr = reStr + FMakerid + " "
'        end if
'        if (Fgift_scope="1") then
'            reStr = reStr + "��ü ���� �� "
'        elseif (Fgift_scope="2") then
'            reStr = reStr + "�̺�Ʈ��ϻ�ǰ "
'        elseif (Fgift_scope="3") then
'            reStr = reStr + "���ú귣���ǰ "
'        elseif (Fgift_scope="4") then
'            reStr = reStr + "�̺�Ʈ�׷��ǰ"
'        elseif (Fgift_scope="5") then
'            reStr = reStr + "���û�ǰ"
'        end if
'
'        if (Fgift_type="1") then
'            reStr = reStr + "��� ������"
'        elseif (Fgift_type="2") then
'            if (Fgift_range2=0) then
'                reStr = reStr + CStr(Fgift_range1) + " �� �̻� ���Ž� "
'            else
'                reStr = reStr + CStr(Fgift_range1) + "~" + CStr(Fgift_range2) + " �� ���Ž� "
'            end if
'        elseif (Fgift_type="3") then
'            if (Fgift_range2=0) then
'                reStr = reStr + CStr(Fgift_range1) + " �� �̻� ���Ž� "
'            else
'                reStr = reStr + CStr(Fgift_range1) + "~" + CStr(Fgift_range2) + " �� ���Ž� "
'            end if
'        end if
'        reStr = reStr &"'"&  FGiftName &"' "
'        reStr = reStr &  Cstr(Fgiftkind_orgcnt) & " �� "
'
'        if (Fgiftkind_type=2) then
'            reStr = reStr + "[1+1]"
'             reStr = reStr & "(�� "& Cstr(Fgiftkind_cnt) & " ��)"
'        elseif (Fgiftkind_type=3) then
'            reStr = reStr + "[1:1]"
'             reStr = reStr & "(�� "& Cstr(Fgiftkind_cnt) & " ��)"
'        end if
'         reStr = reStr + " ����"
'
'
'        if Fgiftkind_limit<>0 then
'            reStr = reStr & " ������ [" & Fgiftkind_limit & "]"
'            remainEa = Fgiftkind_limit-Fgiftkind_givecnt
'            if (remainEa<0) then remainEa=0
'             reStr = reStr & " ���糲������ " & remainEa
'        end if
'        fnComGetEventConditionStr = reStr
' End Function

menupos = requestCheckVar(request("menupos"),10)
evt_code = requestCheckVar(Request("evt_code"),10)
gift_code = requestCheckVar(Request("gift_code"),10)
'gift_type = 2

if evt_code = "" then
	Alert_return("�߸��� �����Դϴ�. ���� �̺�Ʈ�� ����ϼ���.")
	dbget.close()	:	response.End
end if

'==============================================================================
set cEvtCont = new cevent_list
	cEvtCont.frectevt_code = evt_code	'�̺�Ʈ �ڵ�

	'�̺�Ʈ ���� ��������
	cEvtCont.fnGetEventCont_off
	evt_kind = cEvtCont.FOneItem.fevt_kind
	evt_name = cEvtCont.FOneItem.fevt_name
	evt_startdate = cEvtCont.FOneItem.Fevt_startdate
	evt_enddate = cEvtCont.FOneItem.Fevt_enddate
	evt_prizedate =	cEvtCont.FOneItem.Fevt_prizedate
	evt_state =	cEvtCont.FOneItem.Fevt_state
	IF datediff("d",now,evt_enddate) <0 THEN evt_state = 9 '�Ⱓ �ʰ��� ����ǥ��
	evt_regdate	= cEvtCont.FOneItem.fevt_regdate
	evt_using = cEvtCont.FOneItem.Fevt_using
	shopid = cEvtCont.FOneItem.fshopid
	shopname = cEvtCont.FOneItem.fshopname
	evt_opendate = cEvtCont.FOneItem.fopendate
	evt_closedate = cEvtCont.FOneItem.fclosedate

	'�̺�Ʈ ȭ�鼳�� ���� ��������
	cEvtCont.fnGetEventDisplay_off
	chkdisp = cEvtCont.FOneItem.FChkDisp
	issale = cEvtCont.FOneItem.fissale
	isgift = cEvtCont.FOneItem.fisgift
	israck = cEvtCont.FOneItem.fisrack
	isprize = cEvtCont.FOneItem.fisprize
	isracknum = cEvtCont.FOneItem.fisracknum
	partMDid = cEvtCont.FOneItem.FpartMDid
	evt_forward	= db2html(cEvtCont.FOneItem.Fevt_forward)
	brand = cEvtCont.FOneItem.Fbrand
	evt_comment = cEvtCont.FOneItem.fevt_comment
 	chkdisp	= cEvtCont.FOneItem.fchkdisp
	img_basic = cEvtCont.FOneItem.fimg_basic

set cEvtCont = nothing



'==============================================================================
set cEvent = new cevent_list
	cEvent.Frectevt_code = evt_code
	cEvent.fnGetEventConts

	gift_name 		= cEvent.foneitem.fevt_name
	gift_startdate	= cEvent.foneitem.fevt_startdate
	gift_enddate	= cEvent.foneitem.fevt_enddate
	gift_status  	= cEvent.foneitem.fevt_state
	opendate		= cEvent.foneitem.FOpenDate

	evt_sStateDesc = cEvent.foneitem.fevt_statedesc
	sStateDesc = cEvent.foneitem.fevt_statedesc		'�ű� ����϶� �̺�Ʈ�� ���¸� �����´�.
set cEvent = nothing



'==============================================================================
dim isregstate

isregstate = true


'//�űԵ��
if gift_code = "" then

	'�̺�Ʈ ���¿� ����ǰ ���� ��Īó��(�������� ���´� ��� ������)
	if gift_status < 6 then gift_status = 0
	giftkind_cnt = 1

'//����
else
	isregstate = false

	set clsGift = new cgift_list
		clsGift.frectgift_code = gift_code
		clsGift.fnGetGiftConts_off

		gift_name			= clsGift.foneitem.fgift_name
		gift_scope 			= clsGift.foneitem.fgift_scope
		evt_code			= clsGift.foneitem.fevt_code
		gift_type			= clsGift.foneitem.fgift_type
		gift_range1			= clsGift.foneitem.fgift_range1
		gift_range2 		= clsGift.foneitem.fgift_range2
		giftkind_code		= clsGift.foneitem.fgiftkind_code

		makerid				= clsGift.foneitem.fmakerid

		itemgubun			= clsGift.foneitem.fitemgubun
		shopitemid			= clsGift.foneitem.fshopitemid
		itemoption			= clsGift.foneitem.fitemoption
		shopitemname		= clsGift.foneitem.fshopitemname

		gift_itemgubun		= clsGift.foneitem.fgift_itemgubun
		gift_shopitemid		= clsGift.foneitem.fgift_shopitemid
		gift_itemoption		= clsGift.foneitem.fgift_itemoption

		giftkind_type		= clsGift.foneitem.fgiftkind_type
		giftkind_cnt		= clsGift.foneitem.fgiftkind_cnt
		giftkind_limit		= clsGift.foneitem.fgiftkind_limit
		gift_startdate		= clsGift.foneitem.fgift_startdate
		gift_enddate		= clsGift.foneitem.fgift_enddate
		gift_status			= clsGift.foneitem.fgift_status
		gift_using     		= clsGift.foneitem.fgift_using
		regdate				= clsGift.foneitem.fregdate
		adminid 			= clsGift.foneitem.fadminid
		giftkind_name 		= clsGift.foneitem.fgiftkind_name
		opendate			= clsGift.foneitem.fopendate
		closedate			= clsGift.foneitem.fclosedate
		gift_itemname		= clsGift.foneitem.fgift_itemname
		gift_img			= clsGift.foneitem.fgift_img

		gift_imgurl			= clsGift.foneitem.GetMobileGiftImage
		gift_img_50X50_url	= clsGift.foneitem.GetMobileGiftImage50X50

		receiptstring		= clsGift.foneitem.fnGetReceiptString

		gift_scope_add		= clsGift.foneitem.fgift_scope_add
		giftkind_limit_sold	= clsGift.foneitem.fgiftkind_limit_sold

	set clsGift = nothing

	  '�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	sStateDesc 	= fnSetCommonCodeArr_off("gift_status",False)
end if
%>

<script language="javascript">

	//����ǰ ���
	function jsSubmitGift(){
		var frm = document.frmReg;

		// ====================================================================
		// ��¥ ����
		if(!frm.gift_name.value){
			alert("������ �Է��� �ּ���");
			frm.gift_name.focus();
			return;
		}

		if(!frm.gift_startdate.value ){
		  	alert("�������� �Է����ּ���");
		 	frm.gift_startdate.focus();
		  	return;
	  	}

	  	if(frm.gift_enddate.value){
		  	if(frm.gift_startdate.value > frm.gift_enddate.value){
			  	alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			  	return;
		  	}
		}

		var nowDate = "<%=date()%>";

		if(frm.gift_status.value==7){
			if(frm.opendate.value !=""){
				nowDate = '<%IF opendate <> ""THEN%><%=FormatDate(opendate,"0000-00-00")%><%END IF%>';
			}

			/*
			if(frm.gift_startdate.value < nowDate){
				alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
				return;
			}
			*/
		}

		// ====================================================================
		if (frm.gift_scope.value == "") {
			alert("��������� �����ϼ���.");
			return;
		} else if (frm.gift_scope.value == "1") {
			//
		} else if (frm.gift_scope.value == "5") {
			// ��� ��ǰ
			if (frm.shopitemid.value == "") {
				alert("��ǰ�� �������ּ���.");
				return;
			}
		} else if (frm.gift_scope.value == "6") {
			// ��� �귣��
			if (frm.makerid.value == "") {
				alert("�귣�带 �������ּ���.");
				return;
			}
		} else if (frm.gift_scope.value == "7") {
			// ���������������
			if (frm.gift_scope_add.value == "") {
				alert("����������������� �Է����ּ���.");
				return;
			}
		}

		// ====================================================================
		if (frm.gift_type.value == "") {
			alert("���������� �������ּ���.");
			return;

		} else if (frm.gift_type.value != "1") {
			if ((frm.gift_range1.value*0 != 0) || (frm.gift_range2.value*0 != 0) || (frm.gift_range1.value == "") || (frm.gift_range2.value == "")) {
				alert("���������� ��Ȯ�� �Է��ϼ���.");
				return;
			}

			if (frm.gift_range1.value*1 == 0) {
				if (confirm("���������� 0 ���� �����Ͽ����ϴ�. �����Ͻðڽ��ϱ�?") != true) {
					return;
				}
			}

		} else {
			frm.gift_range1.value = 0;
			frm.gift_range2.value = 0;
		}

		// ====================================================================
		if(frm.gift_shopitemid.value == ""){
			alert("����ǰ�� �˻��ؼ� �Է����ּ���.");
			return;
		}

		if ((frm.giftkind_cnt.value*0 != 0) || (frm.giftkind_cnt.value == "")) {
			alert("����ǰ ������ ��Ȯ�� �Է��ϼ���.");
			return;
		}

		if (frm.chkLimit.checked == true) {
			if ((frm.giftkind_limit.value*0 != 0) || (frm.giftkind_limit.value == "")) {
				alert("����ǰ ���� ������ ��Ȯ�� �Է��ϼ���.");
				return;
			}
		} else {
			frm.giftkind_limit.value = 0;
			frm.giftkind_limit_sold.value = 0;
		}

		// ====================================================================
		if(frm.gift_itemname.value == ""){
			frm.gift_itemname.value = frm.giftkind_name.value;
		}

		// ====================================================================
		if (confirm("�����Ͻðڽ��ϱ�?") == true) {
			jsChkGiftScope(frm.gift_scope.value);
			jsResetHiddenData();

			frm.submit();
		}

	}

	//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	// �������
	// ǥ�ø� �ٲ۴�. �Ⱥ��̴� �κ� ����Ÿ ������ �����Ҷ� �Ѵ�.
	function jsChkGiftScope(iVal){
		jsHideAll();

		if(iVal == 1){
			// ��ü����
		} else if (iVal == 5) {
			// ��ϻ�ǰ
			document.all.showitemid.style.display = "";
		} else if (iVal==6) {
			// ��Ϻ귣��
			document.all.showmakerid.style.display = "";
		} else if (iVal==7) {
			// �����Ǵ�
			document.all.showaddcondition.style.display = "";
		}else{
			// ERROR
		}

		if (iVal != 5) {
			// ��������� ��ϻ�ǰ�� �ƴϸ� ���������Ѵ�.
			var chk2 	= document.getElementById('tmpgiftkind_givecnt2');
			var chk3 	= document.getElementById('tmpgiftkind_givecnt3');

			chk2.checked=false;
			chk3.checked=false;

			jsResetGiftNo();
		}
	}

	function jsHideAll() {
		document.all.showmakerid.style.display = "none";
		document.all.showitemid.style.display = "none";
		document.all.showaddcondition.style.display = "none";
	}

	function jsResetHiddenData() {
		// �귣������ ����
		if (document.all.showmakerid.style.display == "none") {
			document.all.makerid.value = "";
		}

		// ��ǰ���� ����
		if (document.all.showitemid.style.display == "none") {
			document.all.itemgubun.value = "";
			document.all.shopitemid.value = "";
			document.all.itemoption.value = "";
			document.all.shopitemname.value = "";
		}

		// ��������������� ����
		if (document.all.showaddcondition.style.display == "none") {
			document.all.gift_scope_add.value = "";
		}
	}

	function jsResetGiftNo() {
		// ����ǰ ��������
		var chk2 	= document.getElementById('tmpgiftkind_givecnt2');
		var chk3 	= document.getElementById('tmpgiftkind_givecnt3');

		if ((chk2.checked == true) || (chk3.checked == true) || (document.all.gift_type.value == 1)) {
			document.all.gift_range1.readOnly=true;
			document.all.gift_range2.readOnly=true;
			document.all.giftkind_cnt.readOnly=true;
			document.all.gift_range1.style.backgroundColor='#E6E6E6';
			document.all.gift_range2.style.backgroundColor='#E6E6E6';
			document.all.giftkind_cnt.style.backgroundColor='#E6E6E6';

			if (document.all.gift_type.value == 1) {
				document.all.giftkind_cnt.readOnly=false;
				document.all.giftkind_cnt.style.backgroundColor='';
			}
		} else {
			document.all.gift_range1.readOnly=false;
			document.all.gift_range2.readOnly=false;
			document.all.giftkind_cnt.readOnly=false;
			document.all.gift_range1.style.backgroundColor='';
			document.all.gift_range2.style.backgroundColor='';
			document.all.giftkind_cnt.style.backgroundColor='';

			return;
		}

		if (chk2.checked == true) {
			// 1+1
			document.all.gift_range1.value=1;
			document.all.gift_range2.value=0;
			document.all.giftkind_cnt.value=1;
		} else if (chk3.checked == true) {
			// 1:1
			document.all.gift_range1.value=1;
			document.all.gift_range2.value=0;
			document.all.giftkind_cnt.value=1;
		} else {
			// ����
		}
	}

	function jsChkGiftType(iVal){
		if(iVal==1){
			document.all.gift_range1.readOnly=true;
			document.all.gift_range2.readOnly=true;
			document.all.gift_range1.style.backgroundColor='#E6E6E6';
			document.all.gift_range2.style.backgroundColor='#E6E6E6';

			document.all.gift_range1.value=0;
			document.all.gift_range2.value=0;
		}else{
			document.all.gift_range1.readOnly=false;
			document.all.gift_range2.readOnly=false;
			document.all.gift_range1.style.backgroundColor='';
			document.all.gift_range2.style.backgroundColor='';
		}

		if (iVal != 3) {
			// ���������� ������ �ƴϸ� ���������Ѵ�.
			var chk2 	= document.getElementById('tmpgiftkind_givecnt2');
			var chk3 	= document.getElementById('tmpgiftkind_givecnt3');

			chk2.checked=false;
			chk3.checked=false;

			jsResetGiftNo();
		}
	}

	// 1+1 ,1:1 üũ
	function jsCheckKT(ev,ch){

		var chk 	= document.getElementById(ev);
		var chftf 	= chk.checked;
		var chk2 	= document.getElementById('tmpgiftkind_givecnt2');
		var chk3 	= document.getElementById('tmpgiftkind_givecnt3');

		chk2.checked=false;
		chk3.checked=false;

		if (document.all.gift_scope.value != 5) {
			alert("��������� ��ϻ�ǰ���� �����ؾ߸� üũ�� �� �ֽ��ϴ�.");
			return;
		}

		if (document.all.gift_type.value != 3) {
			alert("���������� �������� �����ؾ߸� üũ�� �� �ֽ��ϴ�.");
			return;
		}

		chk.checked=chftf;

		if(chftf){
			document.frmReg.giftkind_givecnt.value= chk.value;
		}else{
			document.frmReg.giftkind_givecnt.value=0;
		}
		jsResetGiftNo();
	}

	function jsChkLimit(){
		if(document.frmReg.chkLimit.checked){
			document.all.giftkind_limit.readOnly=false;
			document.all.giftkind_limit.style.backgroundColor='';
		}else{
			document.all.giftkind_limit.readOnly=true;
			document.all.giftkind_limit.style.backgroundColor='#E6E6E6';
			document.frmReg.giftkind_limit.value = "";
		}
	}

	// ����ǰ��ϳ��� ��������
	function jsImport(ec){
		var pp = window.open('/admin/offshop/gift/popGiftList.asp?eC='+ec,'popim','scrollbars=yes,resizable=yes,width=900,height=600');

	}

	// ����ǰ �˻�
	function jsSearchGiftItem(){
		var winkind;
		winkind = window.open('popgiftKindReg.asp?giftkind_name='+ urlencode(document.frmReg.giftkind_name.value),'popkind','width=800, height=500,scrollbars=yes,resizable=yes');
		// winkind = window.open('popgiftKindReg.asp?giftkind_name='+ document.frmReg.giftkind_name.value,'popkind','width=800, height=500,scrollbars=yes,resizable=yes');
		winkind.focus();
	}

	// ��ϻ�ǰ �˻�
	function jsSearchTargetItem(itemgubun){
		var winkind;
		winkind = window.open('popTargetItemReg.asp?shopitemname='+document.frmReg.shopitemname.value + '&itemgubun=' + itemgubun,'jsSearchTargetItem','width=800, height=600,scrollbars=yes,resizable=yes');
		winkind.focus();
	}

	function urlencode(plaintxt) {
		   return escape(plaintxt).replace("+","%2C");
	}

	window.onload = function() {
		jsChkGiftScope(document.all.gift_scope.value);
	}

function popUploadGiftItemimage(frm) {
	var mode, imagekind, pk;

	if (frm.gift_code.value == "") {
		alert("���� ����ǰ ������ �����ϼ���");
		return;
	}

	if (frm.gift_img.value == "") {
		mode = "addimage";
	} else {
		mode = "editimage";
	}

	imagekind = "mobilegiftitemimage";
	pk = frm.gift_code.value;


	var popwin = window.open("/common/pop_upload_image.asp?mode=" + mode + "&imagekind=" + imagekind + "&pk=" + pk + "&50X50=Y","popUploadGiftItemimage","width=390 height=120 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<table width="100%" border="0" align="left" class="a" cellpadding="0" cellspacing="1"  >
<form name="frmReg" method="post" action="giftProc.asp">
<input type="hidden" name="mode" value="giftedit">
<input type="hidden" name="gift_code" value="<%=gift_code%>">
<input type="hidden" name="giftkind_givecnt" value="0">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�̺�Ʈ�ڵ�</td>
			<td bgcolor="#FFFFFF">
				<%=evt_code%>
			</td>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">���뼥</td>
			<td bgcolor="#FFFFFF">
				<%= shopid %>(<%= shopname %>)
			</td>
		</tr>
		<tr height="25">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�̺�Ʈ��</td>
			<td bgcolor="#FFFFFF">
				<%=evt_name%>
			</td>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�̺�Ʈ�Ⱓ</td>
			<td bgcolor="#FFFFFF">
				������ : <%= evt_startdate %> ~ ������ : <%= evt_enddate %>
			</td>
		</tr>
		<tr height="25">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
			<td bgcolor="#FFFFFF">
				<%=replace(evt_sStateDesc,"���¿���","����")%>
			</td>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�귣��</td>
			<td bgcolor="#FFFFFF">
				<%= brand %>
			</td>
		</tr>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center"></td>
			<td bgcolor="#FFFFFF" colspan="3"><input type="button" class="button" value="���� ����ǰ���� ��������" onClick="jsImport('<%= evt_code %>');"></td>
		</tr>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> ����</td>
			<td bgcolor="#FFFFFF" width="400">
				<font color="gray"><%=gift_name%></font><input type="hidden" name="gift_name" value="<%=gift_name%>">
			</td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> �Ⱓ</td>
			<input type="hidden" name="gift_startdate" value="<%=gift_startdate%>">
			<input type="hidden" name="gift_enddate" value="<%=gift_enddate%>">
			<td bgcolor="#FFFFFF">
				<font color="gray">
				������ :
				<%=gift_startdate%>
				~ ������ :
				<%=gift_enddate%>
				</font>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�������</td>
			<td bgcolor="#FFFFFF">
				<%sbGetOptCommonCodeArr_off "gift_scope", gift_scope, isregstate, True, "onchange='jsChkGiftScope(this.value);'" %>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����</td>
			<input type="hidden" name="gift_status" value="<%=gift_status%>">
			<td bgcolor="#FFFFFF">
				<% if gift_code = "" then %>
						<font color="gray"><%=replace(sStateDesc,"���¿���","����")%></font>
				<% else %>
						<%=replace(fnGetCommCodeArrDesc_off(sStateDesc,gift_status),"���¿���","����")%>
				<% end if %>
				<input type="hidden" name="opendate" value="<%=opendate%>">
				<input type="hidden" name="closedate" value="<%=closedate%>">
				<%IF opendate <> "" THEN%><span style="padding-left:10px;">����ó����: <%=opendate%></span><%END IF%>
				<%IF closedate <> "" THEN%><br><span style="padding-left:42px;">����ó����: <%=closedate%></span><%END IF%>
			</td>
		</tr>

		<tr id="showmakerid" style="display:none">
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">��Ϻ귣��</td>
			<td  width="400" bgcolor="#FFFFFF" colspan="3">
				<% drawSelectBoxDesignerwithName "makerid", makerid %>
			</td>
		</tr>

		<tr id="showitemid" style="display:none">
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">��ϻ�ǰ</td>
			<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
			<input type="hidden" name="shopitemid" value="<%= shopitemid %>">
			<input type="hidden" name="itemoption" value="<%= itemoption %>">
			<td bgcolor="#FFFFFF" colspan="3">
				<input type="text" name="shopitemname"  value="<%= shopitemname %>" size="40" maxlength="60" style="background-color:#E6E6E6;" readonly>
				<input type="button" class="button" value="ON ��ǰ�˻�" onClick="jsSearchTargetItem('10');">
				<input type="button" class="button" value="OFF ��ǰ�˻�" onClick="jsSearchTargetItem('90');">
			</td>
		</tr>

		<tr id="showaddcondition" style="display:none">
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">���������������</td>
			<td bgcolor="#FFFFFF" colspan="3">
				<input type="text" name="gift_scope_add" size="30" value="<%= gift_scope_add %>">
				* ���� : 2011�⵵ ���б� ���Ի� ����, �ֺ�����
			</td>
		</tr>

		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">��������</td>
			<td width="400" bgcolor="#FFFFFF">
				<%sbGetOptCommonCodeArr_off "gift_type", gift_type, isregstate, True,"onchange='jsChkGiftType(this.value);'" %>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">��������</td>
			<td bgcolor="#FFFFFF">
				<% if gift_code = "" then %>
					<input type="text" name="gift_range1" size="10" style="text-align:right" value="0"> �̻�
					~ <input type="text" name="gift_range2" size="10" style="text-align:right" value="0"> �̸�
				<% else %>
					<input type="text" name="gift_range1" size="10" style="text-align:right;<%IF gift_type= "1" THEN%>background-color:#E6E6E6; readonly<%ELSE%>"<%END IF%> value="<%=gift_range1%>"> �̻�
					~ <input type="text" name="gift_range2" size="10" style="text-align:right;<%IF gift_type= "1" THEN%>background-color:#E6E6E6; readonly<%ELSE%>"<%END IF%> value="<%=gift_range2%>"> �̸�
				<% end if %>
				(ex. 20�� �̻�: 20~0)
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">����ǰ��</td>
			<td  width="400" bgcolor="#FFFFFF">
				<input type="hidden" name="gift_itemgubun" value="<%= gift_itemgubun %>">
				<input type="hidden" name="gift_shopitemid" value="<%= gift_shopitemid %>">
				<input type="hidden" name="gift_itemoption" value="<%= gift_itemoption %>">
				<input type="hidden" name="giftkind_code" value="<%=giftkind_code%>">
				<input type="text" name="giftkind_name"  value="<%=giftkind_name%>" size="40" maxlength="60" style="background-color:#E6E6E6;" readonly>
				<input type="button" class="button" value="����ǰ�˻�" onClick="jsSearchGiftItem();">
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">����ǰ����</td>
			<td bgcolor="#FFFFFF">
					<input type="text" name="giftkind_cnt" size="4" maxlength="10" value="<%=giftkind_cnt%>" style="text-align:right;"> ����
						<label title="���ϻ�ǰ���� 1+1" ><input type="checkbox" name="tmpgiftkind_givecnt2" onclick="jsCheckKT('tmpgiftkind_givecnt2');"  value="2" <%IF CStr(giftkind_type) = "2" THEN%>checked<%END IF%>>1+1(���ϻ�ǰ) </label>
						<label title="�ٸ���ǰ���� 1:1" ><input type="checkbox" name="tmpgiftkind_givecnt3" onclick="jsCheckKT('tmpgiftkind_givecnt3');" value="3" <%IF CStr(giftkind_type) = "3" THEN%>checked<%END IF%>>1:1(�ٸ���ǰ) </label>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����ǰ��������</td>
			<td bgcolor="#FFFFFF">
				<input type="checkbox" name="chkLimit" onClick="jsChkLimit();" <%IF giftkind_limit > 0 THEN%>checked<%END IF%> <% if (shopid = "") or (shopid = "all") then %>disabled<% end if %>>����
				<input type="text" name="giftkind_limit" size="4" value="<%=giftkind_limit%>" style="text-align:right;" <%IF giftkind_limit = 0 THEN%>style="background-color:#E6E6E6;" readonly<%END IF%>> ��(�������� ���� ��쿡�� �Է�)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�����Ҹ����</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="giftkind_limit_sold" size="4" value="<%=giftkind_limit_sold%>" style="text-align:right;">
			</td>
		</tr>
		<% if gift_code <> "" then %>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�������</td>
			<td bgcolor="#FFFFFF" colspan=3>
				<input type="radio" name="gift_using" value="Y" <%IF gift_using = "Y" THEN%>checked<%END IF%>>���
				<input type="radio" name="gift_using" value="N" <%IF gift_using = "N" THEN%>checked<%END IF%>>������
			</td>
		</tr>
		<% end if %>
		<tr height="25">
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">���������</td>
			<td bgcolor="#FFFFFF" colspan=3>
				<%= receiptstring %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">����ǰ��<br>(����Ͼ�ǥ��)</td>
			<td  width="400" bgcolor="#FFFFFF">
				<input type="text" name="gift_itemname"  value="<%=gift_itemname%>" size="40" maxlength="60">
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">�̹���<br>(����Ͼ�ǥ��)</td>
			<td bgcolor="#FFFFFF">
				<% if (gift_code <> "") then %>
					<% if (gift_img <> "") then %>
						<img src="<%= gift_img_50X50_url %>"><br>
						<img src="<%= gift_imgurl %>"><br>
						<input type="button" class="button" value="�����ϱ�" onclick="popUploadGiftItemimage(frmReg)">
					<% else %>
						<input type="button" class="button" value="����ϱ�" onclick="popUploadGiftItemimage(frmReg)">
					<% end if %>
				<% end if %>
				<input type="hidden" name="gift_img" value="<%= gift_img %>">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<input type="button" onclick="jsSubmitGift();" value="����" class="button">
		<input type="button" onclick="history.back();" value="���" class="button">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->