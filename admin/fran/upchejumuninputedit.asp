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
''대표 샵 아이디 : 가맹점대표 or 직영점대표 or 해외점대표중 선택.
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


''기본 매입 구분
dim DIFFCenterMWDivExists
DIFFCenterMWDivExists = False

dim DefaultItemMwDiv
DefaultItemMwDiv = GetDefaultItemMwdivByBrand(odesinger)


''대표 샵 아이디 설정 - 마진 높은기준으로.
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
    response.write "<script>alert('마진이 설정 되어 있지 않습니다. 관리자 문의 요망');</script>"
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
		alert('선택 내역이 없습니다.');
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

	if (confirm('선택 내역을 삭제 하시겠습니까?')){
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
		alert('선택 아이템이 없습니다.');
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
					alert('갯수는 정수만 가능합니다.');
					frm.baljuitemno.focus();
					return;
				}

				if (frm.baljuitemno.value.length<1){
					alert('수량을 입력하세요.');
					frm.baljuitemno.focus();
					return;
				}

				if (!IsInteger(frm.realitemno.value)){
					alert('갯수는 정수만 가능합니다.');
					frm.realitemno.focus();
					return;
				}

				if (frm.realitemno.value.length<1){
					alert('수량을 입력하세요.');
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

	if (confirm('저장 하시겠습니까?')){
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
					alert('갯수는 정수만 가능합니다.');
					frm.baljuitemno.focus();
					return;
				}

				if (frm.baljuitemno.value.length<1){
					alert('수량을 입력하세요.');
					frm.baljuitemno.focus();
					return;
				}

				if (!IsInteger(frm.realitemno.value)){
					alert('갯수는 정수만 가능합니다.');
					frm.realitemno.focus();
					return;
				}

				if (frm.realitemno.value.length<1){
					alert('수량을 입력하세요.');
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

	if (confirm('저장 하시겠습니까?')){
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
	//	alert('주문접수 상태가 아니면 수정하실 수 없습니다.');
	//	return;
	//}

	var popwin;
	var suplyer, shopid;

	if (frm.suplyer.value.length<1){
		alert('공급처를 먼저 선택하세요.');
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
	//	alert('주문접수 상태에선 수정하실 수 없습니다.');
	//	return;
	//}

	var ret = confirm('수정 하시겠습니까?');

	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}

function DelThis(frm){
	//if (jumunwait!=true){
	//	alert('주문접수 상태가 아니면 수정하실 수 없습니다.');
	//	return;
	//}

	var ret = confirm('삭제 하시겠습니까?');

	if (ret){
		frm.mode.value="deldetail";
		frm.submit();
	}
}

function DelMaster(frm){
	//if (jumunwait!=true){
	//	alert('주문접수 상태가 아니면 수정하실 수 없습니다.');
	//	return;
	//}

	var ret = confirm('삭제 하시겠습니까?');

	if (ret){
		frm.mode.value="delmaster";
		frm.submit();
	}
}

function ModiMaster(frm){
	if (frm.beasongdate!=undefined){
		frm.songjangname.value = frm.songjangdiv.options[frm.songjangdiv.selectedIndex].text;
	}

	var ret = confirm('수정 하시겠습니까?');

	if (ret){
		frm.mode.value="modimaster";
		frm.submit();
	}
}

function ReActItems(iidx,igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner){
	if (iidx!='<%= idx %>'){
		alert('주문서가 일치하지 않습니다. 다시시도해 주세요.');
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
		alert('출고일을 입력해 주세요.');
		frm.ipgodate.focus();
		if (!calendarOpen2(frm.ipgodate)) { return };
	}

	if (frm.beasongdate!=undefined){
		frm.songjangname.value = frm.songjangdiv.options[frm.songjangdiv.selectedIndex].text;
	}

	var ret = confirm('출고처리 하시겠습니까?');

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

	var ret2 = confirm('입고일 : ' + frmMaster.ipgodate.value + ' OK?');
	if (!ret2) return;

	var idivcode = getCheckboxValue(frmMaster,'divcode');

	if (idivcode=="121"){
		imsg = "[온라인위탁재고->가맹점용위탁] 인경우 \r\n온라인 내역에 출고로 잡히고 \r\n가맹점으로 위탁입고됩니다. \r\n입고 확정으로 진행 하시겠습니까?";
	}else if(idivcode=="131"){
		imsg = "[온라인위탁재고->가맹점용매입] 인경우 \r\n온라인 내역에 출고로 잡히고 \r\n가맹점으로 매입입고됩니다. \r\n입고 확정으로 진행 하시겠습니까?";
	}else if(idivcode=="201"){
		imsg = "[온라인매입재고->가맹점용매입] 인경우 \r\n온라인 내역에 출고로 잡히고 \r\n가맹점으로 매입입고됩니다. \r\n입고 확정으로 진행 하시겠습니까?";
	}else{
		imsg = " 입고 확정으로 진행 하시겠습니까?";
	}

	var ret = confirm(imsg);

	if (ret){

		frmMaster.mode.value= "franupcheipgofinish";
		frmMaster.targetid.value= frmMaster.suplyer.value;
		frmMaster.submit();
	}
}

function DelAlink(frm,alinkcode){
	if (confirm('관련된 입출고 내역을 삭제 하시겠습니까?')){
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

	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red"><strong>주문정보</strong></font>
				        &nbsp;
				        <b>[ <%= ojumunmaster.FOneItem.Fbaljucode %> ]</b>
				        &nbsp;
				        <% if (Not IsNULL(ojumunmaster.FOneItem.FALinkCode)) and (ojumunmaster.FOneItem.FALinkCode<>"") then %>
							관련입출:<%= ojumunmaster.FOneItem.FALinkCode %>
							<% if (not IsNULL(ojumunmaster.FOneItem.Fipchuldeldt)) then %>
								<font color="red">삭제됨</font>
							<% end if %>
							&nbsp;총소비가:<%= FormatNumber(ojumunmaster.FOneItem.Fipchulsellcash,0) %>
							<!-- &nbsp;총공급가:<%= FormatNumber(ojumunmaster.FOneItem.Fipchulsuplycash,0) %> -->
							&nbsp;총매입가:<%= FormatNumber(ojumunmaster.FOneItem.Fipchulbuycash,0) %>
							<!-- 관련 입출고 삭제기능 없앰. (2011-11-28 eastone)
							<input type="button" class="button" value="관련 입출고 삭제" onClick="DelAlink(frmMaster,'<%= ojumunmaster.FOneItem.FALinkCode %>');">
							-->
						<% end if %>

				    </td>
				    <td align="right">
						<input type="button" class="button" value="목록으로 이동" onclick="document.location='upchejumunlist.asp?page=<%= opage %>&designer=<%= odesinger %>&statecd=<%= ostatecd %>'">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- 상단바 끝 -->

	<tr bgcolor="#FFFFFF">
		<td width="110" bgcolor="<%= adminColor("tabletop") %>" >공급처(브랜드)</td>
		<td width="400">
			<input type=hidden name="suplyer" value="<%= ojumunmaster.FOneItem.Ftargetid %>">
			<%= ojumunmaster.FOneItem.Ftargetid %>&nbsp;(<%= ojumunmaster.FOneItem.Ftargetname %>)
		</td>
		<td width="110" bgcolor="<%= adminColor("tabletop") %>" >발주처(주문자)</td>
		<td><%= ojumunmaster.FOneItem.Fbaljuid %>&nbsp;(<%= ojumunmaster.FOneItem.Fbaljuname %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >주문일시</td>
		<td><%= ojumunmaster.FOneItem.Fregdate %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" >입고요청일</td>
		<td>
			<input type="text" class="text" name="yyyymmdd" value="<%= yyyymmdd %>" size=12 readonly ><a href="javascript:calendarOpen(frmMaster.yyyymmdd);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">진행상태</td>
		<td colspan="3">
			<input type=radio name="statecd" value="0" <% if ojumunmaster.FOneItem.FStatecd="0" then response.write "checked" %> >주문접수
			<input type=radio name="statecd" value="1" <% if ojumunmaster.FOneItem.FStatecd="1" then response.write "checked" %> >업체주문확인
			<input type=radio name="statecd" value="5" <% if ojumunmaster.FOneItem.FStatecd="5" then response.write "checked" %> >업체배송준비
			<input type=radio name="statecd" value="7" <% if ojumunmaster.FOneItem.FStatecd="7" then response.write "checked" %> >업체출고완료
			<input type=radio name="statecd" value="8" <% if ojumunmaster.FOneItem.FStatecd="8" then response.write "checked" %> >입고대기(도착완료)
			<% if ojumunmaster.FOneItem.FStatecd="9" then %>
			<input type=radio name="statecd" value="9" <% if ojumunmaster.FOneItem.FStatecd="9" then response.write "checked" %> >입고완료
				<% if (not IsNULL(ojumunmaster.FOneItem.Fipchuldeldt)) or (IsNULL(ojumunmaster.FOneItem.Falinkcode))  then %>
				&nbsp;<input type="button" class="button" value="상태변경" onClick="ModiMaster(frmMaster)">
				<% else %>
				&nbsp;<input type="button" class="button" value="상태변경" onClick="alert('관련 입출고 삭제후 사용가능합니다.')">
				<% end if %>
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >운송장입력</td>
		<td>
			택배사 : <% drawSelectBoxDeliverCompany "songjangdiv", ojumunmaster.FOneItem.Fsongjangdiv %>
			운송장번호: <input type="text" class="text" name="songjangno" size=14 maxlength=16 value="<%= ojumunmaster.FOneItem.Fsongjangno %>" >
			<input type="hidden" name="songjangname" value="">
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" >출고일</td>
		<td>
			<input type="text" class="text" name="beasongdate" value="<%= ojumunmaster.FOneItem.Fbeasongdate %>" size=12 readonly ><a href="javascript:calendarOpen(frmMaster.beasongdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
		</td>
	</tr>


	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >매입구분</td>
		<td colspan="3">
			<input type=radio name="divcode" value="101" <% if (ojumunmaster.FOneItem.Fdivcode="101") or (IsNULL(ojumunmaster.FOneItem.Fdivcode) and (DefaultItemMwDiv="M")) then response.write "checked" %> >
			<% if (DefaultItemMwDiv="M") then %>
			<b>가맹점용 개별매입</b>
			<% else %>
			가맹점용 개별매입
			<% end if %>
			&nbsp;
			<input type=radio name="divcode" value="111" <% if (ojumunmaster.FOneItem.Fdivcode="111") or (IsNULL(ojumunmaster.FOneItem.Fdivcode) and (DefaultItemMwDiv="W"))  then response.write "checked" %> >
			가맹점용 개별위탁
			&nbsp;&nbsp;
			<% if ojumunmaster.FOneItem.FStatecd="8" then %>
			<input type="button" class="button" value="입고처리" onclick="IpgoFinish()">
			<% end if %>
			&nbsp;&nbsp;
			온라인 : <%= oupchemwinfo.FOneItem.GetOnlineMwDivName %>&nbsp;<%= oupchemwinfo.FOneItem.GetOnlineDefaultmargine %>%
			&nbsp;가맹점: <%= oupchemwinfo.FOneItem.GetfranChargeDivName %>&nbsp;<%= oupchemwinfo.FOneItem.GefranDefaultmargine %>%
		</td>
	</tr>


	<% if (ojumunmaster.FOneItem.FStatecd="6") then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">출고일</td>
		<td colspan="3"><input type=text name="ipgodate" value="<%= ojumunmaster.FOneItem.Fipgodate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.ipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		(재고 정산과 관련있음)

		</td>
	</tr>
	<% elseif (ojumunmaster.FOneItem.FStatecd>6) then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">입고일</td>
		<td colspan="3"><input type="text" class="text" name="ipgodate" value="<%= ojumunmaster.FOneItem.Fipgodate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.ipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		(재고 정산과 관련있음)
		</td>
	</tr>
	<% end if %>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">소비자가합계(주문)</td>
		<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsellcash,0) %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">매입공급가합계(주문)</td>
		<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunbuycash,0) %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">소비자가합계(확정)</td>
		<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsellcash,0) %></b></td>
		<td bgcolor="<%= adminColor("tabletop") %>">매입공급가합계(확정)</td>
		<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalbuycash,0) %></b></td>
	</tr>

	<!-- 샵별 출고가는 다를수 있는데...어떤 데이타를 표시한건지
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#DDDDFF" width=100>총 공급가</td>
		<td colsapn="3"><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsuplycash,0) %> / <%= FormatNumber(ojumunmaster.FOneItem.Fjumunsuplycash,0) %></td>
	</tr>
	-->

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">주문브랜드</td>
		<td colspan="3"><textarea class="textarea_ro" cols=80 rows=3 readonly><%= ojumunmaster.FOneItem.FBrandList %></textarea></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">기타요청사항</td>
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
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red"><strong>상세내역</strong></font>
				        &nbsp;
			        	<font color="#FF0000">텐배</font>&nbsp;
			        	<font color="#000000">업배</font>&nbsp;
			        	<font color="#0000FF">오프전용</font>
				    </td>
				    <td align="right">
						총건수:  <%= ojumundetail.FResultCount %>
			        	&nbsp;
			        	<% if not isfixed then %>
							<input type="button" class="button" value="선택내역삭제" onClick="DelArr()">
						<% end if %>
						<% if not isfixed then %>
							<input type="button" class="button" value="상품추가" onclick="AddItems(frmMaster)">
						<% end if %>

					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- 상단바 끝 -->

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="20"><!-- <input type="checkbox"  name="ckall" onClick="AnSelectAllFrame(this.checked)"> --></td>
	    <td width="50">이미지</td>
		<td width="80">바코드</td>
		<td>브랜드ID</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="50">주문시<br>판매가</td>
		<td width="50">매입가</td>
		<td width="30">마진</td>
		<td width="60">주문<br>수량</td>
		<td width="60">확정<br>수량</td>
		<td width="60">확정<br>수량</td>
		<td width="30">센터<br>매입<br>구분</td>
		<% if isfixed then %>
		<td >비고</td>
		<!-- td width="30">개별<br>입고</td -->
		<% else %>
		<td width="90">비고</td>
		<!-- td width="30">개별<br>입고</td -->
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
		<div ><font color=red>온:<%= FormatNumber(ojumundetail.FItemList(i).Fonlinesellcash,0) %></font></div>
		<% end if %>
	    </td>

		<input type="hidden" name="suplycash" value="<%= ojumundetail.FItemList(i).Fsuplycash %>">
		<td align=right>
			<input type="text" class="text" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>" size="7" maxlength="9" style="text-align:right">

			<% if (ojumundetail.FItemList(i).Fbuycash<>ojumundetail.FItemList(i).Fonlinebuycash) and ((ojumundetail.FItemList(i).FItemDefaultMwDiv="W") and (ojumundetail.FItemList(i).FoffChargeDiv="4")) then %>
			<div ><font color=red>온:<%= ojumundetail.FItemList(i).Fonlinebuycash %></font></div>
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
		<td align="center">총계</td>
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
			<b>입고 완료된 내역은 수정 하실 수 없습니다.</b>
		<% else %>
			<input type="button" class="button" value="전체저장" onclick="SaveALL()">
			&nbsp;
			<input type="button" class="button" value="전체삭제" onclick="DelMaster(frmMaster)">
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
    alert('센터 매입구분이 일치하지 않습니다. - 관리자 문의 요망 ');
</script>
<% end if  %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
