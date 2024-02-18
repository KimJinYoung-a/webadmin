<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 장비자산관리
' History : 2008년 06월 27일 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
'// deprecated : property_gubun, part_sn
'// new : accountGubun, department_id
dim idx ,equip_code ,equip_gubun ,equip_name ,equip_spec ,equip_mainimage ,property_gubun, BIZSECTION_CD, BIZSECTION_NM
dim manufacture_sn ,manufacture_company ,manufacture_manager ,manufacture_tel ,buy_company_name, info_gubun_dic, paymentrequestidx
dim buy_date ,buy_cost ,buy_vat ,buy_sum ,using_userid ,using_date ,out_date ,state, i, info_importance_C, info_importance_I
dim durability_month ,etc ,part_sn ,regdate ,lastupdate ,reguserid ,lastuserid ,isusing, accountGubun, department_id, locate_gubun
dim monthlyDeprice, remainValue201412, info_gubun, info_gubun_display, info_gubun_value, info_importance_A, accountassetcode
	idx = requestCheckVar(request("idx"),10)
	property_gubun = requestCheckVar(request("property_gubun"),10)

dim oequip
set oequip = new CEquipment
	oequip.FRectIdx = idx

	if idx <> "" then
		oequip.getOneEquipment
	end if

if oequip.ftotalcount > 0 then
	accountassetcode = oequip.FOneItem.faccountassetcode
	paymentrequestidx = oequip.FOneItem.fpaymentrequestidx
	idx = oequip.FOneItem.fidx
	equip_code = oequip.FOneItem.fequip_code
	equip_gubun = oequip.FOneItem.fequip_gubun
	equip_name = oequip.FOneItem.fequip_name
	equip_spec = oequip.FOneItem.fequip_spec
	equip_mainimage = oequip.FOneItem.fequip_mainimage
	property_gubun = oequip.FOneItem.fproperty_gubun
	manufacture_sn = oequip.FOneItem.fmanufacture_sn
	manufacture_company = oequip.FOneItem.fmanufacture_company
	manufacture_manager = oequip.FOneItem.fmanufacture_manager
	manufacture_tel = oequip.FOneItem.fmanufacture_tel
	buy_company_name = oequip.FOneItem.fbuy_company_name
	buy_date = Left(oequip.FOneItem.fbuy_date,10)
	buy_cost = oequip.FOneItem.fbuy_cost
	buy_vat = oequip.FOneItem.fbuy_vat
	buy_sum = oequip.fOneItem.fbuy_sum
	using_userid = oequip.FOneItem.fusing_userid
	using_date = oequip.FOneItem.fusing_date
	out_date = Left(oequip.fOneItem.fout_date,10)
	state = oequip.FOneItem.fstate
	durability_month = oequip.FOneItem.fdurability_month
	etc = oequip.FOneItem.fetc
	part_sn = oequip.FOneItem.fpart_sn
	accountGubun = oequip.FOneItem.FaccountGubun
	department_id = oequip.FOneItem.Fdepartment_id
	locate_gubun = oequip.FOneItem.Flocate_gubun
	BIZSECTION_CD = oequip.FOneItem.FBIZSECTION_CD
	BIZSECTION_NM = oequip.FOneItem.FBIZSECTION_NM
	regdate = oequip.FOneItem.fregdate
	lastupdate = oequip.FOneItem.flastupdate
	reguserid = oequip.FOneItem.freguserid
	lastuserid = oequip.FOneItem.flastuserid
	isusing = oequip.FOneItem.fisusing
	monthlyDeprice = oequip.FOneItem.FmonthlyDeprice
	remainValue201412 = oequip.FOneItem.FremainValue201412
	info_gubun = oequip.FOneItem.Finfo_gubun
	info_importance_C = oequip.FOneItem.Finfo_importance_C
	info_importance_I = oequip.FOneItem.Finfo_importance_I
	info_importance_A = oequip.FOneItem.Finfo_importance_A

	if info_gubun <> "" then
		set info_gubun_dic = oequip.FOneItem.Finfo_gubun_dic
	end if
end if

dim oinfoequipgubun
set oinfoequipgubun = new CEquipment
	oinfoequipgubun.FPageSize = 200
	oinfoequipgubun.FCurrPage = 1
	oinfoequipgubun.getInfoEquipmentGubunList

dim olog
set olog = new CEquipment
	olog.FPageSize = 50
	olog.FCurrPage = 1
	olog.frectidx = idx

	if olog.frectidx <> "" then
		olog.getEquipmentlogList
	end if

if state = "" then state = "1"

%>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

//저장
function regEquip(frm){

	<% if (idx = "") then %>
	if (frm.accountGubun.value.length<1){
		alert('물품구분을 선택하세요.');
		frm.accountGubun.focus();
		return;
	}
	<% end if %>

	if (frm.state.value.length<1){
		alert('상태값을 선택하세요.');
		frm.state.focus();
		return;
	}

	if (frm.buy_sum.value.length<1){
		alert('구매가격을 입력하세요.');
		frm.buy_sum.focus();
		return;
	}

	if (frm.buy_cost.value.length<1){
		alert('공급가를 선택하세요.');
		frm.buy_cost.focus();
		return;
	}

	if (frm.equip_gubun.value.length<1){
		alert('장비구분을 선택하세요.');
		frm.equip_gubun.focus();
		return;
	}

	if (frm.state.value == '5'){
		if (frm.out_date.value.length<1){
			alert('폐기날짜를 선택하세요.');
			frm.out_date.focus();
			return;
		}
	}

	if (frm.isusing.value.length<1){
		alert('사용여부를 선택하세요.');
		frm.isusing.focus();
		return;
	}

	// 83090, 소모품 은 수정가능(2016-04-06, skyer9)
	<% if (Left(Now(), 7) > Left(regdate, 7)) and (CStr(regdate) <> "") and (accountGubun <> "83090") then %>
	if ((frm.BIZSECTION_CD.value != frm.org_BIZSECTION_CD.value) || (frm.buy_date.value != frm.org_buy_date.value) || (frm.buy_sum.value != frm.org_buy_sum.value) || (frm.buy_cost.value != frm.org_buy_cost.value)) {
		<% if C_ADMIN_AUTH or C_MngPart or C_SYSTEM_Part or C_PSMngPart then %>
		if(!confirm("이전달 등록한 장비입니다. 손익부서, 구매일자, 구매가격, 공급가를 수정하시겠습니까?")){
			return;
		}
		<% else %>
		alert("이전달 등록한 내역은 손익부서, 구매일자, 구매가격, 공급가를 수정할 수 없습니다.");
		return;
		<% end if %>
	}

	/*
	if ((frm.org_state.value == '5') && (frm.org_out_date.value < "<%= (Left(Now(), 7) + "-01") %>")) {
		alert("이전달 폐기내역은 수정할 수 없습니다.");
		return;
	}

	if ((frm.state.value == '5') && (frm.out_date.value < "<%= (Left(Now(), 7) + "-01") %>")) {
		alert("폐기일자를 이전달로 수정할 수 없습니다.");
		return;
	}
	*/

	/*
	if ((frm.state.value != frm.org_state.value) || (frm.out_date.value != frm.org_out_date.value) || (frm.isusing.value != frm.org_isusing.value)) {
		alert("이전달 등록한 내역은 상태, 폐기일자, 사용여부를 수정할 수 없습니다.");
		return;
	}
	*/
	<% end if %>

	if (confirm('저장 하시겠습니까?')) {
		var info_gubun = document.frmreg.info_gubun.value;

		if (info_gubun != "") {
			for (var i = 0; ; i++) {
				var e = document.getElementById("info_gubun" + i);
				if (!e) {
					break;
				}
				//alert(e.name) //크롬에서 안됨?
				//if(e.name == ("info_gubun" + info_gubun)) {
					var info_gubun_idx = document.getElementById("info_gubun_idx" + i);
					var info_gubun_val = document.getElementById("info_gubun_val" + i);

					document.frmreg.info_gubun_idx_arr.value = document.frmreg.info_gubun_idx_arr.value + "__|__" + info_gubun_idx.value;
					document.frmreg.info_gubun_val_arr.value = document.frmreg.info_gubun_val_arr.value + "__|__" + info_gubun_val.value;
				//}
			}
		}
		frm.submit();
	}
}

function selectChange(comp){
	if (comp.name=="state"){
		if (comp.value=="5"){
			divstate50_5.style.display="";
		}else{
			divstate50_5.style.display="none";
		}
	}
}

//이미지 삭제
function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

//이미지 확대화면 새창으로 보여주기
function jsImgView(sImgUrl){
	var wImgView;

	wImgView = window.open('/common/equipment/pop_equipment_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

function jsSetImg(sImg, sName, sSpan){

	document.domain = '10x10.co.kr';

	var winImg;
	winImg = window.open('/common/equipment/pop_equipment_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

//자금관리부서 선택
function jsGetPart(){
	var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popGetBizOne','width=600, height=500, resizable=yes, scrollbars=yes');
	winP.focus();
}

//자금관리부서 등록
function jsSetPart(selUP, sPNM){
	document.frmreg.BIZSECTION_CD.value = selUP;
	document.frmreg.BIZSECTION_NM.value = sPNM;
}

function jsCalcMonthlyDeprice(frm) {
	if (frm.buy_date.value == "") {
		alert("구매일자를 입력하세요.");
		return;
	}

	if (frm.buy_cost.value == "") {
		alert("공급가를 입력하세요.");
		return;
	}

	if (frm.buy_date.value < "2015-01-01") {
		alert("2014년 이전 구매자산은 수기로 월감가상각비를 입력해야 합니다.");
		return;
	}

	frm.monthlyDeprice.value = Math.round(frm.buy_cost.value / 60);
}

function showHideInfoGubun() {
	var info_gubun = document.frmreg.info_gubun.value;
	var cc, ii, aa;

	var cc = document.getElementById("info_importance_C");
	var ii = document.getElementById("info_importance_I");
	var aa = document.getElementById("info_importance_A");

	if (info_gubun == "-1") {
		cc.style.display = "none";
		ii.style.display = "none";
		aa.style.display = "none";
	} else {
		cc.style.display = "";
		ii.style.display = "";
		aa.style.display = "";
	}

	for (var i = 0; ; i++) {
		var e = document.getElementById("info_gubun" + i);
		if (!e) {
			break;
		}

		if(e.name == ("info_gubun" + info_gubun)) {
			e.style.display = "";
		} else {
			e.style.display = "none";
		}
	}
}

function calcInfoImportance() {
}

//결제요청서보기
function pop_paymentrequestidx(paymentrequestidx){
	var pop_paymentrequestidx = window.open('/admin/approval/payreqList/regPayRequest.asp?ipridx='+paymentrequestidx,'pop_paymentrequestidx','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_paymentrequestidx.focus();
}

// 사용자(사용처) 제거
function fnDelUsingUser() {
	document.frmreg.username.value="";
	document.frmreg.using_userid.value="";
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="frmreg" method="post" action="/common/equipment/do_equipment.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mode" value="equipmentreg">
<input type="hidden" name="equip_code" value="<%= equip_code %>">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="2">
				* 기본정보
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="F4F4F4">물품코드</td>
			<td>
				<%= equip_code %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td bgcolor="F4F4F4">물품구분</td>
			<td>
				<% if (idx = "") then %>
					<% drawEquipmentAccountCode "accountGubun" ,accountGubun, "" %>
				<% else %>
					<%= GetEquipmentAccountCodeName(accountGubun) %>
					<input type="hidden" name="accountGubun" value="<%= accountGubun %>">
				<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td bgcolor="F4F4F4">회계자산관리코드</td>
			<td>
				<input type="text" name="accountassetcode" value="<%= accountassetcode %>" size=20>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td bgcolor="F4F4F4">결제요청서IDX</td>
			<td>
				<input type="text" name="paymentrequestidx" value="<%= paymentrequestidx %>" size=10>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">손익부서</td>
			<td>
				<input type="text" name="BIZSECTION_CD" value="<%= BIZSECTION_CD %>" size="15"  class="text_ro"> <input type="text" name="BIZSECTION_NM" value="<%= BIZSECTION_NM %>" class="text_ro" size="15">
				<input type="hidden" name="org_BIZSECTION_CD" value="<%= BIZSECTION_CD %>">
				<a href="javascript:jsGetPart();"> <img src="/images/icon_search.jpg" border="0"></a>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="F4F4F4">구매일자</td>
			<td>
				<input type="text" id="buyDt" name="buy_date" readonly size="11" maxlength="10" value="<%= buy_date %>" class="text_ro" style="text-align:center;" />
				<input type="hidden" name="org_buy_date" value="<%= buy_date %>">
				<img src="/images/calicon.gif" align="absmiddle" border="0" id="btnBuyDt" style="cursor:pointer;" />
				<script type="text/javascript">
					var CAL_BuyDate = new Calendar({
						inputField : "buyDt", trigger    : "btnBuyDt",
						bottomBar: true, dateFormat: "%Y-%m-%d",
						onSelect: function() {
							this.hide();
						}
					});
				</script>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">구매가격</td>
			<td>
				<input type="text" class="text" name="buy_sum" value="<%= buy_sum %>" size="13" maxlength="13">
				<input type="hidden" name="org_buy_sum" value="<%= buy_sum %>">
				(부가세 포함가)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">공급가</td>
			<td>
				<input type="text" class="text" name="buy_cost" value="<%= buy_cost %>" size="12" maxlength="13">
				<input type="hidden" name="org_buy_cost" value="<%= buy_cost %>">
				(부가세 제외)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">월감가상각</td>
			<td>
				<input type="text" class="text" name="monthlyDeprice" value="<%= monthlyDeprice %>" size="10" maxlength="13">
				<input type="button" class="button" value="자동계산" onClick="jsCalcMonthlyDeprice(frmreg);">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">2014년잔존가치</td>
			<td>
				<input type="text" class="text" name="remainValue201412" value="<%= remainValue201412 %>" size="12" maxlength="13">
				(2014년까지의 구매자산만)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="F4F4F4">상태</td>
			<td>
				<% DrawEquipMentGubun "50","state",state ," onchange='selectChange(frmreg.state)'" %>
				<input type="hidden" name="org_state" value="<%= state %>">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" id="divstate50_5" style="display:none;">
			<td bgcolor="F4F4F4">폐기정보</td>
			<td>
				폐기(반납)날짜 : <input type="text" id="outDt" name="out_date" readonly size="11" maxlength="10" value="<%= out_date %>" class="text_ro" style="text-align:center;" />
				<input type="hidden" name="org_out_date" value="<%= out_date %>">
				<img src="/images/calicon.gif" align="absmiddle" border="0" id="btnOutDt" style="cursor:pointer;" />
				<script type="text/javascript">
					var CAL_OutDate = new Calendar({
						inputField : "outDt", trigger    : "btnOutDt",
						bottomBar: true, dateFormat: "%Y-%m-%d",
						onSelect: function() {
							this.hide();
						}
					});
				</script>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="F4F4F4">사용여부</td>
			<td>
				<% if (Left(Now(), 7) > Left(regdate, 7)) and (CStr(regdate) <> "") then %>
					<% if (Not C_ADMIN_AUTH) and (Not C_OFF_AUTH) and (Not C_MngPart) and (Not C_PSMngPart) then %>
					<%= isusing %>
					<input type="hidden" name="isusing" value="<%= isusing %>">
					<% else %>
					<% drawSelectBoxUsingYN  "isusing" ,isusing %> [관리자뷰]
					<% end if %>
				<% else %>
				<% drawSelectBoxUsingYN  "isusing" ,isusing %>
				<% end if %>
				<input type="hidden" name="org_isusing" value="<%= isusing %>">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="25">
			<td bgcolor="F4F4F4">작성일자</td>
			<td>
				<%= regdate %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="25">
			<td bgcolor="F4F4F4">최종수정</td>
			<td>
				<%= lastupdate %>
			</td>
		</tr>
		</table>
		<p>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="2">
				- 정보자산정보
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="F4F4F4">정보자산구분</td>
			<td>
				<% drawInfoEquipmentGubun "info_gubun" ,info_gubun, "onchange='showHideInfoGubun()'" %>
				<input type="hidden" name="info_gubun_idx_arr" value="">
				<input type="hidden" name="info_gubun_val_arr" value="">
			</td>
		</tr>
		<%
		info_gubun_display = ""
		if (info_gubun = "-1") then
			info_gubun_display = "none"
		end if
		%>
		<tr bgcolor="#FFFFFF" height="30" id="info_importance_C" style="display:<%= info_gubun_display %>;">
			<td width="100" bgcolor="F4F4F4">기밀성(C)</td>
			<td>
				<% drawInfoImportance "info_importance_C" ,info_importance_C, "onchange='calcInfoImportance()'" %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30" id="info_importance_I" style="display:<%= info_gubun_display %>;">
			<td width="100" bgcolor="F4F4F4">무결성(I)</td>
			<td>
				<% drawInfoImportance "info_importance_I" ,info_importance_I, "onchange='calcInfoImportance()'" %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30" id="info_importance_A" style="display:<%= info_gubun_display %>;">
			<td width="100" bgcolor="F4F4F4">가용성(A)</td>
			<td>
				<% drawInfoImportance "info_importance_A" ,info_importance_A, "onchange='calcInfoImportance()'" %>
			</td>
		</tr>
		<% if oinfoequipgubun.fresultcount > 0 then %>
		<% for i=0 to oinfoequipgubun.FResultCount - 1 %>
		<%
		''info_gubun_display, info_gubun_value
		if (info_gubun = oinfoequipgubun.FItemList(i).Finfo_gubun) then
			info_gubun_display = ""
		else
			info_gubun_display = "none"
		end if

		if (info_gubun <> "") then
			info_gubun_value = info_gubun_dic.item(CStr(oinfoequipgubun.FItemList(i).Finfo_GbnIdx))
		else
			info_gubun_value = ""
		end if

		%>
		<tr bgcolor="#FFFFFF" height="30" id="info_gubun<%= i %>" name="info_gubun<%= oinfoequipgubun.FItemList(i).Finfo_gubun %>" style="display:<%= info_gubun_display %>;">
			<td width="100" bgcolor="F4F4F4"><%= oinfoequipgubun.FItemList(i).Finfo_GbnName %></td>
			<td>
				<input type="text" class="text" id="info_gubun_val<%= i %>" value="<%= info_gubun_value %>" size="30" maxlength="30">
				<input type="hidden" id="info_gubun_idx<%= i %>" value="<%= oinfoequipgubun.FItemList(i).Finfo_GbnIdx %>">
			</td>
		</tr>
		<% next %>
		<% end if %>
		</table>
		<p>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="2">
				- 장비정보
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="F4F4F4">물품구분상세</td>
			<td>
				<% textboxEquipMentGubunNew "10","equip_gubun", "equip_gubun_name", equip_gubun," ","" %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="F4F4F4">기본이미지</td>
			<td>
				<input type="hidden" name="equip_mainimage" value="<%=equip_mainimage%>">
		   		<input type="button" name="btnBan2010" value="기본이미지" onClick="jsSetImg('<%=equip_mainimage%>','equip_mainimage','equip_mainimagediv')" class="button">
	   			<div id="equip_mainimagediv" style="padding: 5 5 5 5">
	   				<% IF equip_mainimage <> "" THEN %>
		   				<img src="<%=equip_mainimage%>" border="0" width=50 height=50 onclick="jsImgView('<%=equip_mainimage%>');" alt="누르시면 확대 됩니다">
		   				<a href="javascript:jsDelImg('equip_mainimage','equip_mainimagediv');"><img src="/images/icon_delete2.gif" border="0"></a>
		   				<br>
		   				※이미지 다운로드 방법: 이미지에서 마우스오른쪽버튼 클릭후	"다른이름으로사진저장" 누르시면 됩니다.
	   				<% END IF%>
	   			</div>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="F4F4F4">장비명</td>
			<td>
				<input type="text" class="text" name="equip_name" value="<%= equip_name %>" size="60" maxlength="60">
				(ex : 삼보 드림시스 74SC)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="F4F4F4">장비시리얼</td>
			<td>
				<input type="text" class="text" name="manufacture_sn" value="<%= manufacture_sn %>" size="60" maxlength="60">
				(ex : PN17AS , xx-xxx-xxx)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">제조사</td>
			<td>
				<input type="text" class="text" name="manufacture_company" value="<%= manufacture_company %>" size="60" maxlength="60">
				(ex : 삼성전자, LG전자)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">제조사담당자</td>
			<td>
				<input type="text" class="text" name="manufacture_manager" value="<%= manufacture_manager %>" size="32" maxlength="32">
				(ex : 홍길동대리)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">제조사연락처</td>
			<td>
				<input type="text" class="text" name="manufacture_tel" value="<%= manufacture_tel %>" size="16" maxlength="16">
				(ex : xxx-xxxx-xxxx)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">구매처</td>
			<td>
				<input type="text" class="text" name="buy_company_name" value="<%= buy_company_name %>" size="60" maxlength="60">
				(ex : 삼성몰, 인터파크, DELL코리아)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">상세사양</td>
			<td>
				<textarea class="textarea" cols="60" rows="5" name="equip_spec"><%= equip_spec %></textarea>
			</td>
		</tr>
		</table>
		<p>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
		<tr bgcolor="#FFFFFF">
			<td colspan="2">
				- 사용정보
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">위치</td>
			<td >
				<% textboxEquipmentGubunNew "30", "locate_gubun", "locate_gubun_name", locate_gubun," ","" %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="F4F4F4">사용자(사용처)</td>
			<td>
				<% gettenbytenuser "using_userid", using_userid, "" ,"18","" %>
				<input type="button" class="button" value="제거" onclick="fnDelUsingUser()" />
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="F4F4F4">사용일</td>
			<td>
				<input type="text" id="UseDt" name="using_date" readonly size="11" maxlength="10" value="<%= using_date %>" class="text_ro" style="text-align:center;" />
				<img src="/images/calicon.gif" align="absmiddle" border="0" id="btnUseDt" style="cursor:pointer;" />
				<script type="text/javascript">
					var CAL_UseDate = new Calendar({
						inputField : "UseDt", trigger    : "btnUseDt",
						bottomBar: true, dateFormat: "%Y-%m-%d",
						onSelect: function() {
							this.hide();
						}
					});
				</script>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="F4F4F4">기타코맨트</td>
			<td >
				<textarea class="textarea" cols="80" rows="5" name="etc"><%= etc %></textarea>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center">
	<td>
		<p>
		<input type="button" value="저장" onclick="regEquip(frmreg);" class="button">
	</td>
</tr>

</form>
</table>

<script type="text/javascript">
	selectChange(frmreg.state);
</script>

<%
'/로그리스트
if olog.fresultcount > 0 then
%>
	<!-- 리스트 시작 -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			최근 변동 검색결과 : <b><%= olog.FTotalCount %></b> ※ 50건 까지 노출 됩니다.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>장비코드</td>
		<td>회계자산<br>관리코드</td>
		<td>결제요청서<br>IDX</td>
		<td>자산구분</td>
		<td>부서</td>
		<td>사용자(사용처)</td>
		<td>장비<br>구분</td>
		<td>상태</td>
		<td>제품명</td>
		<td>사용<br>여부</td>
		<td>로그정보</td>
	</tr>
	<% if olog.FResultCount > 0 then %>
	<% for i=0 to olog.FResultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF" onMouseOver= this.style.background='f1f1f1'; onMouseOut=this.style.background='#ffffff';>
		<td width=130>
			<a href="javascript:pop_Equipmentreg('<%= olog.FItemList(i).Fidx %>');" onfocus="this.blur()">
			<%= olog.FItemList(i).Fequip_code %></a>
		</td>
		<td>
			<%= olog.FItemList(i).faccountassetcode %>
		</td>
		<td width=60>
			<a href="#" onclick="pop_paymentrequestidx('<%= olog.FItemList(i).fpaymentrequestidx %>'); return false;">
			<%= olog.FItemList(i).fpaymentrequestidx %></a>
		</td>
		<td width=80>
			<%= olog.FItemList(i).GetAccountGubunName %>
		</td>
		<td width=240>
			<%= olog.FItemList(i).FdepartmentNameFull %>
		</td>
		<td width=100>
			<%= olog.FItemList(i).fusingusername %>
			<% if olog.FItemList(i).fstatediv <> "Y" then %>
				<font color="red">[퇴사]</font>
			<% end if %>

			<% if olog.FItemList(i).Fusing_userid <> "" then %>
				<Br><%= olog.FItemList(i).Fusing_userid %>
			<% end if %>
		</td>
		<td width=100>
			<%= olog.FItemList(i).Fequip_gubun_name %>
		</td>
		<td width=80>
			<%= olog.FItemList(i).fstate_name %>
		</td>
		<td align="left">
			<%= olog.FItemList(i).Fequip_name %>
		</td>
		<td width=30>
			<%= olog.FItemList(i).fisusing %>
		</td>
		<td align="left">
			<%= olog.FItemList(i).flogregdate %>
			<br><%= olog.FItemList(i).flogreguserid %>
		</td>
	</tr>
	<% next %>

	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>

	</table>
<% end if %>

<%
set oequip = Nothing
set olog = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
