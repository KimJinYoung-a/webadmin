<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 수기 매출 관리
' History : 2012.08.07 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<%
dim acURL , shopid ,cnt ,cnt2 ,j ,i ,isPreExists ,PriceEditEnable , menupos ,regdate
dim itemgubunarr ,itemidarr ,itemoptionarr ,itemnamearr ,itemoptionnamearr ,sellcasharr ,suplycasharr ,shopbuypricearr
dim itemnoarr ,makeridarr ,extbarcodearr
dim itemgubunarr2 ,itemidarr2 ,itemoptionarr2 ,itemnamearr2 ,itemoptionnamearr2
dim sellcasharr2 ,suplycasharr2 ,shopbuypricearr2 ,itemnoarr2 ,makeridarr2 ,extbarcodearr2
dim itemgubunarr3 ,itemidarr3 ,itemoptionarr3 ,itemnamearr3 ,itemoptionnamearr3 ,sellcasharr3 ,suplycasharr3
dim shopbuypricearr3 ,itemnoarr3 ,makeridarr3 ,extbarcodearr3
	itemgubunarr = request("itemgubunarr")
	itemidarr	= request("itemidarr")
	itemoptionarr = request("itemoptionarr")
	itemnamearr		= request("itemnamearr")
	itemoptionnamearr = request("itemoptionnamearr")
	sellcasharr = request("sellcasharr")
	suplycasharr = request("suplycasharr")
	shopbuypricearr = request("shopbuypricearr")
	itemnoarr = request("itemnoarr")
	makeridarr = request("makeridarr")
	extbarcodearr = request("extbarcodearr")
	itemgubunarr2 = request("itemgubunarr2")
	itemidarr2	= request("itemidarr2")
	itemoptionarr2 = request("itemoptionarr2")
	itemnamearr2	= request("itemnamearr2")
	itemoptionnamearr2 = request("itemoptionnamearr2")
	sellcasharr2 = request("sellcasharr2")
	suplycasharr2 = request("suplycasharr2")
	shopbuypricearr2 = request("shopbuypricearr2")
	itemnoarr2 = request("itemnoarr2")
	makeridarr2 = request("makeridarr2")
	extbarcodearr2 = request("extbarcodearr2")
	shopid = requestCheckVar(request("shopid"),32)
	menupos = requestCheckVar(request("menupos"),10)
	regdate = requestCheckVar(request("regdate"),10)

if not(C_ADMIN_USER) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('권한이 없습니다');"
	response.write "	self.close();"
	response.write "</script>"
	response.end	:	dbget.close()
end if

PriceEditEnable = false
if C_ADMIN_USER then PriceEditEnable = true		'//본사 직원일경우 수정권한

if C_ADMIN_USER then
'' 매장인경우
elseif (C_IS_SHOP) then
	shopid = C_STREETSHOPID
end if

itemgubunarr = split(itemgubunarr,"|")
itemidarr	= split(itemidarr,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
shopbuypricearr = split(shopbuypricearr,"|")
itemnoarr = split(itemnoarr,"|")
makeridarr = split(makeridarr,"|")
extbarcodearr = split(extbarcodearr,"|")
itemgubunarr2 = split(itemgubunarr2,"|")
itemidarr2	= split(itemidarr2,"|")
itemoptionarr2 = split(itemoptionarr2,"|")
itemnamearr2		= split(itemnamearr2,"|")
itemoptionnamearr2 = split(itemoptionnamearr2,"|")
sellcasharr2 = split(sellcasharr2,"|")
suplycasharr2 = split(suplycasharr2,"|")
shopbuypricearr2 = split(shopbuypricearr2,"|")
itemnoarr2 = split(itemnoarr2,"|")
makeridarr2 = split(makeridarr2,"|")
extbarcodearr2 = split(extbarcodearr2,"|")

cnt = uBound(itemidarr)
cnt2 = uBound(itemidarr2)

'//새로입력받은 내역과 기존내역 비교
for j=0 to cnt2-1
	isPreExists = false

	'//동일내역일경우 itemno 합산
	for i=0 to cnt-1
		if (itemgubunarr(i)=itemgubunarr2(j)) and (itemidarr(i)=itemidarr2(j)) and (itemoptionarr(i)=itemoptionarr2(j)) then
			itemnoarr(i) = CStr(CLng(itemnoarr(i)) + CLng(itemnoarr2(j)))
			isPreExists = true
			exit for
		end if
	next

	'//나머지 3 에 저장
	if Not isPreExists then
		itemgubunarr3 = itemgubunarr3 + itemgubunarr2(j) + "|"
		itemidarr3	= itemidarr3 + itemidarr2(j) + "|"
		itemoptionarr3 = itemoptionarr3 + itemoptionarr2(j) + "|"
		itemnamearr3		= itemnamearr3 + itemnamearr2(j) + "|"
		itemoptionnamearr3  = itemoptionnamearr3 + itemoptionnamearr2(j) + "|"
		sellcasharr3 = sellcasharr3 + sellcasharr2(j) + "|"
		suplycasharr3 = suplycasharr3 + suplycasharr2(j) + "|"
		shopbuypricearr3 = shopbuypricearr3 + shopbuypricearr2(j) + "|"
		itemnoarr3 = itemnoarr3 + itemnoarr2(j) + "|"
		makeridarr3 = makeridarr3 + makeridarr2(j) + "|"
		extbarcodearr3 = extbarcodearr3 + extbarcodearr2(j) + "|"
	end if
next

itemgubunarr2 = ""
itemidarr2	= ""
itemoptionarr2 = ""
itemnamearr2	= ""
itemoptionnamearr2 = ""
sellcasharr2 = ""
suplycasharr2 = ""
shopbuypricearr2 = ""
itemnoarr2 = ""
makeridarr2 = ""
extbarcodearr2 = ""

'//기존내역을 2에 할당
for i=0 to cnt-1
	itemgubunarr2 = itemgubunarr2 + itemgubunarr(i) + "|"
	itemidarr2	= itemidarr2 + itemidarr(i) + "|"
	itemoptionarr2 = itemoptionarr2 + itemoptionarr(i) + "|"
	itemnamearr2	= itemnamearr2 + itemnamearr(i) + "|"
	itemoptionnamearr2 = itemoptionnamearr2 + itemoptionnamearr(i) + "|"
	sellcasharr2 = sellcasharr2 + sellcasharr(i) + "|"
	suplycasharr2 = suplycasharr2 + suplycasharr(i) + "|"
	shopbuypricearr2 = shopbuypricearr2 + shopbuypricearr(i) + "|"
	itemnoarr2 = itemnoarr2 + itemnoarr(i) + "|"
	makeridarr2 = makeridarr2 + makeridarr(i) + "|"
	extbarcodearr2 = extbarcodearr2 + extbarcodearr(i) + "|"
next

'//기존내역과 신규내역 합치기
itemgubunarr = itemgubunarr2 + itemgubunarr3
itemidarr	= itemidarr2 + itemidarr3
itemoptionarr = itemoptionarr2 + itemoptionarr3
itemnamearr	= itemnamearr2 + itemnamearr3
itemoptionnamearr = itemoptionnamearr2 + itemoptionnamearr3
sellcasharr = sellcasharr2 + sellcasharr3
suplycasharr = suplycasharr2 + suplycasharr3
shopbuypricearr = shopbuypricearr2 + shopbuypricearr3
itemnoarr = itemnoarr2 + itemnoarr3
makeridarr = makeridarr2 + makeridarr3
extbarcodearr = extbarcodearr2 + extbarcodearr3

'//신규상품 추가시 팝업으로 넘어갈 경로		'/공용팝업으로 액션 페이지를 통채로 넘긴다
acURL =Server.HTMLEncode("/admin/offshop/maechul/manualmaechul_process.asp")

if regdate = "" then regdate = date()
%>

<script language="javascript">

	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//선택 색변환
	function CheckThis(frm){
		frm.cksel.checked=true;
		AnCheckClick(frm.cksel);
	}

	//바코드 상품등록
	function oneaddmanualItem(shopid){
		var tmpshopid = shopid;
		if (tmpshopid==''){
			alert('매장을 선택하세요.');
			frm.shopid.focus();
			return;
		}

		if (frm.barcode.value==''){
			alert('바코드를 입력 하세요.');
			frm.barcode.focus();
			return;
		}

		var tmpbarcode = frm.barcode.value;
		frm.barcode.value = '';

		var oneaddmanualItem = window.open("/admin/offshop/maechul/manualmaechul_process.asp?mode=oneaddmanualItem&barcode="+tmpbarcode+"&shopid="+tmpshopid+"&menupos=<%=menupos%>", "oneaddmanualItem", "width=50,height=50,scrollbars=yes,resizable=yes");
		oneaddmanualItem.focus();
	}

	function getOnload(){
	    frm.barcode.select();
	    frm.barcode.focus();
	}
	window.onload = getOnload;

	//일괄 상품등록
	function addmanualItem(shopid ,acURL){
		var tmpshopid = shopid;
		if (tmpshopid==''){
			alert('매장을 선택하세요.');
			frm.shopid.focus();
			return;
		}

		var addmanualItem = window.open("/common/offshop/pop_itemAddInfo2_off.asp?shopid="+tmpshopid+"&menupos=<%=menupos%>", "addmanualItem", "width=1024,height=768,scrollbars=yes,resizable=yes");
		addmanualItem.focus();
	}

	//상품등록
	function ReActItems(igubun,iitemid,iitemoption,isellcash,isuplycash,ishopbuyprice,iitemno,iitemname,iitemoptionname,imakerid,iextbarcode){
		var frmMaster = document.frm;

		frmMaster.itemgubunarr2.value = igubun;
		frmMaster.itemidarr2.value = iitemid;
		frmMaster.itemoptionarr2.value = iitemoption;
		frmMaster.sellcasharr2.value = isellcash;
		frmMaster.suplycasharr2.value = isuplycash;
		frmMaster.shopbuypricearr2.value = ishopbuyprice;
		frmMaster.itemnoarr2.value = iitemno;
		frmMaster.itemnamearr2.value = iitemname;
		frmMaster.itemoptionnamearr2.value = iitemoptionname;
		frmMaster.makeridarr2.value = imakerid;
		frmMaster.extbarcodearr2.value = iextbarcode;
		frmMaster.submit();
	}

	//수정 & 삭제
	function arredit(gubun){
		var msfrm = document.frm;
		var frm;

		if (!CheckSelected()){
			alert('선택아이템이 없습니다.');
			return;
		}

		msfrm.itemgubunarr.value = '';
		msfrm.itemidarr.value = '';
		msfrm.itemoptionarr.value = '';
		msfrm.itemnamearr.value = '';
		msfrm.itemoptionnamearr.value = '';
		msfrm.sellcasharr.value = '';
		msfrm.suplycasharr.value = '';
		msfrm.shopbuypricearr.value = '';
		msfrm.itemnoarr.value = '';
		msfrm.makeridarr.value = '';
		msfrm.extbarcodearr.value = '';

		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {

				//수정
				if (gubun=='E'){
					if (!IsInteger(frm.sellcash.value)){
						alert('판매가는 정수만 가능합니다.');
						frm.sellcash.focus();
						return;
					}
					if (!IsInteger(frm.suplycash.value)){
						alert('공급가는 정수만 가능합니다.');
						frm.suplycash.focus();
						return;
					}
					if (!IsInteger(frm.shopbuyprice.value)){
						alert('매장공급가는 정수만 가능합니다.');
						frm.shopbuyprice.focus();
						return;
					}
					if (!IsInteger(frm.itemno.value)){
						alert('갯수는 정수만 가능합니다.');
						frm.itemno.focus();
						return;
					}

					msfrm.itemgubunarr.value = msfrm.itemgubunarr.value + frm.itemgubun.value + "|";
					msfrm.itemidarr.value = msfrm.itemidarr.value + frm.itemid.value + "|";
					msfrm.itemoptionarr.value = msfrm.itemoptionarr.value + frm.itemoption.value + "|";
					msfrm.itemnamearr.value = msfrm.itemnamearr.value + frm.itemname.value + "|";
					msfrm.itemoptionnamearr.value = msfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
					msfrm.sellcasharr.value = msfrm.sellcasharr.value + frm.sellcash.value + "|";
					msfrm.suplycasharr.value = msfrm.suplycasharr.value + frm.suplycash.value + "|";
					msfrm.shopbuypricearr.value = msfrm.shopbuypricearr.value + frm.shopbuyprice.value + "|";
					msfrm.itemnoarr.value = msfrm.itemnoarr.value + frm.itemno.value + "|";
					msfrm.makeridarr.value = msfrm.makeridarr.value + frm.makerid.value + "|";
					msfrm.extbarcodearr.value = msfrm.extbarcodearr.value + frm.extbarcode.value + "|";

				//삭제
				}else if (gubun=='D'){
					if (!frm.cksel.checked){
						msfrm.itemgubunarr.value = msfrm.itemgubunarr.value + frm.itemgubun.value + "|";
						msfrm.itemidarr.value = msfrm.itemidarr.value + frm.itemid.value + "|";
						msfrm.itemoptionarr.value = msfrm.itemoptionarr.value + frm.itemoption.value + "|";
						msfrm.itemnamearr.value = msfrm.itemnamearr.value + frm.itemname.value + "|";
						msfrm.itemoptionnamearr.value = msfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
						msfrm.sellcasharr.value = msfrm.sellcasharr.value + frm.sellcash.value + "|";
						msfrm.suplycasharr.value = msfrm.suplycasharr.value + frm.suplycash.value + "|";
						msfrm.shopbuypricearr.value = msfrm.shopbuypricearr.value + frm.shopbuyprice.value + "|";
						msfrm.itemnoarr.value = msfrm.itemnoarr.value + frm.itemno.value + "|";
						msfrm.makeridarr.value = msfrm.makeridarr.value + frm.makerid.value + "|";
						msfrm.extbarcodearr.value = msfrm.extbarcodearr.value + frm.extbarcode.value + "|";
					}
				}
			}
		}

		msfrm.submit();
	}

	//매출전송
	function arrinsert(shopid){
		var msfrm = document.frm;
		var msupfrm = document.frmreg;
		var frm;

		var tmpshopid = shopid;
		if (tmpshopid==''){
			alert('매장을 선택하세요.');
			frm.shopid.focus();
			return;
		}
		if (msfrm.regdate.value==''){
			alert('매출날짜를 입력하세요.');
			msfrm.regdate.focus();
			return;
		}
		if (!CheckSelected()){
			alert('선택아이템이 없습니다.');
			return;
		}

		msupfrm.itemgubunarr.value = '';
		msupfrm.itemidarr.value = '';
		msupfrm.itemoptionarr.value = '';
		msupfrm.itemnamearr.value = '';
		msupfrm.itemoptionnamearr.value = '';
		msupfrm.sellcasharr.value = '';
		msupfrm.suplycasharr.value = '';
		msupfrm.shopbuypricearr.value = '';
		msupfrm.itemnoarr.value = '';
		msupfrm.makeridarr.value = '';
		msupfrm.extbarcodearr.value = '';

		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					if (!IsInteger(frm.sellcash.value)){
						alert('판매가는 정수만 가능합니다.');
						frm.sellcash.focus();
						return;
					}
					if (!IsInteger(frm.suplycash.value)){
						alert('공급가는 정수만 가능합니다.');
						frm.suplycash.focus();
						return;
					}
					if (!IsInteger(frm.shopbuyprice.value)){
						alert('매장공급가는 정수만 가능합니다.');
						frm.shopbuyprice.focus();
						return;
					}
					if (!IsInteger(frm.itemno.value)){
						alert('갯수는 정수만 가능합니다.');
						frm.itemno.focus();
						return;
					}

					msupfrm.itemgubunarr.value = msupfrm.itemgubunarr.value + frm.itemgubun.value + "|";
					msupfrm.itemidarr.value = msupfrm.itemidarr.value + frm.itemid.value + "|";
					msupfrm.itemoptionarr.value = msupfrm.itemoptionarr.value + frm.itemoption.value + "|";
					msupfrm.itemnamearr.value = msupfrm.itemnamearr.value + frm.itemname.value + "|";
					msupfrm.itemoptionnamearr.value = msupfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
					msupfrm.sellcasharr.value = msupfrm.sellcasharr.value + frm.sellcash.value + "|";
					msupfrm.suplycasharr.value = msupfrm.suplycasharr.value + frm.suplycash.value + "|";
					msupfrm.shopbuypricearr.value = msupfrm.shopbuypricearr.value + frm.shopbuyprice.value + "|";
					msupfrm.itemnoarr.value = msupfrm.itemnoarr.value + frm.itemno.value + "|";
					msupfrm.makeridarr.value = msupfrm.makeridarr.value + frm.makerid.value + "|";
					msupfrm.extbarcodearr.value = msupfrm.extbarcodearr.value + frm.extbarcode.value + "|";
				}
			}
		}

		var ret = confirm('매출내역을 전송하시겠습니까?');
		if (ret){

			msupfrm.shopid.value = tmpshopid;
			msupfrm.shopregdate.value = msfrm.regdate.value;
			msupfrm.mode.value = 'addmanualItem';
			msupfrm.action='/admin/offshop/maechul/manualmaechul_process.asp';
			msupfrm.target='view';
			msupfrm.submit();
		}
	}

</script>

<!-- 표 상단바 시작-->
※ 매장과 계약된 브랜드 상품만 등록 됩니다.
<table width="100%" align="center" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="post" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
<input type="hidden" name="itemidarr" value="<%= itemidarr %>">
<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
<input type="hidden" name="itemnamearr" value="<%= itemnamearr %>">
<input type="hidden" name="itemoptionnamearr" value="<%= itemoptionnamearr %>">
<input type="hidden" name="sellcasharr" value="<%= sellcasharr %>">
<input type="hidden" name="suplycasharr" value="<%= suplycasharr %>">
<input type="hidden" name="shopbuypricearr" value="<%= shopbuypricearr %>">
<input type="hidden" name="itemnoarr" value="<%= itemnoarr %>">
<input type="hidden" name="makeridarr" value="<%= makeridarr %>">
<input type="hidden" name="extbarcodearr" value="<%= extbarcodearr %>">
<input type="hidden" name="itemgubunarr2" value="">
<input type="hidden" name="itemidarr2" value="">
<input type="hidden" name="itemoptionarr2" value="">
<input type="hidden" name="itemnamearr2" value="">
<input type="hidden" name="itemoptionnamearr2" value="">
<input type="hidden" name="sellcasharr2" value="">
<input type="hidden" name="suplycasharr2" value="">
<input type="hidden" name="shopbuypricearr2" value="">
<input type="hidden" name="itemnoarr2" value="">
<input type="hidden" name="makeridarr2" value="">
<input type="hidden" name="extbarcodearr2" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				매장 : <%= getoffshopname(shopid) %><input type="hidden" name="shopid" value="<%= shopid %>">
				&nbsp;&nbsp;
				매출날짜 :
				<input type="text" name="regdate" size=10 maxlength=10 value="<%=regdate%>" onClick="jsPopCal('regdate');" style="cursor:hand;" readonly>
			</td>
		</tr>
		<tr>
			<td align="right">
				바코드(물류&공용)로 등록 :
				<input type="text" name="barcode" size=20 maxlength=20 onKeyPress="if(window.event.keyCode==13) oneaddmanualItem('<%=shopid%>');">
			</td>
		</tr>
		</table>
    </td>
</tr>
</form>
<form name="frmreg" method="post" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="shopregdate" value="">
	<input type="hidden" name="shopid" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="itemnamearr" value="">
	<input type="hidden" name="itemoptionnamearr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="shopbuypricearr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="makeridarr" value="">
	<input type="hidden" name="extbarcodearr" value="">
</form>
</table>
<!-- 표 상단바 끝-->
<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	<input type="button" class="button" value="선택수정" onclick="arredit('E')">
    	<input type="button" class="button" value="선택삭제" onclick="arredit('D')">
    </td>
    <td align="right">
    	<input type="button" value="상품검색" onclick="addmanualItem('<%=shopid%>','<%=acURL%>');" class="button">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<%
itemgubunarr = split(itemgubunarr,"|")
itemidarr	= split(itemidarr,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
shopbuypricearr = split(shopbuypricearr,"|")
itemnoarr = split(itemnoarr,"|")
makeridarr = split(makeridarr,"|")
extbarcodearr = split(extbarcodearr,"|")

cnt = ubound(itemidarr)
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= replace(cnt,"-1","0") %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onClick="ckAll(this)"></td>
	<td>물류코드</td>
	<td>공용<br>바코드</td>
	<td>브랜드</td>
	<td>상품명(옵션명)</td>
	<td>판매가</td>
	<td>수량</td>

	<% if (C_ADMIN_USER) then %>
		<td>업체<br>매입가</td>
		<td>매장<br>출고가</td>
	<% elseif (C_IS_SHOP) then %>
		<td>매장<br>출고가</td>
	<% elseif (C_IS_Maker_Upche) then %>
		<td>업체<br>매입가</td>
	<% else %>
		<td>업체<br>매입가</td>
		<td>매장<br>출고가</td>
	<% end if %>
</tr>
<% for i=0 to cnt-1 %>
<form name="frmBuyPrc_<%= i %>" method="post" action="">
<input type="hidden" name="itemgubun" value="<%= itemgubunarr(i) %>">
<input type="hidden" name="itemid" value="<%= itemidarr(i) %>">
<input type="hidden" name="itemoption" value="<%= itemoptionarr(i) %>">
<input type="hidden" name="makerid" value="<%= makeridarr(i) %>">
<input type="hidden" name="extbarcode" value="<%= extbarcodearr(i) %>">
<tr align="center" bgcolor="#FFFFFF">
	<td width="20">
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	</td>
	<td width="80">
		<%= itemgubunarr(i) %><%= CHKIIF(itemidarr(i)>=1000000,format00(8,itemidarr(i)),format00(6,itemidarr(i))) %><%= itemoptionarr(i) %>
	</td>
	<td width="90"><%= extbarcodearr(i) %></td>
	<td><%= makeridarr(i) %></td>
	<td align="left">
		<%= itemnamearr(i) %>
		<input type="hidden" name="itemname" value="<%= itemnamearr(i) %>">
		<input type="hidden" name="itemoptionname" value="<%= itemoptionnamearr(i) %>">

		<% if itemoptionnamearr(i) <> "" then %>
			(<%= itemoptionnamearr(i) %>)
		<% end if %>
	</td>

	<% if Not (PriceEditEnable) then %>
		<td align="right">
			<input type="hidden" name="sellcash" value="<%= sellcasharr(i) %>">
			<%= FormatNumber(sellcasharr(i),0) %>
		</td>

	<% else %>
		<td width="80">
			<input type="text" class="text" name="sellcash" value="<%= sellcasharr(i) %>" onKeyup="CheckThis(frmBuyPrc_<%= i %>)" size="8" maxlength="8">
		</td>
	<% end if %>

	<td width="60">
		<input type="text" class="text" name="itemno" value="<%= itemnoarr(i) %>" onKeyup="CheckThis(frmBuyPrc_<%= i %>)" size="4" maxlength="4">
	</td>

	<% if (C_ADMIN_USER) then %>
		<td width="80" align="right">
			<% if suplycasharr(i)<>"" and not isnull(suplycasharr(i)) then %>
				<%= FormatNumber(suplycasharr(i),0) %>
			<% else %>
				0
			<% end if %>

			<input type="hidden" name="suplycash" value="<%= suplycasharr(i) %>">
		</td>
		<td width="80" align="right">
			<% if shopbuypricearr(i)<>"" and not isnull(shopbuypricearr(i)) then %>
				<%= FormatNumber(shopbuypricearr(i),0) %>
			<% else %>
				0
			<% end if %>

			<input type="hidden" name="shopbuyprice" value="<%= shopbuypricearr(i) %>">
		</td>
	<% elseif (C_IS_SHOP) then %>
		<td width="80" align="right">
			<% if suplycasharr(i)<>"" and not isnull(suplycasharr(i)) then %>
				<%= FormatNumber(shopbuypricearr(i),0) %>
			<% else %>
				0
			<% end if %>

			<input type="hidden" name="suplycash" value="<%= suplycasharr(i) %>">
			<input type="hidden" name="shopbuyprice" value="<%= shopbuypricearr(i) %>">
		</td>
	<% elseif (C_IS_Maker_Upche) then %>
		<td width="80" align="right">
			<% if suplycasharr(i)<>"" and not isnull(suplycasharr(i)) then %>
				<%= FormatNumber(suplycasharr(i),0) %>
			<% else %>
				0
			<% end if %>

			<input type="hidden" name="suplycash" value="<%= suplycasharr(i) %>">
			<input type="hidden" name="shopbuyprice" value="<%= shopbuypricearr(i) %>">
		</td>
	<% else %>
		<td width="80" align="right">
			<% if suplycasharr(i)<>"" and not isnull(suplycasharr(i)) then %>
				<%= FormatNumber(suplycasharr(i),0) %>
			<% else %>
				0
			<% end if %>

			<input type="hidden" name="suplycash" value="<%= suplycasharr(i) %>">
		</td>
		<td width="80" align="right">
			<% if shopbuypricearr(i)<>"" and not isnull(shopbuypricearr(i)) then %>
				<%= FormatNumber(shopbuypricearr(i),0) %>
			<% else %>
				0
			<% end if %>

			<input type="hidden" name="shopbuyprice" value="<%= shopbuypricearr(i) %>">
		</td>
	<% end if %>
</tr>
</form>
<% next %>
<% if (cnt>0) then %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">
		<input type="button" class="button" value="매출전송" onclick="arrinsert('<%=shopid%>');">
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">
		판매 상품이 없습니다.
	</td>
</tr>
<% end if %>
</table>
<iframe id="view" name="view" width=0 hegiht=0 frameborder="0" scrolling="no"></iframe>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->