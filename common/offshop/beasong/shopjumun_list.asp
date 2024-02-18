<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.22 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->
<%
dim showshopselect, loginidshopormaker ,tmodlvType ,ojumun , i , orderno, UserHpAuto, certsendgubun, odlvTypedefault, totrealprice
	orderno = requestcheckvar(request("orderno"),16)
	UserHpAuto = requestcheckvar(request("UserHpAuto"),16)
	certsendgubun = requestcheckvar(request("certsendgubun"),32)

totrealprice=0
showshopselect = false
loginidshopormaker = ""
if certsendgubun="" then certsendgubun = "KAKAOTALK"
odlvTypedefault="1"

if C_ADMIN_USER or C_IS_OWN_SHOP then
	showshopselect = true
	loginidshopormaker = request("shopid")
elseif (C_IS_SHOP) then
	'직영/가맹점
	loginidshopormaker = C_STREETSHOPID
else
	if (C_IS_Maker_Upche) then
		loginidshopormaker = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
			loginidshopormaker = "--"		'표시안한다. 에러.
		else
			showshopselect = true
			loginidshopormaker = request("shopid")
		end if
	end if
end if

set ojumun = new cupchebeasong_list
	ojumun.frectorderno = orderno
	ojumun.frectshopid = loginidshopormaker

if (orderno <> "") then
	ojumun.fshopjumun_list()
end if

'//체크박스 disabled 조건 처리
function checkblock(CurrState)
	checkblock = false

	'// 배송입력이 있는 경우 true
	if Not IsNull(CurrState) then
		checkblock = true
	end if
end Function
%>

<script language="javascript">

	//주문수정
	function jumundetail(masteridx){
		frm.masteridx.value=masteridx;
		frm.action='/common/offshop/beasong/shopbeasong_input.asp';
		frm.submit();
	}

	// 배송자 일괄 지정
	function chodlvTypedefault(){
		var odlvTypedefault = document.getElementById("odlvTypedefault");		//getElementsByName

		if (odlvTypedefault.value==''){
			alert('일괄지정 하실 기준 배송자가 지정되어 있지 않습니다.');
			odlvTypedefault.focus();
			return;
		}

		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				frm.cksel.checked = true;
				AnCheckClick(frm.cksel);
				frm.odlvType.value=odlvTypedefault.value;
			}
		}
	}
	
	//매장에서 상품 배송 요청
	function jumuninput(upfrm,isAuto){
		if (upfrm.orderno.value==''){
			alert('주문번호를 입력 하세요');
			upfrm.orderno.focus();
			return;
		}
		var orderno = upfrm.orderno.value;

		if (isAuto=='auto'){
			if (upfrm.certsendgubun.value==''){
				alert('핸드폰번호를 발송할 인증 구분이 없습니다.');
				upfrm.certsendgubun.focus();
				return;
			}
			var certsendgubun = upfrm.certsendgubun.value;

			if (upfrm.UserHpAuto.value==''){
				alert('인증 받으실 휴대폰번호를 입력 하세요');
				upfrm.UserHpAuto.focus();
				return;
			}
			var UserHpAuto = upfrm.UserHpAuto.value;
		}

		upfrm.itemgubunarr.value='';
		upfrm.itemidarr.value='';
		upfrm.itemoptionarr.value='';
		upfrm.masteridx.value='';
		upfrm.masteridxarr.value='';
		upfrm.odlvTypearr.value='';

		if (!CheckSelected()){
			alert('선택 상품이 없습니다.');
			return;
		}

		var frm; var odlvType='';
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					if (frm.odlvType.value==''){
						alert('배송구분을 선택 하세요');
						frm.odlvType.focus();
						return;
					}
					// comm_cd : B031 매입출고정산 / B012 업체특정 / B013 출고특정
					// 물류배송
					if (frm.odlvType.value*1 == '1') {
						if (frm.comm_cd.value == 'B012') {
							alert("해당 상품은 업체배송 or 매장배송만 가능 합니다.");
							frm.odlvType.focus();
							return;
						}
					}
					// 업체배송
					if (frm.odlvType.value*1 == '2') {
						if (frm.comm_cd.value == 'B031' || frm.comm_cd.value == 'B013') {
							alert("해당 상품은 물류배송 or 매장배송만 가능 합니다.");
							frm.odlvType.focus();
							return;
						}
					}
					
/*
					if (frm.defaultbeasongdiv.value*1 == 0) {
						if (frm.odlvType.value*1 != 0) {
							alert("지정할수 없는 배송자입니다. 매장배송을 선택하세요.");
							frm.odlvType.focus();
							return;
						}
					}

					if (odlvType!='' && odlvType != frm.odlvType.value){
						alert('배송구분은 상품별로 다르게 지정하실수 없습니다.');
						frm.odlvType.focus();
						return;
					}
*/

					upfrm.odlvTypearr.value = upfrm.odlvTypearr.value + frm.odlvType.value + "," ;
					upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "," ;
					upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "," ;
					upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "," ;
					upfrm.shopidarr.value = upfrm.shopidarr.value + frm.shopid.value + "," ;

					// 계산이 끝난후 현재 계약 조건을 저장한다.
					odlvType = frm.odlvType.value;
				}
			}
		}

		if (isAuto=='auto'){
			if (confirm('고객님( ' + UserHpAuto + ' )께 주소입력 링크를 '+ certsendgubun + ' 으로 발송 하시겠습니까?')){
				upfrm.mode.value='userjumun';
				upfrm.action='/common/offshop/beasong/shopbeasong_process.asp';
				upfrm.submit();
			}
		}else{
			if (confirm('주소를 수기로 매장에서 직접 입력 하시겠습니까?')){
				upfrm.mode.value='shopjumun';
				upfrm.action='/common/offshop/beasong/shopjumun_address.asp';
				upfrm.submit();
			}
		}
	}

	//폼로드시 셀렉트
	function getOnload(){
	    frm.orderno.select();
	    frm.orderno.focus();
	    
	    <% if (session("poslogin") = 1) then %>
	    // POS에서 넘어온 것이면
	        setTimeout(function(){
	            reqPosSign();
            },100);
	    <% end if %>
	}

	window.onload = getOnload;

	//폼전송
	function gosubmit(){
		frm.submit();
	}

	function CheckThis(frm){
		frm.cksel.checked=true;
		AnCheckClick(frm.cksel);
	}

	// 초기값 전체 선택
	function AllCheck(){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if ( frm.cksel.disabled != true){
					frm.cksel.checked = true;
					AnCheckClick(frm.cksel);
				}
			}
		}
	}

    /* pos 통신 */
    function reqPosSign(){
        window.location = "jsTenPosCall://tenPosReqSign?resultPosSign";
    }
    
    function resultPosSign(ival){
        document.frm.UserHpAuto.value=ival;
    }
</script>

<!-- 검색 시작 -->
<form name="frm" method="post" action="">
<input type="hidden" name="itemgubunarr">
<input type="hidden" name="itemidarr">
<input type="hidden" name="itemoptionarr">
<input type="hidden" name="shopidarr">
<input type="hidden" name="masteridxarr">
<input type="hidden" name="mode">
<input type="hidden" name="masteridx">
<input type="hidden" name="odlvTypearr">
<input type="hidden" name="detailidxarr">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 매장 :
		<% if (showshopselect = true) then %>
			<% 'drawSelectBoxOffShop "shopid",loginidshopormaker %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",loginidshopormaker, "21") %>
		<% else %>
			<%= loginidshopormaker %>
		<% end if %>
		&nbsp;&nbsp;
		* 주문번호 : <input type="text" name="orderno" value="<%= orderno %>" size="16" onKeyPress="if(window.event.keyCode==13) gosubmit('');">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		* 휴대폰번호 : <input type="text" name="UserHpAuto" value="<%= UserHpAuto %>" size=16 maxlength=16>
		<% drawcertsendgubun "certsendgubun", certsendgubun, "", "N" %>
		<input type="button" value="선택상품 주소입력 링크발송" class="button" onclick="jumuninput(frm,'auto')">
		&nbsp;&nbsp;&nbsp;
		<input type="button" value="POS 사인패드요청"  class="button" onclick="reqPosSign()">
		
	</td>
	<td align="right">
		<input type="button" value="선택상품 주소입력(직원수기입력)" class="button" onclick="jumuninput(frm,'')">
	</td>
</tr>
</table>
</form>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ojumun.FTotalCount %></b><br>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>판매매장</td>
	<td>주문번호</td>
	<td>상품코드</td>
	<td>브랜드ID</td>
	<td>상품명<br>옵션명</td>
	<td>판매금액</td>
	<td>실결제액</td>
	<td>판매수량</td>
	<td>합계</td>
	<td>판매일</td>
	<td>기본배송구분</td>
	<td>
		배송자지정
		<!--<Br><% 'Drawbeasonggubun "odlvTypedefault", odlvTypedefault," id='odlvTypedefault'" %>
		<input type="button" value="일괄지정" class="button" onclick="chodlvTypedefault();">-->
	</td>
	<td>배송상태</td>
	<td>비고</td>
</tr>
<% if ojumun.FTotalCount>0 then %>
<% for i=0 to ojumun.FTotalCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">
<input type="hidden" name="orderno" value="<%= ojumun.FItemList(i).forderno %>">
<input type="hidden" name="itemgubun" value="<%= ojumun.FItemList(i).fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ojumun.FItemList(i).fitemid %>">
<input type="hidden" name="itemoption" value="<%= ojumun.FItemList(i).fitemoption %>">
<input type="hidden" name="shopid" value="<%= ojumun.FItemList(i).fshopid %>">
<input type="hidden" name="masteridx" value="<%= ojumun.FItemList(i).fmasteridx %>">
<input type="hidden" name="detailidx" value="<%= ojumun.FItemList(i).fdetailidx %>">
<input type="hidden" name="comm_cd" value="<%= ojumun.FItemList(i).fcomm_cd %>">

<% if ojumun.FItemList(i).fcurrstate = "" or isnull(ojumun.FItemList(i).fcurrstate) then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>
<tr align="center" bgcolor="#FFFFaa">
<% end if %>
	<td>
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if checkblock(ojumun.FItemList(i).FCurrState) then response.write " disabled" %>>
	</td>
	<td>
		<%= ojumun.FItemList(i).fshopname %>
	</td>
	<td>
		<%= ojumun.FItemList(i).forderno %>
	</td>
	<td>
		<%=ojumun.FItemList(i).fitemgubun%>-<%=CHKIIF(ojumun.FItemList(i).fitemid>=1000000,Format00(8,ojumun.FItemList(i).fitemid),Format00(6,ojumun.FItemList(i).fitemid))%>-<%=ojumun.FItemList(i).fitemoption%>
	</td>
	<td>
		<%=ojumun.FItemList(i).fmakerid%>
	</td>
	<td>
		<%= ojumun.FItemList(i).fitemname %><br><%= ojumun.FItemList(i).fitemoptionname %>
	</td>
	<td><%= FormatNumber(ojumun.FItemList(i).fsellprice,0) %></td>
	<td><%= FormatNumber(ojumun.FItemList(i).frealsellprice,0) %></td>
	<td>
		<%= ojumun.FItemList(i).fitemno %>
	</td>
	<td><%= FormatNumber(ojumun.FItemList(i).frealsellprice*ojumun.FItemList(i).fitemno,0) %></td>
	<td>
		<%= ojumun.FItemList(i).fIXyyyymmdd %>
	</td>
	<input type="hidden" name="defaultbeasongdiv" value="<%= ojumun.FItemList(i).Fdefaultbeasongdiv %>">
	<td>
		<% if (ojumun.FItemList(i).Fdefaultbeasongdiv <> 0) then %>
			<%= ojumun.FItemList(i).getDefaultBeasongDivName %>
		<% end if %>
	</td>
	<%
	tmodlvType = ojumun.FItemList(i).fodlvType

	if (tmodlvType = "") or (IsNull(tmodlvType)) then
		' 물류배송(매입출고정산 , 출고특정) , 업체배송(업체특정) , 매장배송(ALL)
		if ojumun.FItemList(i).fcomm_cd="B031" or ojumun.FItemList(i).fcomm_cd="B013" then
			tmodlvType = "1"
		elseif ojumun.FItemList(i).fcomm_cd="B012" then
			tmodlvType = "2"
		else
			tmodlvType = "1"
		end if
	end if
	%>
	<td>
		<% Drawbeasonggubun "odlvType", tmodlvType," onchange='CheckThis(frmBuyPrc"& i &");'" %>
	</td>
	<td>
		<%= ojumun.FItemList(i).shopNormalUpcheDeliverState %>
		<%
		'//결제완료 상태가 아니라면  출고요청 번호 보여줌
		if ojumun.FItemList(i).FCurrState<>"" or not isnull(ojumun.FItemList(i).FCurrState) then
		%>
			<!--
			<br>(일렬번호 : <%= ojumun.FItemList(i).fmasteridx %>)
			-->
		<% end if %>
	</td>
	<td>
		<%
		'//주문대기 상태 일때 주소 수정 가능
		if ojumun.FItemList(i).FCurrState="0" then
		%>
			<input type="button" onclick="jumundetail(<%= ojumun.FItemList(i).fmasteridx %>);" value="주문수정" class="button">
		<%
		end if
		%>
	</td>
</tr>
</form>
<%
totrealprice = totrealprice + (ojumun.FItemList(i).frealsellprice*ojumun.FItemList(i).fitemno)
next
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=9>합계</td>
	<td><%= FormatNumber(totrealprice,0) %></td>
	<td colspan=9></td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<script type="text/javascript">
	<% if ojumun.FTotalCount>0 then %>
		AllCheck();
	<% end if %>
</script>

<br>

* 200 건 까지 검색 됩니다.<br>
* 배송정보를 재입력하시려면 먼저 입력된 배송정보를 삭제하세요.
<%
set ojumun = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->