<%@ language=vbscript %>
<% option explicit %>
<%
'Server.ScriptTimeOut = 60
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim yyyy1, mm1, gubun, page
dim yyyy_t, mm_t
yyyy1 = requestCheckvar(request("yyyy1"),10)
mm1 = requestCheckvar(request("mm1"),10)
gubun = requestCheckvar(request("gubun"),16)
page = requestCheckvar(request("page"),10)



if (gubun="") then gubun="chk0"
if (page="") then page=1

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

yyyy_t  = request("yyyy1")
mm_t    = request("mm1")

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FPageSize = 3000
ojungsan.FCurrPage = page
ojungsan.FRectGubun = gubun
ojungsan.FRectYYYYMM = yyyy1 + "-" + mm1

' if (gubun="witakchulgo") or (gubun="witakchulgoJS") then
'     if (gubun="witakchulgoJS") then ojungsan.FRectNotIncDivcode999="on"
' 	ojungsan.SearchWitakMaeipChulgoJungsanList
' end if

dim i, precode, ischeckd, isdisabled
dim checkdate1, checkdate2

%>
<script language='javascript'>
function popConfirm(yyyymm){
    var popwin = window.open('checkDuplicatedJungsan.asp?yyyymm=' + yyyymm,'checkDuplicatedJungsan','width=800,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popConfirm2(yyyymm){
    var popwin = window.open('checkDuplicatedJungsan_etc.asp?yyyymm=' + yyyymm,'checkDuplicatedJungsan','width=800,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function SelectCkMonly(opt){
	var bool = opt.checked;

	for (var i=0;i<document.forms.length;i++){
		var frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.hideMw.value=="M") {
			    frm.cksel.checked = bool;
			    AnCheckClick(frm.cksel);
			}
		}
	}


}

function SaveArr(igubun){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	upfrm.mode.value= igubun;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

    upfrm.idx.value = "";
    upfrm.yyyy.value = frmDumi.yyyy1.value;
    upfrm.mm.value  = frmDumi.mm1.value;

	if (!pass) {
		ret = confirm('선택 내역이 없습니다. \r\n\r\n ' + upfrm.yyyy.value + '-' + upfrm.mm.value + ' 정산대상 내역으로 저장 하시겠습니까?');
		if (!ret){
			return;
		}else{

		}
	}else{
		ret = confirm('선택 내역을 ' + upfrm.yyyy.value + '-' + upfrm.mm.value + ' 정산대상 내역으로 저장 하시겠습니까?');
	}



	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + ",";
				}
			}
		}
		upfrm.mode.value=igubun;
		upfrm.submit();
	}
}

function dobatch(frm,mode){



	frm.mode.value=mode;
	var ret = confirm('일괄처리를 실행 하시겠습니까?');
	if(ret){
		frm.submit();
	}
}

function etcBrandCpnJungsan(comp){
	if (confirm("브랜드쿠폰 차감 정산 작성 하시겠습니까?")){
		comp.form.submit();
	}
}

function etcBrandCpnIdxJungsan(comp){
	if (confirm("브랜드쿠폰 차감 정산 작성 하시겠습니까?")){
		comp.form.submit();
	}
}

function jsAddextBeasongPay(comp){
	if (confirm("제휴몰 누락배송비 매출등록 하시겠습니까?")){
		comp.form.submit();
	}
}

function popJungsanCheck(idifftp){
	var popwin = window.open("","popJungsanCheck","width=1200,height=800,scrollbars=yes,resizable=yes,status=yes");
	popwin.location.href="/admin/jungsan/popJungsanCheck.asp?difftp="+idifftp;

	popwin.focus();

}
</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="4" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">

	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" height="40" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
        <td align="left">
        	정산대상년월:<% DrawYMBox yyyy1,mm1 %>
        </td>

        <td rowspan="2" align="right" width="50">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td align="left">
			<input type="radio" name="gubun" value="chk0" <% if gubun="chk0" then response.write "checked" %> > 검토 - 제휴
			&nbsp;
			<input type="radio" name="gubun" value="chk2" <% if gubun="chk2" then response.write "checked" %> > 검토 - OFF
			&nbsp;
			<input type="radio" name="gubun" value="chk1" <% if gubun="chk1" then response.write "checked" %> > 검토 - ON

			&nbsp;|&nbsp;

			<input type="radio" name="gubun" value="act0" <% if gubun="act0" then response.write "checked" %> > 정산 BATCH 처리


			&nbsp;|&nbsp;
			<input type="radio" name="gubun" value="ext0" <% if gubun="ext0" then response.write "checked" %> > 기타정산 - 추가배송비 / 기타출고 / 브랜드쿠폰 / 매입마진쉐어


			&nbsp;|&nbsp;
			<input type="radio" name="gubun" value="chk9" <% if gubun="chk9" then response.write "checked" %> > 정산검토
			&nbsp;
			<input type="radio" name="gubun" value="act9" <% if gubun="act9" then response.write "checked" %> > 정산오픈처리
        </td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name=barchForm method=post action="dobatch.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
	<input type="hidden" name="mm" value="<%= mm1 %>">
	</form>

	<% if gubun="chk0" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<!-- 검토 -->
		<strong>1. 제휴 정산내역 수신 확인</strong>
		<br>&nbsp; - <a href="/admin/maechul/extjungsandata/extJungsanDataStatistic.asp?menupos=1656" target="_menu1656">[경영]승인자료>>제휴몰정산통계</a> : 구글시트를 참고하여 정산대사 수신금액이 맞는지 검토한다.
		<br>&nbsp; - <a href="https://docs.google.com/spreadsheets/d/1MNrTeCz1RvLE-Neuoh7RCQR5NstxWmeg5BLJAKTqH2o/edit#gid=441301323" target="_441301323">[googlesheet]제휴몰정산비교</a>
		<br>&nbsp; - <a href="/admin/maechul/extjungsandata/extJungsanDataList.asp?menupos=1652&mimap=on" target="_menu1652">[경영]승인자료>>제휴몰정산내역</a> : 미매핑내역이 없어야한다.(상품의경우)
		<br>&nbsp; - [경영]승인자료>>제휴몰정산통계 의 [정산vs주문입력검토] 의 쿠폰오차가 없도록 수정한다.
		<br>
		<br>
		<strong>2. 제휴 매입가 확인</strong>
		<br>&nbsp; - <a href="/admin/etc/difforder/orderMarginErrList.asp?menupos=3956" target="_menu3956">[입점제휴]제휴몰관리>>제휴사 마진 체크</a>
		: 제휴 입력 매입가가 제대로 되어 있는지 확인한다. 제휴unit에서 업체와 협의해서 별도 마진 진행하는경우도 있다.
		<br>
		<br>
	<% if (FALSE) then %>
		<strong>3. </strong>
	<% end if%>

		</td>
	</tr>
	<% elseif gubun="chk1" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<strong>1. 매입가 확인 / 쿠폰적용가확인 / 과세면세 / 매입가소수점</strong>
		<br>&nbsp; - [입점제휴]제휴몰관리>>제휴사 마진 체크</a>
		: 텐바이텐 매입가 관련 나머지 탭들 확인
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp; - <a href="/admin/etc/difforder/buycashErrList.asp?menupos=3956&vTab=2" target="_menu3956_tab">매입가오류체크</a>
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : ezwel의 경우 판매가를 100원 단위로 보낸다. 역마진이 있다.<br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 10x10의 경우 쿠폰번호 있는경우 (자사부담쿠폰일경우)역마진이 있을 수 있다.<br>
		&nbsp;&nbsp;&nbsp;&nbsp; - <a href="/admin/etc/difforder/buycashOverList.asp?menupos=3956&vTab=3" target="_menu3956_tab">원매입가보다 상품쿠폰 적용 매입가가 큰경우</a>
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 상품옵션변경시 잘못 들어가는경우가 간혹 있다.<br>
		&nbsp;&nbsp;&nbsp;&nbsp; - <a href="/admin/etc/difforder/taxErrList.asp?menupos=3956&vTab=4" target="_menu3956_tab">면세 오등록 체크</a>
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp; - <a href="/admin/etc/difforder/buycashPrimeList.asp?menupos=3956&vTab=5" target="_menu3956_tab">상품/옵션공급가소수점</a>
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 업체/위탁상품의 경우 매입가에 소수점이 들어가면 곤란하다 (일별배치로 처리)<br>


		<br>
		<strong>2. 주문원장, 정산확정일 검토</strong>
		<% if (FALSE) then %><br>&nbsp; - 주문원장 오차 검토 --> 매출로그 검토에포함됨 <% end if %>
		<br>&nbsp; - <a href="#" onClick="popJungsanCheck('900');return false;">정산확정일 검토</a>

		<br>
		<br>
		<strong>3. 매출로그 검토</strong>
		<br>&nbsp; -  <a href="#" onClick="popJungsanCheck('');return false;">pop</a>
		<br>
	<% if (FALSE) then %>
		<strong>4.</strong>
		<br>

		</td>
	<% end if %>
	</tr>
	<% elseif gubun="chk2" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<strong>1. 매장 포스입력 데이터 확인</strong>
		<br>&nbsp; - <a href="/admin/offshop/offshopjumun_error.asp?menupos=1183" target="_menu1183">[OFF]오프_통계관리>>매입가 오류 판매내역</a>
		: 계약 미지정 판매내역 확인
		<br>
		<br>
		<strong>2. 매출로그 검토</strong>
		<br>&nbsp; - <a href="#" onClick="popJungsanCheck('200');return false;">pop</a>
	<% if (FALSE) then %>
		<strong>3. 매장 주문 원장확인</strong>
		<br>&nbsp; -
		<br>

		<br>
		<br>
		<strong>3. </strong>
		<br>&nbsp; -
		<br>
	<% end if %>
		<br>
		<br>
		</td>
	</tr>
	<% elseif gubun="act0" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<strong>1. OFF 정산 일괄처리</strong>
		<br>&nbsp; - <a href="/admin/upchejungsan/popjungsanmakebatch.asp?targetGbn=OF&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>" target="_popjungsanmakebatch_onff">OFF 정산작성</a>
		<br>
		<br>

		<strong>2. ON 정산 일괄처리</strong>
		<br>&nbsp; - <a href="/admin/upchejungsan/popjungsanmakebatch.asp?targetGbn=ON&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>" target="_popjungsanmakebatch_onff">ON 정산작성</a>
		<br>
		<br>
		<strong>3. Class 정산 일괄처리</strong>
		<br>&nbsp; - 2020/03 개인도 수수료 정산으로 변경됨
		<br>

		<br>
		<br>
		</td>
	</tr>
	<% elseif gubun="ext0" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<strong>1. 추가 배송비정산(업체추가정산)</strong>
		<br>&nbsp; - <a href="/admin/upchejungsan/upchedeliverypay.asp?menupos=975" target="_menu975"> [정산]정산내역작성>>[ON]배송비정산</a>
		<br>
		<br>

		<strong>2. 기타출고정산</strong>
		<br>&nbsp; - <a href="/admin/upchejungsan/etcchulgojungsan.asp?menupos=321&gubun=witakchulgoJS" target="_menu321"> [정산]정산내역작성>>[ON]기타출고정산</a>
		<br>
		<br>
		<strong>3. 브랜드쿠폰 차감정산 </strong>
		<form name="frmbrandcpn" method="post" action="dobatch.asp">
		<input type="hidden" name="mode" value="brandcpn">
		<br>
		&nbsp; - 정산월 <input type="text" name="jyyyymm" value="<%=yyyy1%>-<%=mm1%>" size="5" maxlength="7">
		&nbsp; - 차수 <input type="text" name="differencekey" value="0" size="1" maxlength="2">
		&nbsp; - 브랜드ID <input type="text" name="makerid" value="<%= "ETUDEHOUSE" %>" size="12" maxlength="32">
		&nbsp; - 업체부담율 <input type="text" name="upchepro" value="<%= "70" %>" size="2" maxlength="3" style="text-align:right"> %
		&nbsp; <input type="button" value="차감 정산 작성" onClick="etcBrandCpnJungsan(this)">
		</form>
		<form name="frmbrandcpnidx" method="post" action="dobatch.asp">
		<input type="hidden" name="mode" value="brandcpnidx">
		<br>
		&nbsp; - 정산월 <input type="text" name="jyyyymm" value="<%=yyyy1%>-<%=mm1%>" size="5" maxlength="7">
		&nbsp; - 차수 <input type="text" name="differencekey" value="0" size="1" maxlength="2">
		&nbsp; - 브랜드ID <input type="text" name="makerid" value="<%= "ETUDEHOUSE" %>" size="12" maxlength="32">
		&nbsp; - 쿠폰번호 <input type="text" name="cpnidx" value="<%= "7777" %>" size="4" maxlength="32">
		&nbsp; - 업체부담율 <input type="text" name="upchepro" value="<%= "70" %>" size="2" maxlength="3" style="text-align:right"> %
		&nbsp; <input type="button" value="차감 정산 작성" onClick="etcBrandCpnIdxJungsan(this)">
		</form>
		<br><br>
		<strong>4. 매입마진쉐어 차감정산</strong>
		<br>&nbsp; - <a href="/admin/shopmaster/sale/maeipSaleMarginList.asp?menupos=3967" target="_menu3967"> [ON]상품관리>>할인 마진쉐어(매입상품)</a>
		<br><br>
		<strong>5. 배송비분담 프로모션 비용정산</strong>
		<br>&nbsp; - <a href="/admin/sitemaster/halfDeliveryPay/index.asp?menupos=4155" target="_menu4155"> [ON]상품관리>>배송비부담설정</a>
		<br><br>
		<strong>6. 제휴 누락배송비 매출등록</strong>
		<form name="frmextbeasongPay" method="post" action="dobatch.asp">
		<input type="hidden" name="mode" value="addextbeasongPay">
		<br>
		&nbsp; - 제휴몰 <input type="text" name="sitename" value="" size="12" maxlength="32">
		&nbsp; - 출고일자 <input type="text" name="yyyymmdd" value="" size="10" maxlength="10">
		&nbsp; - 배송비 금액 <input type="text" name="beasongPay" value="" size="12" maxlength="32">
        &nbsp; <input type="button" value="배송비 매출 등록" onClick="jsAddextBeasongPay(this)">
        </form>
		<br>
		<br>

		<br>
		<br>
		</td>
	</tr>
	<% elseif gubun="chk9" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<strong>1. 정산/매출로그 비교 ON</strong>
		<br>&nbsp; - <a href="#" onClick="popJungsanCheck('300');return false;"> pop</a>
		<br>
		<br>

		<strong>2. 정산/매출로그 비교 OF</strong>
		<br>&nbsp; - <a href="#" onClick="popJungsanCheck('400');return false;"> pop</a>
		<br>

		<br>
		<br>
		</td>
	</tr>
	<% elseif gubun="act9" then %>
	<tr bgcolor="#FFFFFF">
		<td>
		 <input type="button" class="button" value="수정중->업체확인중 일괄처리 ON" onClick="javascript:dobatch(barchForm,'finishflag1');"><br><br>

		 <input type="button" class="button" value="수정중->업체확인중 일괄처리 OF" onClick="javascript:dobatch(barchForm,'finishflagoff1');"><br><br>


		</td>
	</tr>
	<% end if %>
</table>


<form name="frmArrupdate" method="post" action="dobatch.asp">
<input type="hidden" name="idx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
</form>
<%
SET ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
