<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 바코드 출력 온라인 / 오프라인 통합
' Hieditor : 2016.12.15 한용민 생성
'/////////////////// 이파일 수정시 밑에 파일도 모두 동일하게 같이 고쳐야 한다. ////////////////////////
' SCM : /common/barcode/barcodeprint_on_off_multi.asp
' 		/partner/common/barcode/barcodeprint_on_off_multi.asp
' 		/partner/common/barcode/barcodeprint_on_off_multi_pop.asp
' LOGICS : /v2/common/barcode/barcodeprint_on_off_multi.asp
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<!-- #include virtual="/lib/classes/stock/ipchulbarcodecls.asp"-->
<!-- #include virtual="/lib/classes/stock/ipchullocationcls_UTF8.asp"-->

<!-- #include virtual="/lib/classes/stock/ipchulproductcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->

<%
dim onoffgubun
	onoffgubun = requestCheckVar(request("onoffgubun"), 10)

if (C_IS_SHOP) then
	'/가맹점 일경우
	if getoffshopdiv(C_STREETSHOPID) = "3" then
		if onoffgubun="" then onoffgubun="OFFLINE"
	end if
end if

dim i,page,research, currencyChar, currencyunit, currencyunit_pos, printername, olocation ,oproduct, isforeignprint, makerid, itemid, prdcode, generalbarcode
dim iA ,arrTemp,arrItemid, listgubun, printpriceyn, makeriddispyn, iPageSize, isdispsql, isdispconfirm, papername, itemcopydispyn
dim ocstoragemaster, itemoptionyn, titledispyn
%>

<script type="text/javascript" src="/js/barcode.js"></script>
<script type="text/javascript" src="/js/ttpbarcode_utf8.js"></script>
<script type="text/javascript" src="/js/DOSHIBAbarcode_utf8.js"></script>
<script type="text/javascript">

function gosubmit(){
	frmgubun.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frmgubun" method="get" action="" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="150" bgcolor="<%= adminColor("gray") %>">온라인 / 오프라인 구분</td>
	<td align="left">
		<% drawonoffgubun "onoffgubun", onoffgubun, " onchange='gosubmit();'" %>
	</td>
	<td align="right">
		<a href="http://imgstatic.10x10.co.kr/offshop/sample/print/TOSHIBA_TEC_B-FV4_itembarcode_manual_v3.zip" target="_blank">
		<font color="red">TEC B-FV4 상품바코드 프린트 설치 설명서</font></a>
		/ <a href="http://imgstatic.10x10.co.kr/offshop/sample/print/TSC_TTP-243_itembarcode_manual_v3.zip" target="_blank">
		<font color="red">TSC TTP-243 상품바코드 프린트 설치 설명서</font></a>
		<Br>
		<a href="http://imgstatic.10x10.co.kr/offshop/sample/font/10X10_FONT_manual_v1.zip" target="_blank">
		<font color="red">10X10 폰트 설치 설명서</font></a>
		<Br>
		<a href="https://imgstatic.10x10.co.kr/offshop/sample/print/SHOWCARD_print_manual_v1.docx" target="_blank" onfocus="this.blur" >
		<font color="red">쇼카드출력설명서</font></a>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<% if onoffgubun="ONLINE" then %>
	<!-- #include virtual="/common/barcode/inc_barcodeprint_on.asp"-->
<% elseif onoffgubun="OFFLINE" then %>
	<!-- #include virtual="/common/barcode/inc_barcodeprint_off.asp"-->
<% else %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center"><font color="red"><strong>온라인 / 오프라인 구분을 선택 하세요.</strong></font></td>
	</tr>
	</table>
<% end if %>

<%
session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
