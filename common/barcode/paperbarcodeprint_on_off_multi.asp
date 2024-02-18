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
' Description : 바코드 출력 온라인 / 오프라인 통합(이문재 이사님 지시. 페이지 내부에서 분기함)
' Hieditor : 2016.12.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
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

dim i,page,research, currencyChar, currencyunit, olocation ,oproduct, isforeignprint, makerid, itemid, prdcode, generalbarcode
dim iA ,arrTemp,arrItemid, listgubun, printpriceyn, makeriddispyn, iPageSize, isdispsql, isdispconfirm, itembarcodearr
dim tmptdcnt, tmptrcnt, menupos, papername, wd, ht, qt, msg, imgPath, itemcopydispyn
dim ocstoragemaster, itemoptionyn
	tmptdcnt=0
	tmptrcnt=0
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<style type="text/css">

@font-face {
  font-family: 10X10;
  font-style: normal;
  font-weight: 400;
  src: url(http://fiximage.10x10.co.kr/fonts/10X10.eot);
  src: local('10X10 Regular'),
       local('10X10R'),
       url(http://fiximage.10x10.co.kr/fonts/10X10.eot?#iefix) format('embedded-opentype'),
       url(http://fiximage.10x10.co.kr/fonts/10X10.woff) format('woff'),
       url(http://fiximage.10x10.co.kr/fonts/10X10.ttf) format('truetype');
}

.currencychardefault {font-family:'gulim';}
.currencychar10X10 {font-family:'10X10';}

</style>
</head>
<body topmargin=0 leftmargin=0>

<%
'-------------------- 용지 규격 ------------------------- '/2017.01.04 한용민
'/ 치수 절대 건들지 말것. 쇼카드 용지 규격에 픽셀단위로 맞추어 놓음
'/ 칸 수 가로3칸 X 세로7칸 = 총 21칸
'/ 용지전체 : 662px X 1021px
'/ 칸 크기 : 208px X 133px
'/ 칸과 칸 사이에 간격 크기 : 가로 2칸 간격(각 19px X 133px) , 세로 6칸 간격(각 208px X 15px)
'-------------------- 용지 규격 -------------------------
%>

<% if onoffgubun="ONLINE" then %>
	<!-- #include virtual="/common/barcode/inc_paperbarcodeprint_on.asp"-->
<% elseif onoffgubun="OFFLINE" then %>
	<!-- #include virtual="/common/barcode/inc_paperbarcodeprint_off.asp"-->
<% else %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center"><font color="red"><strong>온라인 / 오프라인 구분을 선택 하세요.</strong></font></td>
	</tr>
	</table>
<% end if %>

</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
session.codePage = 949
%>