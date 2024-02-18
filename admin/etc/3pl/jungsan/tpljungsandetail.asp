<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/jungsanCls.asp"-->
<%

dim idx, gubun, rectorder, masteridx
dim yyyymm, i, j
masteridx      = requestCheckvar(request("idx"),10)
gubun   = requestCheckvar(request("gubun"),16)

if gubun="" then gubun="st"

dim sqlStr

dim otpljungsan, otpljungsanmaster, otpljungsanrealdetail, otpljungsangubundetail
set otpljungsanmaster = new CTplJungsan
otpljungsanmaster.FRectIdx = masteridx
otpljungsanmaster.GetTPLJungsanMasterList

if (otpljungsanmaster.FResultCount<1) then
    dbget_TPL.Close : dbget.Close(): response.end
end if


set otpljungsan = new CTplJungsan
otpljungsan.FRectMasterIdx = masteridx
otpljungsan.FRectGubun = gubun
otpljungsan.GetTplJungsanDetailList

yyyymm = otpljungsanmaster.FItemList(0).FYYYYmm


set otpljungsanrealdetail = new CTplJungsan
otpljungsanrealdetail.FRectMasterIdx = masteridx
otpljungsanrealdetail.FRectTplCompanyID = otpljungsanmaster.FItemList(0).Ftplcompanyid

set otpljungsangubundetail = new CTplJungsan

select case gubun
    case "cbm"
        otpljungsanrealdetail.FPageSize = 1000
        otpljungsanrealdetail.GetTplJungsanCbmList
    case else
        otpljungsanrealdetail.FPageSize = 5000
        otpljungsanrealdetail.FRectGubun = gubun
        otpljungsanrealdetail.GetTplJungsanEtcList

        otpljungsangubundetail.FRectGubun = gubun
        otpljungsangubundetail.GetTplJungsanGubunDetailList
end select


dim duplicated

%>
<script>
function addEtcList(iid,igubun){
	window.open('popetclistadd.asp?idx=' + iid + '&gubun=' + igubun,'popetc','width=700, height=150, location=no,menubar=no,resizable=yes,scrollbars=no,status=no,toolbar=no');
}

function DelDetail(frm){
	var ret = confirm('선택 내역을 삭제 하시겠습니까?');
	if (ret){
		frm.mode.value="deldetail";
		frm.submit();
	}
}

function ModiDetail(frm){
	var ret = confirm('선택 내역을 수정 하시겠습니까?');
	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="idx" value="<%= masteridx %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	업체ID : <b><%= otpljungsanmaster.FItemList(0).Ftplcompanyid %></b>
        	&nbsp;
			<input type="radio" name="gubun" value="cbm" <% if gubun="cbm" then response.write "checked" %> > 임대비용
			<input type="radio" name="gubun" value="ipchul" <% if gubun="ipchul" then response.write "checked" %> > 입출고비용
			<input type="radio" name="gubun" value="etc" <% if gubun="etc" then response.write "checked" %> > 기타비용
        </td>
        <td align="right">
        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="120">구분</td>
        <td width="120">구분상세</td>
        <td width="200">유형</td>
        <td width="80">단가</td>
        <td width="80">수량</td>
        <td width="80">금액</td>
        <td width="80">전월CBM</td>
        <td width="80">금월CBM</td>
        <td width="80">평균CBM</td>
		<td>코멘트</td>
    </tr>
    <% for i=0 to otpljungsan.FResultCount-1 %>
    <tr bgcolor="#FFFFFF" align="center" height="25">
      <td ><%= otpljungsan.FItemList(i).Fgubunname %></td>
      <td ><%= otpljungsan.FItemList(i).Fgubundetailname %></td>
      <td ><%= otpljungsan.FItemList(i).Ftypename %></td>
      <td align="right"><%= FormatNumber(otpljungsan.FItemList(i).Funitprice,0) %></td>
      <% if (otpljungsan.FItemList(i).Fgubunname = "임대비") and (otpljungsan.FItemList(i).Fgubundetailname = "상품보관") then %>
      <td align="right"><%= otpljungsan.FItemList(i).Favgcbm %></td>
      <% else %>
      <td align="right"><%= FormatNumber(otpljungsan.FItemList(i).Fitemno,0) %></td>
      <% end if %>
      <td align="right"><%= FormatNumber(otpljungsan.FItemList(i).FtotPrice,0) %></td>
      <% if (otpljungsan.FItemList(i).Fgubunname = "임대비") and (otpljungsan.FItemList(i).Fgubundetailname = "상품보관") then %>
      <td ><%= FormatNumber(otpljungsan.FItemList(i).Fprevcbm, 2) %></td>
      <td ><%= FormatNumber(otpljungsan.FItemList(i).Fcurrcbm, 2) %></td>
      <td ><%= FormatNumber(otpljungsan.FItemList(i).Favgcbm, 2) %></td>
      <% else %>
      <td></td>
      <td></td>
      <td></td>
      <% end if %>
      <td align="left"><%= otpljungsan.FItemList(i).Fcomment %></td>
    </tr>
    <% next %>
</table>

<p />

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<img src="/images/icon_arrow_down.gif" align="absbottom">
			<% if gubun="cbm" then %>
            <font color="red"><strong>CBM</strong>(최대 1,000건표시)</font> <%= otpljungsanrealdetail.FResultCount %> 건
			<% elseif gubun="ipchul" then %>
			<font color="red"><strong>입출고내역</strong></font>(최대 5,000건표시) <%= otpljungsanrealdetail.FResultCount %> 건
			<% elseif gubun="witakoffshop" then %>
			<font color="red"><strong>위탁 오프라인 판매내역</strong></font>(정산에 포함됨)
			<% end if %>
        </td>
        <td align="right">
        	<input type="button" class="button" value="기타내역추가" onclick="addEtcList(<%= otpljungsanmaster.FItemList(0).Fidx %>,'<%= gubun %>')">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<% if (gubun = "cbm") then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<td width="50">구분</td>
<td width="80">상품코드</td>
<td width="50">옵션</td>
<td width="100">바코드</td>
<td >상품명</td>
<td >옵션명</td>
<td width="80">수량</td>
<td width="80">CBM X(mm)</td>
<td width="80">CBM Y(mm)</td>
<td width="80">CBM Z(mm)</td>
<td width="30">삭제</td>
<td width="30">수정</td>
</tr>
<% for i=0 to otpljungsanrealdetail.FResultCount-1 %>
<form name="frmBuyPrcSell_<%= i %>" method="post" action="dotpljungsan.asp">
<input type="hidden" name="idx" value="<%= otpljungsanrealdetail.FItemList(i).Fidx %>">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="mode" value="">
<tr bgcolor="#FFFFFF">
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fitemgubun %></td>
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fitemid %></td>
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fitemoption %></td>
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fbarcode %></td>
<td ><%= otpljungsanrealdetail.FItemList(i).Fitemname %></td>
<td ><%= otpljungsanrealdetail.FItemList(i).Fitemoptionname %></td>
<td align="center">
    <input type="text" size="3" name="itemno" value="<%= otpljungsanrealdetail.FItemList(i).Fitemno %>" style="text-align:right">
</td>
<td align="center">
    <input type="text" size="3" name="cbmX" value="<%= otpljungsanrealdetail.FItemList(i).FcbmX %>" style="text-align:right">
</td>
<td align="center">
    <input type="text" size="3" name="cbmY" value="<%= otpljungsanrealdetail.FItemList(i).FcbmY %>" style="text-align:right">
</td>
<td align="center">
    <input type="text" size="3" name="cbmZ" value="<%= otpljungsanrealdetail.FItemList(i).FcbmZ %>" style="text-align:right">
</td>
<td ><a href="javascript:DelDetail(frmBuyPrcSell_<%= i %>)">삭제</a></td>
<td ><a href="javascript:ModiDetail(frmBuyPrcSell_<%= i %>)">수정</a></td>
</tr>
</form>
<%
'' 버퍼구성제한 초과시 아래 주석제거
if (i mod 1000)=0 then
    response.flush
end if
%>
<% next %>
</table>
<% else %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<td width="120">구분</td>
<td width="120">구분상세</td>
<td width="200">유형</td>
<td width="80">단가</td>
<td width="80">수량</td>
<td width="80">금액</td>
<td width="120">관련코드</td>
<td>코멘트</td>
<td width="30">삭제</td>
<td width="30">수정</td>
</tr>
<% for i=0 to otpljungsanrealdetail.FResultCount-1 %>
<form name="frmBuyPrcSell_<%= i %>" method="post" action="dotpljungsan.asp">
<input type="hidden" name="idx" value="<%= otpljungsanrealdetail.FItemList(i).Fidx %>">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="mode" value="">
<tr bgcolor="#FFFFFF">
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fgubunname %></td>
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fgubundetailname %></td>
<td align="center">
    <select class="select" name="itypename">
    <% for j=0 to otpljungsangubundetail.FResultCount-1 %>
    <% if (otpljungsanrealdetail.FItemList(i).Fgubunname = otpljungsangubundetail.FItemList(j).Fgubunname) and (otpljungsanrealdetail.FItemList(i).Fgubundetailname = otpljungsangubundetail.FItemList(j).Fgubundetailname) then %>
        <option value="<%= otpljungsangubundetail.FItemList(j).Ftypename %>" <%= CHKIIF(otpljungsanrealdetail.FItemList(i).Ftypename = otpljungsangubundetail.FItemList(j).Ftypename, "selected", "") %>><%= otpljungsangubundetail.FItemList(j).Ftypename %></option>
    <% end if %>
    <% next %>
    </select>
</td>
<td align="right">
    <input type="text" size="3" name="unitprice" value="<%= otpljungsanrealdetail.FItemList(i).Funitprice %>" style="text-align:right">
</td>
<td align="right">
    <input type="text" size="3" name="itemno" value="<%= otpljungsanrealdetail.FItemList(i).Fitemno %>" style="text-align:right">
</td>
<td align="right"><%= FormatNumber(otpljungsanrealdetail.FItemList(i).FtotPrice, 0) %></td>
<td align="center">
    <%= otpljungsanrealdetail.FItemList(i).Fmastercode %>
</td>
<td align="left">
    <%= otpljungsanrealdetail.FItemList(i).Fcomment %>
</td>
<td ><a href="javascript:DelDetail(frmBuyPrcSell_<%= i %>)">삭제</a></td>
<td ><a href="javascript:ModiDetail(frmBuyPrcSell_<%= i %>)">수정</a></td>
</tr>
</form>
<%
'' 버퍼구성제한 초과시 아래 주석제거
if (i mod 1000)=0 then
    response.flush
end if
%>
<% next %>
</table>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->
