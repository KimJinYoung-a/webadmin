<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
dim makerid, yyyy1,mm1, jgubun, targetGbn, groupid, page, finishflag, taxtype
makerid = requestCheckvar(request("makerid"),32)
yyyy1   = requestCheckvar(request("yyyy1"),10)
mm1     = requestCheckvar(request("mm1"),10)
jgubun  = requestCheckvar(request("jgubun"),10)
targetGbn= requestCheckvar(request("targetGbn"),10)
groupid  = requestCheckvar(request("groupid"),10)
finishflag = requestCheckvar(request("finishflag"),10)
taxtype   = requestCheckvar(request("taxtype"),10)

if (page="") then page=1

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

if (jgubun="") then
    jgubun = "MM"
end if

dim ojungsanTax
set ojungsanTax = new CUpcheJungsanTax
ojungsanTax.FPageSize = 5000
ojungsanTax.FCurrPage = 1
ojungsanTax.FRectMakerid = makerid
ojungsanTax.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTax.FRectJGubun = jgubun
ojungsanTax.FRectTargetGbn = targetGbn
ojungsanTax.FRectGroupid = groupid
ojungsanTax.FRectFinishFlag = finishflag
ojungsanTax.FRectTaxType = taxtype
ojungsanTax.getMonthUpcheJungsanListAdmAll


dim i
%>

<!-- 엑셀파일로 저장 헤더 부분 -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
Response.ContentType = "application/unknown"
Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="&yyyy1&"-"&mm1&CHKIIF(jgubun="CC","수수료","매입")&"정산내역.xls"
%>
<style type="text/css">
/* 엑셀 다운로드로 저장시 숫자로 표시될 경우 방지 */
.txt {mso-number-format:'\@'}
</style>
</head>
<body>


<% if (jgubun="CC") then %>
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="14" align="left"><strong>* 수수료 정산 내역</strong> <font color=red>(수수료 세금 계산서는 <b>텐바이텐</b>에서 <b>일괄 발행</b>합니다.)</font></td>
    <td colspan="2" align="right">총 <%=ojungsanTax.FTotalcount%> 건 <%=page%> / <%=ojungsanTax.FTotalpage%></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" class="txt">정산월</td>
    <td width="60" class="txt">매출처</td>
    <td width="50" class="txt">과세<br>구분</td>
    <td width="50" class="txt">그룹코드</td>
    <td width="50" class="txt">ERPCode</td>
    <td width="90" >회사명</td>
    <td width="90" >사업자번호</td>
    <td width="90" class="txt">브랜드ID</td>
    <td width="180"class="txt">정산내역</td>
    <td width="90" >고객판매금액<br>(협력사 매출액)</td>
    <td width="80" >수수료</td>
    <td width="80" >배송비/기타<br>(판매금액)</td>
    <td width="100">지급대상액<br>(상품)</td>
  	<td width="80">지급대상액<br>(배송비/기타)</td>
  	<td width="80">지급예정액</td>
    <td width="90" >계산서상태</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td>합계</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="left"></td>
    <td align="left"></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FPrdCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FdlvMeachulsum + ojungsanTax.FSumaryOneItem.FetMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FdlvJungsanSum + ojungsanTax.FSumaryOneItem.FetJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getToTalJungsanSum,0)%></td>
    <td></td>
</tr>
<% for i=0 to ojungsanTax.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td class="txt"><%=ojungsanTax.FItemList(i).Fyyyymm%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).getTargetNm%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).getItemVatTypeName%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fgroupid%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).FerpCust_cd%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fcompany_name%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fcompany_no%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Ftitle%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPrdCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FdlvMeachulsum + ojungsanTax.FItemList(i).FetMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FdlvJungsanSum + ojungsanTax.FItemList(i).FetJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).getToTalJungsanSum,0)%></td>
    <td><%=ojungsanTax.FItemList(i).GetTaxEvalStateName%></td>

</tr>
<% next %>


<% else %>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="13" align="left"><strong>* 매입 정산 내역</strong> (협력사에서 텐바이텐으로 발행해 주셔야 합니다.) (롯데닷컴 판매 내역 및 가맹점 판매 내역은 매입정산으로 처리 됩니다.)</td>
    <td colspan="2" align="right">총 <%=ojungsanTax.FTotalcount%> 건 <%=page%> / <%=ojungsanTax.FTotalpage%></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" class="txt">정산월</td>
    <td width="60" class="txt">매출처</td>
    <td width="50" class="txt">과세<br>구분</td>
    <td width="50" class="txt">그룹코드</td>
    <td width="50" class="txt">ERPCode</td>
   <td width="90" >회사명</td>
   <td width="90" >사업자번호</td>
    <td width="90" class="txt">브랜드ID</td>
    <td width="170" >정산내역</td>
    <td width="90" >협력사 매출액<br>(상품 공급액)</td>
    <td width="80" >배송비/기타</td>
    <td width="100">지급대상액<br>(상품)</td>
  	<td width="80">지급대상액<br>(배송비/기타)</td>
  	<td width="80">지급예정액</td>
    <td width="90" >계산서상태</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
    <td>합계</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="left"></td>
    <td align="left"></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FdlvMeachulsum + ojungsanTax.FSumaryOneItem.FetMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FdlvJungsanSum + ojungsanTax.FSumaryOneItem.FetJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getToTalJungsanSum,0)%></td>
    <td></td>
</tr>
<% for i=0 to ojungsanTax.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td class="txt"><%=ojungsanTax.FItemList(i).Fyyyymm%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).getTargetNm%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).getTaxtypeName%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fgroupid%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).FerpCust_cd%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fcompany_name%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fcompany_no%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Ftitle%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FdlvMeachulsum + ojungsanTax.FItemList(i).FetMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FdlvJungsanSum + ojungsanTax.FItemList(i).FetJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).getToTalJungsanSum,0)%></td>
    <td><%=ojungsanTax.FItemList(i).GetTaxEvalStateName%></td>

</tr>
<% next %>

</table>
<% end if %>
</body>
</html>
<%
set ojungsanTax = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
