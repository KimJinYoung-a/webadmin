<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 퀵배송 출고완료리스트
' History : 2017년 12월 21일 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim xl
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

xl = request("xl")

dim nowdate,date1,date2,Edate
nowdate = now

if (yyyy1="") then
	''date1 = dateAdd("d",-4,nowdate)
	date1 = nowdate
	yyyy1 = Left(CStr(date1),4)
	mm1   = Mid(CStr(date1),6,2)
	dd1   = Mid(CStr(date1),9,2)

	yyyy2 = Left(CStr(nowdate),4)
	mm2   = Mid(CStr(nowdate),6,2)
	dd2   = Mid(CStr(nowdate),9,2)

	Edate = Left(CStr(nowdate+1),10)
else
	Edate = Left(CStr(dateserial(yyyy2, mm2 , dd2)+1),10)
end if

dim objbaljumakeonorder, balju_code

balju_code	= requestCheckVar(request("balju_code"),10)

set objbaljumakeonorder = New COldMiSend
objbaljumakeonorder.FPageSize = 500

if (balju_code <> "") then
	objbaljumakeonorder.FRectBaljuCode = balju_code
else
	objbaljumakeonorder.FRectStartDate = yyyy1 + "-" + mm1 + "-" + dd1
	objbaljumakeonorder.FRectEndDate = Left(CStr(Edate),10)
end if

objbaljumakeonorder.GetBaljuListQuickOrder

dim i, tmp
dim orgitemno, makeonorderitemno


if (xl <> "") then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename=quickorderlist"& replace(date&hour(now)&minute(now),"-","") &".xls"

	%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<style type="text/css">
 body {font-family:tahoma;font-size:12px}
 table {padding:2px;border-spacing:0px;font-family:tahoma;font-size:12px;border-collapse:collapse}
 td {text-align:center}
 .titbg {background-color:#FEE;}
</style>
</head>
<body>
<table>
	<tr>
		<td>업체명</td>
		<td></td>
		<td>버전</td>
		<td>1</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<tr>
		<td>고객주문번호</td>
	    <td>보내는분</td>
	    <td>연락처1</td>
	    <td>주소지</td>
	    <td>상세주소</td>
	    <td>받는분</td>
	    <td>연락처1</td>
	    <td>연락처2</td>
	    <td>주소지</td>
	    <td>상세주소</td>
	    <td>물품</td>
	    <td>포장</td>
	    <td>개수</td>
	    <td>배송타입</td>
	    <td>픽업요청일시</td>
	    <td>배송요청메모</td>
	</tr>
	<% for i=0 to objbaljumakeonorder.FResultCount -1 %>
	<tr>
		<td><%= objbaljumakeonorder.FItemList(i).FOrderSerial %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fbuyname %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fbuyphone %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fbuyaddr1 %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fbuyaddr2 %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Freqname %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Freqhp %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Freqphone %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Freqzipaddr %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Freqaddress %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fitemname %></td>
		<td><%= objbaljumakeonorder.FItemList(i).FpojangName %></td>
		<td><%= objbaljumakeonorder.FItemList(i).FboxNo %></td>
		<td></td>
		<td><%= objbaljumakeonorder.FItemList(i).FpickupReqDate %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fcomment %></td>
	</tr>
	<% next %>
</table>
</body>
</html>
	<%
	dbget.close : response.end
end if

%>
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<script language='javascript'>

function jsPopXL() {
	var popwin = window.open("pop_QuickOrder_list.asp?menupos=44&xl=Y&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>","jsPopXL","width=200 height=100 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cOrderFin(detailidx){
    if (confirm('취소 처리 확인 하시겠습니까?')){
        var popwin = window.open("/admin/ordermaster/misendmaster_main_process.asp?detailidx=" + detailidx + "&mode=cancelFin","misendmaster_process","width=100 height=100 scrollbars=yes resizable=yes");
	    popwin.focus();
    }
}
</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			퀵배송 송장출력일자 : <% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>

<p>

* 최대 500개까지 표시됩니다. <input type="button" class="button" value="엑셀받기" onclick="jsPopXL()">

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmview" method="get">
	<input type="hidden" name="iid" value="">
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18">
			검색결과 : <b><%= objbaljumakeonorder.FResultCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	    <td width="70">송장출력일</td>
		<td width="70">고객주문번호</td>
	    <td width="70">보내는분</td>
	    <td width="70">연락처1</td>
	    <td width="70">주소지</td>
	    <td width="70">상세주소</td>
	    <td width="70">받는분</td>
	    <td width="70">연락처1</td>
	    <td width="70">연락처2</td>
	    <td width="70">주소지</td>
	    <td width="70">상세주소</td>
	    <td width="70">물품</td>
	    <td width="70">포장</td>
	    <td width="70">개수</td>
	    <td width="70">배송타입</td>
	    <td width="70">픽업요청일시</td>
	    <td width="70">배송요청메모</td>
		<td>비고</td>
	</tr>
	<% if objbaljumakeonorder.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF" height="25">
	  	<td colspan="18" align="center">검색결과가 없습니다.</td>
	</tr>
	<% else %>
	<% for i=0 to objbaljumakeonorder.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td><%= objbaljumakeonorder.FItemList(i).Fsongjangprintdate %></td>
		<td><%= objbaljumakeonorder.FItemList(i).FOrderSerial %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fbuyname %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fbuyphone %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fbuyaddr1 %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fbuyaddr2 %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Freqname %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Freqhp %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Freqphone %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Freqzipaddr %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Freqaddress %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fitemname %></td>
		<td><%= objbaljumakeonorder.FItemList(i).FpojangName %></td>
		<td><%= objbaljumakeonorder.FItemList(i).FboxNo %></td>
		<td></td>
		<td><%= objbaljumakeonorder.FItemList(i).FpickupReqDate %></td>
		<td><%= objbaljumakeonorder.FItemList(i).Fcomment %></td>
		<td></td>
	</tr>
  <% next %>
  <% end if %>
</table>


<%
set objbaljumakeonorder = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
