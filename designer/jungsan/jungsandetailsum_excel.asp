<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim id,gubun
id = request("id")
gubun = request("gubun")

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectId = id
ojungsan.FRectgubun = gubun
ojungsan.FRectdesigner = session("ssBctID")
ojungsan.JungsanMasterList

if ojungsan.FresultCount <1 then
	dbget.close()	:	response.End
end if

dim gubunstr
if (gubun = "upche") then
	gubunstr = "업체배송"
elseif (gubun = "maeip") then
	gubunstr = "매입"
elseif (gubun = "witaksell") then
	gubunstr = "특정"
elseif (gubun = "witakchulgo") then
	gubunstr = "기타출고"
end if


%>
<!-- 엑셀파일로 저장 헤더 부분 -->
<%
Response.ContentType = "application/unknown"
Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=" & "온라인 " & ojungsan.FItemList(0).Ftitle & " " & gubunstr & ".xls"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
/* 엑셀 다운로드로 저장시 숫자로 표시될 경우 방지 */
.txt {mso-number-format:'\@'}
</style>
</head>
<body>



<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100">구분</td>
		<td width="100">총건수</td>
		<td width="100">소비자가총액</td>
		<td width="100">공급가총액</td>
		<td width="70">평균마진</td>
		<% if gubun="maeip" then %>
		<td colspan=4>비고</td>
		<% else %>
		<td colspan=6>비고</td>
		<% end if %>
	</tr>
	<% if gubun="upche" then %>
	<tr bgcolor="#CCCCFF">
		<td>업체배송</td>
		<td align=right><%= ojungsan.FItemList(0).Fub_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fub_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fub_totalsuplycash/ojungsan.FItemList(0).Fub_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=center></td>
		<% end if %>
		<td colspan=6><%= nl2br(ojungsan.FItemList(0).Fub_comment) %></td>
	</tr>
	<% end if %>
	<% if gubun="maeip" then %>
	<tr bgcolor="#CCCCFF">
		<td>매입</td>
		<td align=right><%= ojungsan.FItemList(0).Fme_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fme_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fme_totalsuplycash/ojungsan.FItemList(0).Fme_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=center></td>
		<% end if %>
		<td colspan=4><%= nl2br(ojungsan.FItemList(0).Fme_comment) %></td>
	</tr>
	<% end if %>
	<% if gubun="witaksell" then %>
	<tr bgcolor="#CCCCFF">
		<td>특정</td>
		<td align=right><%= ojungsan.FItemList(0).Fwi_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fwi_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fwi_totalsuplycash/ojungsan.FItemList(0).Fwi_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=center></td>
		<% end if %>
		<td colspan=6><%= nl2br(ojungsan.FItemList(0).Fwi_comment) %></td>
	</tr>
	<% end if %>
	<!--
	<tr bgcolor="#FFFFFF">
		<td>특정 오프라인</td>
		<td><%= ojungsan.FItemList(0).Fsh_cnt %></td>
		<td><%= FormatNumber(ojungsan.FItemList(0).Fsh_totalsellcash,0) %></td>
		<td><%= FormatNumber(ojungsan.FItemList(0).Fsh_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fsh_totalsellcash<>0 then %>
		<td><%= CLng((1-ojungsan.FItemList(0).Fsh_totalsuplycash/ojungsan.FItemList(0).Fsh_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td>?</td>
		<% end if %>
		<td><%= nl2br(ojungsan.FItemList(0).Fsh_comment) %></td>
		<td align="center"><img src="/images/icon_search.jpg" width="16" border="0"></a></td>
	</tr>
	-->
	<% if gubun="witakchulgo" then %>
	<tr bgcolor="#CCCCFF">
		<td>기타출고</td>
		<td align=right><%= ojungsan.FItemList(0).Fet_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fet_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fet_totalsuplycash/ojungsan.FItemList(0).Fet_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=right></td>
		<% end if %>
		<td colspan=6><%= nl2br(ojungsan.FItemList(0).Fet_comment) %></td>
	</tr>
	<% end if %>
</table>


<p>

<%
set ojungsan = Nothing


dim ojungsansummary
set ojungsansummary = new CUpcheJungsan
ojungsansummary.FRectId = id
ojungsansummary.FRectgubun = gubun
ojungsansummary.FRectdesigner = session("ssBctID")

'' 1357 이전내역은 정산방식이 다름(재고기준정산)
if (id>1357) and (gubun<>"") then
    ojungsansummary.JungsanDetailListSum
end if
%>

<!-- 아이템별 합계 리스트 시작-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
		<% if gubun="maeip" then %>
		<td colspan="9" align="left">
		<% else %>
		<td colspan="11" align="left">
		<% end if %>

			<b>상품(아이템)별 합계리스트</b>
			&nbsp;&nbsp;
			<% if ojungsansummary.FRectgubun="maeip" then %>
			창고입고확인일 기준으로 등록됩니다.
			<% else %>
			출고일 기준으로 등록됩니다.
			<% end if %>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">상품코드</td>
		<td colspan=3>상품명</td>
		<% if gubun="maeip" then %>
		<td>옵션명</td>
		<% else %>
		<td colspan=3>옵션명</td>
		<% end if %>
		<td width="40">수량</td>
		<td width="70">판매가</td>
		<td width="70">공급가</td>
		<td width="80">공급가합계</td>
    </tr>
<% if ojungsansummary.FResultCount>0 and ojungsansummary.FRectgubun<>"" then %>
    <% suplytotalsum=0 %>
    <% for i=0 to ojungsansummary.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsansummary.FItemList(i).Fsuplycash * ojungsansummary.FItemList(i).FItemNo
    suplytotalsum = suplytotalsum + suplysum

    %>
    <tr bgcolor="#FFFFFF" align="center">
      <td class="txt"><%= ojungsansummary.FItemList(i).FItemID %></td>
      <td align="left" class="txt" colspan=3><%= ojungsansummary.FItemList(i).FItemName %></td>
		<% if gubun="maeip" then %>
		<td class="txt"><%= ojungsansummary.FItemList(i).FItemOptionName %></td>
		<% else %>
		<td class="txt" colspan=3><%= ojungsansummary.FItemList(i).FItemOptionName %></td>
		<% end if %>
      <td><%= ojungsansummary.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsansummary.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsansummary.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsansummary.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsansummary.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td align="center">합계</td>
      <td colspan="7"></td>
		<% if gubun="maeip" then %>
		<% else %>
		<td colspan="2"></td>
		<% end if %>
      <td align="right"><font color="<%= MinusFont(suplytotalsum) %>"><%= FormatNumber(suplytotalsum,0) %></font></td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<% if gubun="maeip" then %>
    	<td colspan="9" align="center">&nbsp;검색내역이 없습니다.</td>
    	<% else %>
    	<td colspan="11" align="center">&nbsp;검색내역이 없습니다.</td>
    	<% end if %>
    </tr>
<% end if %>
</table>
<!-- 아이템별 합계 리스트 끝-->
<p>


<%
set ojungsansummary = Nothing


dim i, suplysum, suplytotalsum, duplicated
dim sumttl1, sumttl2
sumttl1 = 0
sumttl2 = 0

dim ojungsandetail
set ojungsandetail = new CUpcheJungsan
ojungsandetail.FRectId = id
ojungsandetail.FRectgubun = gubun
ojungsandetail.FRectdesigner = session("ssBctID")
ojungsandetail.FRectOrder = "orderserial"


'' 1357 이전내역은 정산방식이 다름(재고기준정산)
if (id>1357) and (gubun<>"")   then
    ojungsandetail.JungsanDetailList
end if
%>
<!-- 주문건별 리스트 시작-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
		<% if gubun="maeip" then %>
		<td colspan="9" align="left">
		<% else %>
		<td colspan="11" align="left">
		<% end if %>

			<b>주문/출고/입고건별 상세리스트</b>
			&nbsp;&nbsp;
			<% if ojungsandetail.FRectgubun="maeip" then %>
			창고입고확인일 기준으로 등록됩니다.
			<% else %>
			출고일 기준으로 등록됩니다.
			<% end if %>

			<% if ojungsandetail.FResultCount>=5000 then %>
			(최대 <%= ojungsandetail.FResultCount %> 건 표시)
			<% end if %>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <% if ojungsandetail.FRectgubun="maeip" then %>
      <td width="70">입고코드</td>
      <% elseif ojungsandetail.FRectgubun="witakchulgo" then %>
      <td width="70">출고코드</td>
      <% else %>
      <td width="70">주문번호</td>
      <% end if %>

      <% if (ojungsandetail.FRectgubun<>"maeip") and (ojungsandetail.FRectgubun<>"witakchulgo") then %>
      <td width="45">구매자</td>
      <td width="45">수령인</td>
      <% elseif (ojungsandetail.FRectgubun="witakchulgo") then %>
      <td width="45"></td>
      <td width="45"></td>
      <% end if %>
      <td colspan=2>상품명</td>
      <td>옵션명</td>
      <td width="35">수량</td>
      <td width="50">판매가</td>
      <td width="50">공급가</td>
      <td width="65">공급가계</td>

      <% if ojungsandetail.FRectgubun="maeip" then %>
      <td width="65">입고일</td>
      <% else %>
      <td width="65">출고일</td>
      <% end if %>
    </tr>
<% if ojungsandetail.FResultCount>0 and ojungsandetail.FRectgubun<>"" then %>
    <% for i=0 to ojungsandetail.FResultCount-1 %>

    <%
	sumttl1 = sumttl1 + ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsellcash
	sumttl2 = sumttl2 + ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash
	%>
    <tr bgcolor="#FFFFFF" align="center">
      <td class="txt"><%= ojungsandetail.FItemList(i).Fmastercode %></td>
      <% if ojungsandetail.FRectgubun<>"maeip" and ojungsandetail.FRectgubun<>"witakchulgo" then %>
      <td><%= ojungsandetail.FItemList(i).FBuyname %></td>
      <td><%= ojungsandetail.FItemList(i).FReqname %></td>
      <% elseif (ojungsandetail.FRectgubun="witakchulgo") then %>
      <td><%= ojungsandetail.FItemList(i).FBuyname %></td>
      <td><%= ojungsandetail.FItemList(i).FReqname %></td>
      <% end if %>
      <td align="left" class="txt" colspan=2><%= ojungsandetail.FItemList(i).FItemName %></td>
      <td class="txt"><%= ojungsandetail.FItemList(i).FItemOptionName %></td>
      <td><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo) %>"><%= ojungsandetail.FItemList(i).FItemNo %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>

      <td><%= ojungsandetail.FItemList(i).FExecDate %></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF" align="center">
    	<td>합계</td>
    	<% if gubun="maeip" then %>
    	<td colspan="6"></td>
    	<% else %>
    	<td colspan="8"></td>
    	<% end if %>
    	<td align="right"><font color="<%= MinusFont(sumttl2) %>"><%= formatNumber(sumttl2,0) %></font></td>
    	<td></td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<% if gubun="maeip" then %>
    	<td colspan="9" align="center">&nbsp;검색내역이 없습니다.</td>
    	<% else %>
    	<td colspan="11" align="center">&nbsp;검색내역이 없습니다.</td>
    	<% end if %>
    </tr>
<% end if %>
</table>
<!-- 주문건별 리스트 끝-->

<%
set ojungsandetail = Nothing
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
