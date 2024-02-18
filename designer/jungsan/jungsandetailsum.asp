<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
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

%>

<script language='javascript'>

function JunsanDetailList(id,gubun){
    location.href = '?id=' + id + '&gubun=' + gubun;
}
function ExcelJunsanDetailList(id,gubun){
    location.href = '/designer/jungsan/jungsandetailsum_excel.asp?id=' + id + '&gubun=' + gubun;
}
</script>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">
        	<b>
        	온라인 <%= ojungsan.FItemList(0).Ftitle %>&nbsp;[<%= ojungsan.FItemList(0).Fdesignerid %>]
        	&nbsp;&nbsp;|&nbsp;&nbsp;
            <%= ojungsan.FItemList(0).Fdifferencekey %> 차
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <font color="<%= ojungsan.FItemList(0).GetTaxtypeNameColor %>"><%= ojungsan.FItemList(0).GetSimpleTaxtypeName %></font>&nbsp;&nbsp;
            </b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100">구분</td>
		<td width="100">총건수</td>
		<td width="100">소비자가총액</td>
		<td width="100">공급가총액</td>
		<td width="70">평균마진</td>
		<td>비고</td>
		<td width="50">상세내역</td>
	</tr>
	<% if gubun="upche" then %>
	<tr bgcolor="#CCCCFF">
	<% else %>
	<tr bgcolor="#FFFFFF">
	<% end if %>
		<td>업체배송</td>
		<td align=right><%= ojungsan.FItemList(0).Fub_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fub_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fub_totalsuplycash/ojungsan.FItemList(0).Fub_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=center></td>
		<% end if %>
		<td><%= nl2br(ojungsan.FItemList(0).Fub_comment) %></td>
		<td align="center">
		  <a href="javascript:JunsanDetailList('<%= id %>','upche')"><img src="/images/icon_search.jpg" width="16" border="0"></a>
		  <a href="javascript:ExcelJunsanDetailList('<%= id %>','upche')"><img src="/images/iexcel.gif" width="16" border="0"></a>
		</td>
	</tr>
	<% if gubun="maeip" then %>
	<tr bgcolor="#CCCCFF">
	<% else %>
	<tr bgcolor="#FFFFFF">
	<% end if %>
		<td>매입</td>
		<td align=right><%= ojungsan.FItemList(0).Fme_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fme_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fme_totalsuplycash/ojungsan.FItemList(0).Fme_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=center></td>
		<% end if %>
		<td><%= nl2br(ojungsan.FItemList(0).Fme_comment) %></td>
		<td align="center">
		  <a href="javascript:JunsanDetailList('<%= id %>','maeip')"><img src="/images/icon_search.jpg" width="16" border="0"></a>
		  <a href="javascript:ExcelJunsanDetailList('<%= id %>','maeip')"><img src="/images/iexcel.gif" width="16" border="0"></a>
		</td>
	</tr>
	<% if gubun="witaksell" then %>
	<tr bgcolor="#CCCCFF">
	<% else %>
	<tr bgcolor="#FFFFFF">
	<% end if %>
		<td>특정</td>
		<td align=right><%= ojungsan.FItemList(0).Fwi_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fwi_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fwi_totalsuplycash/ojungsan.FItemList(0).Fwi_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=center></td>
		<% end if %>
		<td><%= nl2br(ojungsan.FItemList(0).Fwi_comment) %></td>
		<td align="center">
		  <a href="javascript:JunsanDetailList('<%= id %>','witaksell')"><img src="/images/icon_search.jpg" width="16" border="0"></a>
		  <a href="javascript:ExcelJunsanDetailList('<%= id %>','witaksell')"><img src="/images/iexcel.gif" width="16" border="0"></a>
		</td>
	</tr>
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
	<% else %>
	<tr bgcolor="#FFFFFF">
	<% end if %>
		<td>기타출고</td>
		<td align=right><%= ojungsan.FItemList(0).Fet_cnt %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsuplycash,0) %></td>
		<% if ojungsan.FItemList(0).Fet_totalsellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).Fet_totalsuplycash/ojungsan.FItemList(0).Fet_totalsellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=right></td>
		<% end if %>
		<td><%= nl2br(ojungsan.FItemList(0).Fet_comment) %></td>
		<td align="center">
		  <a href="javascript:JunsanDetailList('<%= id %>','witakchulgo')"><img src="/images/icon_search.jpg" width="16" border="0"></a>
		  <a href="javascript:ExcelJunsanDetailList('<%= id %>','witakchulgo')"><img src="/images/iexcel.gif" width="16" border="0"></a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td>총계</td>
		<td></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).GetTotalSellcash,0) %></td>
		<td align=right><%= FormatNumber(ojungsan.FItemList(0).GetTotalSuplycash,0) %></td>
		<% if ojungsan.FItemList(0).GetTotalSellcash<>0 then %>
		<td align=center><%= CLng((1-ojungsan.FItemList(0).GetTotalSuplycash/ojungsan.FItemList(0).GetTotalSellcash)*10000)/100 %> %</td>
		<% else %>
		<td align=right></td>
		<% end if %>
		<td></td>
		<td></td>
	</tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
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
		<td colspan="10" align="left">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
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
		<td>상품명</td>
		<td>옵션명</td>
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
      <td><%= ojungsansummary.FItemList(i).FItemID %></td>
      <td align="left"><%= ojungsansummary.FItemList(i).FItemName %></td>
      <td><%= ojungsansummary.FItemList(i).FItemOptionName %></td>
      <td><%= ojungsansummary.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsansummary.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsansummary.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsansummary.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsansummary.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td align="center">합계</td>
      <td colspan="5"></td>
      <td align="right"><font color="<%= MinusFont(suplytotalsum) %>"><%= FormatNumber(suplytotalsum,0) %></font></td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="10" align="center"><img src="/images/icon_search.jpg" width="16" border="0" align="absbottom">&nbsp;상세내역을 선택하세요.</td>
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
		<td colspan="10" align="left">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
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
      <td>상품명</td>
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
      <td><%= ojungsandetail.FItemList(i).Fmastercode %></td>
      <% if ojungsandetail.FRectgubun<>"maeip" and ojungsandetail.FRectgubun<>"witakchulgo" then %>
      <td><%= ojungsandetail.FItemList(i).FBuyname %></td>
      <td><%= ojungsandetail.FItemList(i).FReqname %></td>
      <% elseif (ojungsandetail.FRectgubun="witakchulgo") then %>
      <td><%= ojungsandetail.FItemList(i).FBuyname %></td>
      <td><%= ojungsandetail.FItemList(i).FReqname %></td>
      <% end if %>
      <td align="left"><%= ojungsandetail.FItemList(i).FItemName %></td>
      <td><%= ojungsandetail.FItemList(i).FItemOptionName %></td>
      <td><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo) %>"><%= ojungsandetail.FItemList(i).FItemNo %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>

      <td><%= ojungsandetail.FItemList(i).FExecDate %></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF" align="center">
    	<td>합계</td>
    	<% if ojungsandetail.FRectgubun="maeip" then %>
    	<td colspan="5"></td>
    	<% else %>
    	<td colspan="7"></td>
    	<% end if %>
    	<td align="right"><font color="<%= MinusFont(sumttl2) %>"><%= formatNumber(sumttl2,0) %></font></td>
    	<td></td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="10" align="center"><img src="/images/icon_search.jpg" width="16" border="0" align="absbottom">&nbsp;상세내역을 선택하세요.</td>
    </tr>
<% end if %>
</table>
<!-- 주문건별 리스트 끝-->

<%
set ojungsandetail = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->