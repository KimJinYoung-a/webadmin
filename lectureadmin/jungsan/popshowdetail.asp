<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim id, yyyymm,designerid
id = RequestCheckvar(request("id"),10)

dim ojungsan, ojungsanmaster
set ojungsanmaster = new CUpcheJungsan
ojungsanmaster.FRectId = id
ojungsanmaster.FrectDesigner = session("ssBctID")
ojungsanmaster.JungsanMasterList

if ojungsanmaster.FresultCount <1 then
	dbget.close()	:	response.End
end if

yyyymm = ojungsanmaster.FItemList(0).FYYYYmm
designerid = ojungsanmaster.FItemList(0).FDesignerid

set ojungsan = new CUpcheJungsan
ojungsan.FRectid = id
ojungsan.FrectDesigner = session("ssBctID")
'ojungsan.FRectgubun = "upche"
'if (id>=179504) then ''2014/02 수정
'  ojungsan.FRectgubun = "lecture"
'end if
ojungsan.JungsanDetailListSum

dim i, suplysum, suplytotalsum, duplicated

suplytotalsum = 0
%>
<table width="760" cellspacing="0" class="a">
<tr>
  <td align="right"><a href="jungsanmaster.asp?menupos=<%= menupos %>&id=<%= id %>">내역확인&gt;&gt;</a></td>
</tr>
</table>
<% if ojungsan.FResultCount>0 then %>
<br>
<table border="0" width="760" class="a">
<tr>
	<td>[ 업체배송 아이템별 합계 ]</td>
	<td align="right">합계 <%= FormatNumber(ojungsanmaster.FitemList(0).Fub_totalsuplycash,0) %></td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="40">상품ID</td>
      <td width="200">상품명</td>
      <td width="80">옵션명</td>
      <td width="40">갯수</td>
      <td width="70">판매가</td>
      <td width="70">공급가</td>
      <td width="70">공급가합계</td>
    </tr>
    <% suplytotalsum=0 %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).Fsuplycash * ojungsan.FItemList(i).FItemNo
    suplytotalsum = suplytotalsum + suplysum

    duplicated = ojungsan.CheckDuplicated(i)
    %>

	<% if duplicated then %>
    <tr bgcolor="#FFFFFF">
	<% else %>
    <tr bgcolor="#FFFFFF">
    <% end if %>
      <td ><%= ojungsan.FItemList(i).FItemID %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="6"></td>
      <td align="right"><%= FormatNumber(suplytotalsum,0) %></td>
    </tr>
</table>
<% end if %>

<% if ojungsan.FResultCount>0 then %>
<%
ojungsan.FRectOrder = "orderserial"
ojungsan.JungsanDetailList
%>
<br>
<table border="0" width="760" class="a">
<tr>
	<td>[ 업체배송 내역 ] - (<font color="#FF0000">배송일 기준</font>입니다. 배송일이 다음달인 경우 다음달 정산에 포함됩니다.)</td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="80">주문번호</td>
      <td width="50">구매자</td>
      <td width="50">수령인</td>
      <td width="120">아이템명</td>
      <td width="80">옵션명</td>
      <td width="40">갯수</td>
      <td width="70">판매가</td>
      <td width="70">공급가</td>
      <td width="100">배송일</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td ><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td ><%= ojungsan.FItemList(i).FBuyname %></td>
      <td ><%= ojungsan.FItemList(i).FReqname %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td ><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><%= ojungsan.FItemList(i).FExecDate %></td>
    </tr>
    <% next %>
</table>
<% end if %>

<%
ojungsan.FRectgubun = "maeip"
ojungsan.JungsanDetailListSum
%>
<% if ojungsan.FResultCount>0 then %>
<br>
<table border="0" width="760" class="a">
<tr>
	<td>[ 매입입고 상품별 합계 ]</td>
	<td align="right">합계 <%= FormatNumber(ojungsanmaster.FitemList(0).Fme_totalsuplycash,0) %></td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1" class="a" align="center" bgcolor=#3d3d3d>
    <tr align="center" bgcolor="#DDDDFF">
      <td width="40">상품ID</td>
      <td>상품명</td>
      <td width="80">옵션명</td>
      <td width="40">수량</td>
      <td width="70">판매가</td>
      <td width="70">공급가</td>
      <td width="80">공급가합계</td>

    </tr>
    <% suplytotalsum=0 %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).Fsuplycash * ojungsan.FItemList(i).FItemNo
    suplytotalsum = suplytotalsum + suplysum

    duplicated = ojungsan.CheckDuplicated(i)
    %>

	<% if duplicated then %>
	<tr bgcolor="#FFFFFF">
	<% else %>
    <tr bgcolor="#FFFFFF">
    <% end if %>
      <td align="center"><%= ojungsan.FItemList(i).FItemID %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="6"></td>
      <td align="right"><%= FormatNumber(suplytotalsum,0) %></td>
    </tr>
</table>

<%
ojungsan.FRectOrder = "orderserial"
ojungsan.JungsanDetailList
%>
<table border="0" width="760" class="a">
<tr>
	<td>[ 매입입고 상세내역 ]</td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1" class="a" align="center" bgcolor=#3d3d3d>
    <tr align="center" bgcolor="#DDDDFF">
      <td width="60">입고코드</td>
      <td width="80">입고일</td>
      <td>상품명</td>
      <td width="80">옵션명</td>
      <td width="40">수량</td>
      <td width="70">판매가</td>
      <td width="70">공급가</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td align="center"><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td align="center"><%= ojungsan.FItemList(i).FExecDate %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
    </tr>
    <% next %>
</table>
<% end if %>


<% if id>1357 then %>
<!-- Modify 20040305 -->
<%
'ojungsan.FrectDesigner = designerid
'ojungsan.FRectStartDay = yyyymm + "-" + "01"
'ojungsan.FRectEndDay   = CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))+1,1))
'ojungsan.FRectYYYYMM   = yyyymm
'ojungsan.FRectPreYYYYMM   = Left(CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))-1,1)),7)

ojungsan.GetWitakJungSanSummary
%>

<% if ojungsan.FResultCount>0 then %>
<br>
<table border="0" width="760" class="a">
<tr>
	<td>[ 특정 아이템별 합계 ] - (<font color="#FF0000">배송완료일 기준</font>입니다. 배송완료일이 다음달인 경우 다음달 정산에 포함됩니다.)</td>
	<td align="right">합계 <%= FormatNumber(ojungsanmaster.FitemList(0).Fwi_totalsuplycash + ojungsanmaster.FitemList(0).Fsh_totalsuplycash + ojungsanmaster.FitemList(0).Fet_totalsuplycash,0) %></td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr align="center" bgcolor="#DDDDFF">
      <td width="40">상품ID</td>
      <td>상품명</td>
      <td width="80">옵션명</td>
      <td width="60">온라인판매</td>
      <td width="60">기타판매</td>
      <td width="60">판매가</td>
      <td width="60">공급가</td>
      <td width="80">공급가합계</td>
    </tr>
    <% suplytotalsum=0 %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).FSuplycash * (ojungsan.FItemList(i).Fsellno + ojungsan.FItemList(i).Foffsellno + ojungsan.FItemList(i).FChulgoNo)
    suplytotalsum = suplytotalsum + suplysum

    duplicated = ojungsan.CheckDuplicated(i)
    %>

	<% if duplicated then %>
    <tr bgcolor="#FFFFFF">
	<% else %>
    <tr bgcolor="#FFFFFF">
    <% end if %>
      <td align="center" ><%= ojungsan.FItemList(i).FItemID %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).Fsellno %></td>
      <td align="center"><%= ojungsan.FItemList(i).FChulgoNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FSellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).FSellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FSuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).FSuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="7"></td>
      <td align="right"><%= FormatNumber(suplytotalsum,0) %></td>
    </tr>
</table>
<%
ojungsan.FRectgubun = "witaksell"
ojungsan.FRectOrder = "orderserial"
ojungsan.JungsanDetailList

dim sumttl1, sumttl2
sumttl1 = 0
sumttl2 = 0
%>
<table border="0" width="760" class="a">
<tr>
	<td>[ 특정 온라인판매 상세내역 ] - (<font color="#FF0000">배송일 기준</font>입니다. 배송일이 다음달인 경우 다음달 정산에 포함됩니다.)</td>
</tr>
</table>
<% if ojungsan.FResultCount>0 then %>

<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr align="center" bgcolor="#DDDDFF">
      <td width="80">주문번호</td>
      <td width="50">구매자</td>
      <td width="50">수령인</td>
      <td>상품명</td>
      <td width="80">옵션명</td>
      <td width="40">수량</td>
      <td width="60">판매가</td>
      <td width="60">공급가</td>
      <td width="80">공급가계</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
	sumttl1 = sumttl1 + ojungsan.FItemList(i).FItemNo*ojungsan.FItemList(i).Fsellcash
	sumttl2 = sumttl2 + ojungsan.FItemList(i).FItemNo*ojungsan.FItemList(i).Fsuplycash
	%>
    <tr bgcolor="#FFFFFF">
      <td align="center" ><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td align="center" ><%= ojungsan.FItemList(i).FBuyname %></td>
      <td align="center" ><%= ojungsan.FItemList(i).FReqname %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center" ><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FItemNo) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FItemNo) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FItemNo) %>"><%= FormatNumber(ojungsan.FItemList(i).FItemNo*ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
    	<td colspan=8></td>
    	<td align=right><%= formatNumber(sumttl2,0) %></td>
    </tr>
</table>
<% else %>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
	<tr bgcolor="#FFFFFF"><td align=center>내역이 존재하지 않습니다.</td></tr>
</table>
<% end if %>

<% end if %>

<% else %>
<%
ojungsan.FrectDesigner = designerid
ojungsan.FRectStartDay = yyyymm + "-" + "01"
ojungsan.FRectEndDay   = CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))+1,1))
ojungsan.FRectYYYYMM   = yyyymm
ojungsan.FRectPreYYYYMM   = Left(CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))-1,1)),7)

ojungsan.GetWitakJungSanByItemView
%>
<% if ojungsan.FResultCount>0 then %>
<br>
<table border="0" width="760" class="a">
<tr>
	<td>[ 특정 아이템별 합계 ]</td>
	<td align="right">합계 <%= FormatNumber(ojungsanmaster.FitemList(0).Fwi_totalsuplycash,0) %></td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="40">상품ID</td>
      <td width="200">상품명</td>
      <td width="80">옵션명</td>
      <td width="40">출고/판매량</td>
      <td width="80">판매가</td>
      <td width="80">공급가</td>
      <td width="80">공급가합계</td>
    </tr>
    <% suplytotalsum=0 %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).FSuplycash_sell * ojungsan.FItemList(i).FjungsanNo
    suplytotalsum = suplytotalsum + suplysum

    duplicated = ojungsan.CheckDuplicated(i)
    %>

	<% if duplicated then %>
    <tr bgcolor="#FFFFFF">
	<% else %>
    <tr bgcolor="#FFFFFF">
    <% end if %>
      <td ><%= ojungsan.FItemList(i).FItemID %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FjungsanNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FSellcash_sell) %>"><%= FormatNumber(ojungsan.FItemList(i).FSellcash_sell,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FSuplycash_sell) %>"><%= FormatNumber(ojungsan.FItemList(i).FSuplycash_sell,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="6"></td>
      <td align="right"><%= FormatNumber(suplytotalsum,0) %></td>
    </tr>
</table>
<% end if %>
<% end if %>



<%
ojungsan.FRectOrder = "orderserial"
ojungsan.FRectgubun = "witakchulgo"
ojungsan.JungsanDetailList
%>
<% if ojungsan.FResultCount>0 then %>
<table border="0" width="760" class="a">
<tr>
	<td>[ 특정 기타판매 상세내역 ] - 판촉, 협찬 및 기타 판매</td>
</tr>
</table>
<table width="760" cellpadding="1" cellspacing="1"  class="a" align="center" bgcolor=#3d3d3d>
    <tr align="center" bgcolor="#DDDDFF">
      <td width="60">입고코드</td>
      <td width="80">출고일</td>
      <td>상품명</td>
      <td width="80">옵션명</td>
      <td width="40">수량</td>
      <td width="60">판매가</td>
      <td width="60">공급가</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td align="center"><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td align="center"><%= ojungsan.FItemList(i).FExecDate %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>

    </tr>
    <% next %>
</table>
<% end if %>
<%
set ojungsan = Nothing
set ojungsanmaster = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->