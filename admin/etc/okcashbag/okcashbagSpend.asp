<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/othermall/okcashbagCls.asp"-->

<%
dim sSdate,sEdate, userid, orderserial
sSdate 		= requestCheckVar(Request("iSD"),10)
sEdate 		= requestCheckVar(Request("iED"),10)
userid 		= requestCheckVar(Request("uId"),32)
orderserial	= requestCheckVar(Request("oSn"),12)

IF sSdate ="" Then
	sSdate= DateSerial(Year(now()),Month(now()),1)
End IF

dim OrderType
OrderType = requestCheckVar(Request("otp"),2)

IF OrderType="" Then OrderType="N"

dim sPageSize
sPageSize = requestCheckVar(Request("ps"),3)
IF sPageSize="" Then sPageSize = 50
IF OrderType="UN" or OrderType ="UC" Then sPageSize=1000
dim SearchDateType
SearchDateType = requestCheckVar(request("dType"),2)
IF SearchDateType="" THEN SearchDateType="od"

dim CurrPage
CurrPage = requestCheckVar(request("pg"),3)

IF CurrPage="" THEN CurrPage =1
dim oCash,intLp
Set oCash = New CashbagCls
oCash.FCurrPage		= CurrPage
oCash.FPageSize		= sPageSize
oCash.FStartDate 	= sSdate
oCash.FEndDate 		= sEdate
oCash.Fuserid	 	= userid
oCash.Forderserial 	= orderserial
oCash.FOrderType 	= OrderType
oCash.FSearchType	= SearchDateType

//사용건수
oCash.getSpendCashbagList

%>

<script language='javascript'>
function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName+'&FN=sfrm','pCal','width=250, height=200');
		winCal.focus();
	}

function NextPage(page){
    sfrm.pg.value = page;
    sfrm.submit();
}
</script>
<!-- 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="sfrm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="pg" >
	<input type="hidden" name="ps" value="<%= sPageSize %>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<select name="dType">
				<option value="od" <% IF SearchDateType="od" Then response.write "selected"%>>주문일 기준</option>
				<option value="ov" <% IF SearchDateType="ov" Then response.write "selected"%>>배송일 기준</option>
				<!--<option value="ud" <% IF SearchDateType="ud" Then response.write "selected"%>>적립일 기준</option>-->
			</select>
			<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
			~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;"> &nbsp;
			<!-- 아이디 <input type="text" size="10" maxlength="32" name="uId" value="<%=userid%>"> &nbsp; -->
			주문번호 <input type="text" size="12" maxlength="12" name="oSn" value="<%=orderserial%>">
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.sfrm.submit();">
		</td>
	</tr>

    </form>
</table>
<!-- 검색폼 끝 -->
<p>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("sky") %>">
	<td align="center" width="100" >주문번호</td>
	<td align="center">주문일자</td>
	<td align="center">배송일자</td>
	<td align="center">주문자</td>
	<td align="center">총결제금액</td>
	<td align="center">사용포인트</td>
	<td align="center">적립포인트</td>
</tr>
<% IF oCash.FResultcount<=0 Then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center"> 일치하는 데이타가 없습니다.</td>
	</tr>
<% ELSE %>

	<% FOR intLp=0 To oCash.FResultcount-1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= oCash.FList(IntLp).FOrderSerial %></td>
		<td align="center"><%= DateValue(oCash.FList(IntLp).FRegdate) %></td>
		<td align="center"><% if DateValue(oCash.FList(IntLp).FBeadaldate)="1900-01-01" then Response.Write "미배송": Else Response.Write DateValue(oCash.FList(IntLp).FBeadaldate): End if %></td>
		<td align="center"><%= oCash.FList(IntLp).FBuyName %></td>
		<td align="center"><%= FormatNumber(oCash.FList(IntLp).FsubtotalPrice,0) %></td>
		<td align="center"><%= FormatNumber(oCash.FList(IntLp).Facctamount,0) %></td>
		<td align="center"><%= FormatNumber(oCash.FList(IntLp).FGainPoint,0) %></td>
	</tr>

	<% NEXT %>
<% End IF %>
</table>
<!-- 하단 페이징 -->
<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if oCash.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCash.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for intLp=0 + oCash.StartScrollPage to oCash.FScrollCount + oCash.StartScrollPage - 1 %>
			<% if intLp>oCash.FTotalpage then Exit for %>
			<% if CStr(CurrPage)=CStr(intLp) then %>
			<font color="red">[<%= intLp %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= intLp %>')">[<%= intLp %>]</a>
			<% end if %>
		<% next %>

		<% if oCash.HasNextScroll then %>
			<a href="javascript:NextPage('<%= intLp %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

</table>
<%
set oCash = Nothing
%>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
