<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : OkCashbag관리
' History : 서동석 생성
'			2023.03.22 한용민 수정(권한 수기 아이디 박혀 있는부분 공통 권한 변수로 자동화. 소스 표준코드로 수정.)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/othermall/okcashbagCls.asp"-->
<%
dim sSdate,sEdate, userid, orderserial, vRdSite, OrderType, sPageSize, SearchDateType, CurrPage
dim oCash,intLp
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)
	userid 		= requestCheckVar(Request("uId"),32)
	orderserial	= requestCheckVar(Request("oSn"),12)
	vRdSite		= requestCheckVar(Request("rdsite"),10)
	OrderType = requestCheckVar(Request("otp"),2)
	SearchDateType = requestCheckVar(request("dType"),2)
	CurrPage = requestCheckVar(request("pg"),3)

If vRdSite = "" Then
	vRdSite = "okcashbag"
End If

IF sSdate ="" Then
	sSdate= DateSerial(Year(now()),Month(now()),1)
End IF

IF OrderType="" Then OrderType="N"
sPageSize = 100
IF SearchDateType="" THEN SearchDateType="od"

IF CurrPage="" THEN CurrPage =1

Set oCash = New CashbagCls
	oCash.FCurrPage		= CurrPage
	oCash.FPageSize		= sPageSize
	oCash.FStartDate 	= sSdate
	oCash.FEndDate 		= sEdate
	oCash.Fuserid	 	= userid
	oCash.Forderserial 	= orderserial
	oCash.FOrderType 	= OrderType
	oCash.FSearchType	= SearchDateType
	oCash.FRdSite		= vRdSite

	IF OrderType="N" Then 		'//정상건
		oCash.getNormalOrder()
	ELSEIF OrderType ="C" Then	'//취소건
		oCash.getCancelOrder()
	ELSEIF OrderType="UN" or OrderType ="UC" Then '// 출력 된 내역 (정상,취소)
		oCash.getUpdatedOrder()
	END IF

%>

<script type="text/javascript">

function jsChkAll(blnChk){
		var frm, blnChk;
		frm = document.rfrm;

		for (var i=0;i<frm.elements.length;i++){
			//check optioon
			var e = frm.elements[i];

			//check itemEA
			if ((e.type=="checkbox")) {
				e.checked = blnChk ;
				AnCheckClick(e);
		}
	}
}

function downloadexcel(){
	document.sfrm.target = "view";
	document.sfrm.action = "/admin/etc/okcashbag/okcashbagReport_down.asp";
	document.sfrm.submit();
	document.sfrm.target = "";
	document.sfrm.action = "";
}

function NextPage(v){
	document.sfrm.target = "";
	document.sfrm.action = "";
	document.sfrm.pg.value=v;
	document.sfrm.submit();
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName+'&FN=sfrm','pCal','width=250, height=200');
	winCal.focus();
}

</script>

<!-- 검색 시작 -->
<form name="sfrm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="pg" value=1>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<select name="dType">
				<option value="od" <% IF SearchDateType="od" Then response.write "selected"%>>주문일 기준</option>
				<option value="ov" <% IF SearchDateType="ov" Then response.write "selected"%>>배송일 기준</option>
				<!--<option value="ud" <% IF SearchDateType="ud" Then response.write "selected"%>>적립일 기준</option>-->
			</select>
			<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
			~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;">
			&nbsp;
			* 아이디 : <input type="text" size="10" maxlength="32" name="uId" value="<%=userid%>">
			&nbsp;
			* 주문번호 : <input type="text" size="12" maxlength="12" name="oSn" value="<%=orderserial%>">
			&nbsp;
			* 제휴업체 : 
			<select name="rdsite">
				<option value="okcashbag" <%=ChkIIF(vRdSite="okcashbag","selected","")%>>okcashbag</option>
				<option value="pickle" <%=ChkIIF(vRdSite="pickle","selected","")%>>pickle</option>
			</select>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="NextPage('');">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<acronym title="정상건"><input type="radio" name="otp" value="N" <% IF OrderType="N" Then response.write "checked" %> onClick="document.sfrm.submit();">정상건</acronym>
			<acronym title="정상건 출력내역"><input type="radio" name="otp" value="UN" <% IF OrderType="UN" Then response.write "checked" %> onClick="document.sfrm.submit();">정상건 출력내역</acronym>
			<acronym title="정상건 출력후 취소내역"><input type="radio" name="otp" value="C" <% IF OrderType="C" Then response.write "checked" %> onClick="document.sfrm.submit();">취소건</acronym>
			<acronym title="취소건 출력내역"><input type="radio" name="otp" value="UC" <% IF OrderType="UC" Then response.write "checked" %> onClick="document.sfrm.submit();">취소건 출력내역</acronym>
		</td>
	</tr>
</table>
<!-- 검색 끝 -->

<Br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">
		<%
		' 관리자 이거나 개발운영팀 이거나 제휴파트 일경우
		If C_ADMIN_AUTH or C_SYSTEM_Part or C_partnership_part Then
		%>
			<input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button">
		<% else %>
			다운권한없음
		<% End If %>
	</td>
</tr>
</table>
</form>
<!-- 액션 끝 -->

<form name="rfrm" method="post" action="" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%= oCash.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= CurrPage %>/ <%= oCash.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width="20"><input type="checkbox" name="chkAll" onClick="jsChkAll(this.checked);"></td>
	<td align="center" width="100" >주문번호</td>
	<td align="center" width="80">장바구니번호</td>
	<td align="center">총결제금액</td>
	<td align="center">주문일자</td>
	<td align="center">배송일자</td>
	<td align="center">주문자</td>
	<td align="center">캐쉬백번호</td>
	<td align="center">적립포인트</td>
</tr>
<% if oCash.FresultCount>0 then %>
<% for IntLp=0 to oCash.FresultCount-1 %>

<tr align="center" bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="chkb" onClick="AnCheckClick(this);" value="<%= oCash.FItemList(IntLp).Fidx %>"></td>
	<td align="center"><%= oCash.FItemList(IntLp).FOrderSerial %></td>
	<td align="center"><%= oCash.FItemList(IntLp).FShoppingBagNo %></td>
	<td align="center"><%= FormatNumber(oCash.FItemList(IntLp).FPointCash,0) %></td>
	<td align="center">
		<% if not(isnull(oCash.FItemList(IntLp).FRegdate)) then %>
			<%= DateValue(oCash.FItemList(IntLp).FRegdate) %>
		<% end if %>
	</td>
	<td align="center"><% if DateValue(oCash.FItemList(IntLp).FBeadaldate)="1900-01-01" then Response.Write "미배송": Else Response.Write DateValue(oCash.FItemList(IntLp).FBeadaldate): End if %></td>
	<td align="center"><%= oCash.FItemList(IntLp).FBuyName %></td>
	<td align="center">0000-****-****-0000</td>
	<td align="center"><%= FormatNumber(oCash.FItemList(IntLp).FPoint,0) %></td>
</tr>   
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="9" align="center">
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
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="16" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</form>
<!-- 리스트 시작 -->

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set oCash = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
