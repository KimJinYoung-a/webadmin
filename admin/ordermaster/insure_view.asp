<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/myorder_Insurecls.asp"-->
<%
	'// 변수 선언 //
	dim OrderIdx
	dim page, searchDiv, searchKey, searchString, param

	dim oInsure, i, lp

	'// 파라메터 접수 //
	OrderIdx = request("OrderIdx")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수

	'// 내용 접수
	set oInsure = new CInsure
	oInsure.FRectOrderIdx = OrderIdx

	oInsure.GetInsureRead

%>
<script language="javascript">
<!--
	// 보증서 삭제
	function GotoInsureDel(){
		if (confirm('전자보증서 정보를 삭제하시겠습니까?\n\n※ 10x10의 정보에서만 삭제되는 것이므로 실제 처리는 U-Safe에서 반드시 확인해주십시요.')){
			document.frm_trans.mode.value="Del";
			document.frm_trans.submit();
		}
	}

	// 전자보증서 팝업
	function insurePrint(iorderserial, mallid)
	{
		var receiptUrl = "https://gateway.usafe.co.kr/esafe/ResultCheck.asp?oinfo=" + iorderserial + "|" + mallid
		window.open(receiptUrl,"insurePop","width=720,height=500,scrollbars=yes");
	}
//-->
</script>
<!-- 전자보증서 정보 시작 -->
<table width="600" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4" height="24" align="left"><b>전자보증서 상세 정보</b></td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
		<td bgcolor="#FFFFFF" width="180"><%=oInsure.FInsureList(0).Forderserial %></td>
		<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">구매일자</td>
		<td bgcolor="#FFFFFF"><%=FormatDate(oInsure.FInsureList(0).Fregdate,"0000.00.00")%></td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">주문 품목</td>
		<td bgcolor="#F8F8FF" colspan="3">
			<%=db2html(oInsure.FInsureList(0).Fitemname)%>
			<% if Not(oInsure.FInsureList(0).Fipkumdate="" or isnull(oInsure.FInsureList(0).Fipkumdate)) then %>(입금일 : <%=FormatDate(oInsure.FInsureList(0).Fipkumdate,"0000.00.00")%>)<% end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">정산 금액</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= CurrFormat(oInsure.FInsureList(lp).FsubtotalPrice) & "원"%></td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">구매자</td>
		<td bgcolor="#FFFFFF" colspan="3"><%=oInsure.FInsureList(0).Fbuyname %></td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">구매자 전화</td>
		<td bgcolor="#FFFFFF"><%=db2html(oInsure.FInsureList(0).Fbuyphone)%></td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">구매자 휴대폰</td>
		<td bgcolor="#FFFFFF"><%=db2html(oInsure.FInsureList(0).Fbuyhp)%></td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">구매자 이메일</td>
		<td bgcolor="#FFFFFF" colspan="3"><%=db2html(oInsure.FInsureList(0).Fbuyemail)%></td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">주문 상태</td>
		<td bgcolor="#F8F8FF" colspan="3"><%=NormalIpkumDivName(oInsure.FInsureList(0).Fipkumdiv)%></td>
	</tr>
	<tr><td height="1" colspan="4" bgcolor="#D0D0D0"></td></tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">발행 결과</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<%
				'발행결과
				if oInsure.FInsureList(0).FinsureCd="0" then
			%>
					<font color=darkblue>정상</font>
					&nbsp;
					<input type="button" class="button" value="전자보증서 출력" onClick="insurePrint('<%=oInsure.FInsureList(0).Forderserial%>','ZZcube1010')">
			<%	else %>
					<font color=darkred>오류</font>
			<%	end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">보증서번호(결과)</td>
		<td bgcolor="#FFFFFF" colspan="3"><%=oInsure.FInsureList(0).FinsureMsg%></td>
	</tr>
	<tr><td height="1" colspan="4" bgcolor="#D0D0D0"></td></tr>
	<tr>
		<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
			<input type="button" class="button" value="삭제" onClick="GotoInsureDel()">
			&nbsp;
			<input type="button" class="button" value="목록" onClick="self.location='Insure_list.asp?menupos=<%=menupos & param %>'">
		</td>
	</tr>
<form name="frm_trans" method="POST" action="doInsure.asp">
<input type="hidden" name="OrderIdx" value="<%=OrderIdx%>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
</form>
</table>
<!-- 세금계산 요청서 정보 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
