<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 taxRefund 관리
' History : 2014.01.17 서동석
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/payment/taxRefundMngCls.asp"-->
<%
dim page,shopid,yyyy1,mm1,dd1,yyyy2,mm2,dd2, fromDate,toDate, Searchtaxrefundkey, schType, scgRealsum, jyyyymm
dim datefg , i, ToTcashsum, intLoop, isedityn, inc3pl
	shopid = requestCheckvar(request("shopid"),32)
	page = requestCheckvar(request("page"),10)
	if page="" then page=1
	yyyy1 = requestCheckvar(request("yyyy1"),4)
	mm1 = requestCheckvar(request("mm1"),2)
	dd1 = requestCheckvar(request("dd1"),2)
	yyyy2 = requestCheckvar(request("yyyy2"),4)
	mm2 = requestCheckvar(request("mm2"),2)
	dd2 = requestCheckvar(request("dd2"),2)

	jyyyymm = requestCheckvar(request("jyyyymm"),7)

	datefg = requestCheckvar(request("datefg"),10)
    inc3pl = requestCheckvar(request("inc3pl"),10)
	Searchtaxrefundkey = requestCheckvar(request("Searchtaxrefundkey"),30)
	schType = requestCheckvar(request("schType"),10)
	scgRealsum = requestCheckvar(request("scgRealsum"),10)
if datefg = "" then datefg = "maechul"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-0)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/매장
if (C_IS_SHOP) then

	'//직영점일때
	if C_IS_OWN_SHOP then

		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if


dim oTaxRefund
set oTaxRefund = new CTaxRefund
	oTaxRefund.FRectShopID = shopid
	oTaxRefund.FRectStartDay = fromDate
	oTaxRefund.FRectEndDay = toDate
	oTaxRefund.frectdatefg = datefg
	oTaxRefund.frecttaxrefundkey = Searchtaxrefundkey
	oTaxRefund.frectscgRealsum = scgRealsum
	oTaxRefund.frectschType = schType
	if (Len(jyyyymm)=7) then
	    oTaxRefund.FRectRefundMonth = jyyyymm
    end if
	''oTaxRefund.FRectInc3pl = inc3pl
	oTaxRefund.FPageSize = 200
	oTaxRefund.FCurrPage = page

	if (shopid<>"") then
		oTaxRefund.GetTaxRefundTargetList
	else
		response.write "<script language='javascript'>"
		response.write "alert('매장을 선택하신 후 검색하세요.');"
		response.write "</script>"
	end if

dim totcnt, totrealsum, totVatsum

Dim defaultrefundCode
Select Case shopid
	'Case "streetshop011"	defaultrefundCode = "20025720390513"		'대학로
	Case "streetshop011"	defaultrefundCode = "20023120332514"		'대학로 '2014/02/05 김진영 수정. 강희란대리님 요청
	'Case "streetshop014"	defaultrefundCode = "20023120332513"		'두타
	Case "streetshop014"	defaultrefundCode = "20025720390514"		'두타 '2014/02/05 김진영 수정. 강희란대리님 요청
	'Case "streetshop018"	defaultrefundCode = "20025710131013"		'김포롯데
	Case "streetshop018"	defaultrefundCode = "20025710131014"		'김포롯데 '2014/02/05 김진영 수정. 강희란대리님 요청
End Select
%>
<script language='javascript'>
function addRefundKey(comp, chkid, btnIid, btnSid){
	document.getElementById(chkid).disabled = false;
	document.getElementById(btnIid).style.display = "none";
	document.getElementById(btnSid).style.display = "block";
	document.getElementById(chkid).focus();
	document.getElementById(chkid).value = "<%=defaultrefundCode%>";
}
function updateRefundKey(comp, chkid){
	if(document.getElementById(chkid).value.length < 20){
		alert('20자 이내로 입력하세요');
		document.getElementById(chkid).value = "<%=defaultrefundCode%>";
		document.getElementById(chkid).focus();
		return false;
	}
	document.frmSvArr.target = "xLink";
	document.frmSvArr.cmdparam.value = "U";
	document.frmSvArr.midx.value = comp;
	document.frmSvArr.refundkey.value = document.getElementById(chkid).value;
	document.frmSvArr.action = "/admin/offshop/payment/taxRefund_process.asp"
	document.frmSvArr.submit();
}
function delRefundKey(comp){
	if(confirm("삭제 하시겠습니까?")){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "D";
		document.frmSvArr.midx.value = comp;
		document.frmSvArr.action = "/admin/offshop/payment/taxRefund_process.asp"
		document.frmSvArr.submit();
	}
}
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="A">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 :
				<% drawmaechuldatefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% drawSelectBoxOffShopAll "shopid",shopid %>
					<% end if %>
				<% else %>
					* 매장 : <% drawSelectBoxOffShopAll "shopid",shopid %>
				<% end if %>
				<!--
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	            -->
	            &nbsp;&nbsp;
	            * 정산월 : <input type="text" name="jyyyymm" value="<%=jyyyymm%>" size="7" maxlength="7"> (YYYY-MM)
			</td>
		</tr>
	    </table>
    </td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
    <td>
    * 검색구분 :
    <select name="schType" class="select">
	    <option value="">전체
	    <option value="0" <%= Chkiif(schType="0","selected","") %> >taxRefund 입력내역
	    <option value="1" <%= Chkiif(schType="1","selected","") %>>taxRefund 미입력내역
	    <option value="2" <%= Chkiif(schType="2","selected","") %>>외국인구매내역
    </select>

    &nbsp;&nbsp;
    * 결제금액 :
    <input type="text" name="scgRealsum" size="10" maxlength="10" value="<%= scgRealsum %>">

    &nbsp;&nbsp;
    * TaxRefund일련번호 :
    <input type="text" name="Searchtaxrefundkey" size="25" maxlength="20" value="<%=Searchtaxrefundkey%>"> (20자리)
    </td>
</tr>

</form>
</table>
<!-- 표 상단바 끝-->
<Br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oTaxRefund.FTotalCount %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>주문번호</td>
	<td>결제액</td>
	<td>부가세</td>
	<!--
	<td>카드</td>
	<td>현금</td>
	<td>마일리지</td>
	<td>상품권</td>
	<td>기프트카드</td>
	-->
	<td>구매일</td>
	<td>외국인여부</td>
	<td>정산월</td>
	<td>TaxRefund코드</td>
	<td>비고</td>
</tr>
<%
if oTaxRefund.FResultCount > 0 then
for i=0 to oTaxRefund.FResultCount -1
totcnt = totcnt +1
totrealsum=totrealsum+oTaxRefund.FItemList(i).Frealsum
totVatsum=totVatsum+CLNG(FIX(oTaxRefund.FItemList(i).Frealsum/11))
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><%= oTaxRefund.FItemList(i).ForderNo %></td>
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).Frealsum,0) %></td>
	<td align="right"><%= FormatNumber(FIX(oTaxRefund.FItemList(i).Frealsum/11),0) %></td>
	<!--
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).Fcardsum,0) %></td>
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).Fcashsum,0) %></td>
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).Fspendmile,0) %></td>
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).FGiftCardPaySum,0) %></td>
	<td align="right"><%= FormatNumber(oTaxRefund.FItemList(i).FTenGiftCardPaySum,0) %></td>
	-->
	<td><%= oTaxRefund.FItemList(i).Fshopregdate %></td>
	<td>
	<%
		Select Case oTaxRefund.FItemList(i).Fbuyergubun
			Case "100"	response.write "내국인"
			Case "200"	response.write "외국인"
			Case Else	response.write "미체크"
		End Select
	%>
	</td>
	<td><%= oTaxRefund.FItemList(i).FrefundMonth %></td>
	<td>
		<input type="text" id="taxrefundkey<%=i%>" name="taxrefundkey" maxlength="20" size="25" value="<%= oTaxRefund.FItemList(i).Ftaxrefundkey %>" disabled >
	</td>

	<td>
	<% if isNULL(oTaxRefund.FItemList(i).Ftaxrefundkey) then %>
	<input type="button" class="button" id="btnI<%=i%>" value="입력" style="display:block;" onClick="addRefundKey('<%= oTaxRefund.FItemList(i).Fidx %>','taxrefundkey<%= i%>','btnI<%=i%>','btnS<%=i%>')">
	<input type="button" class="button" id="btnS<%=i%>" value="저장" style="display:none;" onClick="updateRefundKey('<%= oTaxRefund.FItemList(i).Fidx %>', 'taxrefundkey<%=i%>')">
	<% else %>
	<input type="button" class="button" value="삭제" onClick="delRefundKey('<%= oTaxRefund.FItemList(i).Fidx %>')">
	<% end if %>
	</td>
</tr>
<%
next
%>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td >합계</td>
	<td align="right"><%= FormatNumber(totrealsum,0) %></td>
	<td align="right"><%= FormatNumber(totVatsum,0) %></td>
	<!--
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	-->
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
</tr>
<tr height="20">
    <td colspan="16" align="center" bgcolor="#FFFFFF">
        <% if oTaxRefund.HasPreScroll then %>
		<a href="javascript:goPage('<%= oTaxRefund.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oTaxRefund.StartScrollPage to oTaxRefund.FScrollCount + oTaxRefund.StartScrollPage - 1 %>
    		<% if i>oTaxRefund.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oTaxRefund.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="15">검색 결과가 없습니다.</td>
</tr>
<%
end if
%>
</table>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="midx" value="">
<input type="hidden" name="refundkey" value="">
<input type="hidden" name="refundmonth" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->