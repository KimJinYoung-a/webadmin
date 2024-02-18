<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_refundcheckcls.asp"-->
<%

''http://webadmin.10x10.co.kr/cscenter/refund/refund_check_list.asp?page=4&research=on&menupos=1&divcd=A004&yyyy1=2015&mm1=08&dd1=01&yyyy2=2015&mm2=08&dd2=31&returnmethod=&orderserial=&refundMin=&refundMax=&chkGubun=retbea

dim research, page, i
dim divcd, returnmethod, orderserial, chkGubun, refundMin, refundMax
dim yyyy1, yyyy2, mm1, mm2, dd1, dd2
dim fromDate, toDate
dim exCheckFinish
dim returnmethodIN, retR007, retR910, retR900, dategbn

'===============================================================================
research 		= requestCheckVar(request("research"),32)
page 			= requestCheckVar(request("page"),32)
divcd 			= requestCheckVar(request("divcd"),32)
returnmethod 	= requestCheckVar(request("returnmethod"),32)
orderserial 	= requestCheckVar(request("orderserial"),32)
chkGubun 		= requestCheckVar(request("chkGubun"),32)
refundMin 		= requestCheckVar(request("refundMin"),32)
refundMax 		= requestCheckVar(request("refundMax"),32)
exCheckFinish 	= requestCheckVar(request("exCheckFinish"),32)
retR007 		= requestCheckVar(request("retR007"),32)
retR910 		= requestCheckVar(request("retR910"),32)
retR900 		= requestCheckVar(request("retR900"),32)
dategbn     = requestCheckvar(request("dategbn"),32)
'===============================================================================
yyyy1   = request("yyyy1")
yyyy2   = request("yyyy2")
mm1     = request("mm1")
mm2     = request("mm2")
dd1     = request("dd1")
dd2     = request("dd2")

if (yyyy1="") then
	fromDate = CStr(DateSerial(Year(Now()), (Month(Now()) - 1), 1))
	toDate = CStr(DateSerial(Year(Now()), Month(Now()), 0))

    yyyy1 = CStr(Year(fromDate))
    mm1 = CStr(Month(fromDate))
    dd1 =  CStr(day(fromDate))

    yyyy2 = CStr(Year(toDate))
    mm2 = CStr(Month(toDate))
    dd2 =  CStr(day(toDate))
end if

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if dategbn="" then dategbn="finishdate"

if (retR007 <> "") or (retR910 <> "") or (retR900 <> "") then
	returnmethodIN = "'XXXX'"
	if (retR007 <> "") then
		returnmethodIN = returnmethodIN + ",'R007'"
	end if
	if (retR910 <> "") then
		returnmethodIN = returnmethodIN + ",'R910'"
	end if
	if (retR900 <> "") then
		returnmethodIN = returnmethodIN + ",'R900'"
	end if
end if


'===============================================================================
if (page="") then page = 1
if (research="") then
	''divcd = "A003"
	chkGubun = "err"
	''exCheckFinish = "Y"
end if


'===============================================================================
dim oCCSRefundCheck

set oCCSRefundCheck = new CCSRefundCheck


oCCSRefundCheck.FPageSize = 50
oCCSRefundCheck.FCurrPage = page

oCCSRefundCheck.FRectOrderSerial = orderserial
oCCSRefundCheck.FRectDivCD = divcd
oCCSRefundCheck.FRectReturnMethod = returnmethod
oCCSRefundCheck.FRectStartDate = fromDate
oCCSRefundCheck.FRectEndDate = toDate
oCCSRefundCheck.FRectChkGubun = chkGubun
oCCSRefundCheck.FRectRefundMin = refundMin
oCCSRefundCheck.FRectRefundMax = refundMax

oCCSRefundCheck.FRectExCheckFinish = exCheckFinish
oCCSRefundCheck.FRectReturnMethodIN = returnmethodIN
oCCSRefundCheck.FRectDategbn = dategbn
oCCSRefundCheck.GetRefundCheckList

%>

<script language='javascript'>

function jsSetTitle(divcd) {
	var asidList, asidElements, ele, chkFound;
	var frm = document.frmAct;

	chkFound = false;
	asidList = "-1";
	asidElements = document.getElementsByName("asid");

	for (var i = 0; i < asidElements.length; i++) {
		ele = asidElements[i];
		if (ele.checked == true) {
			chkFound = true;
			asidList = asidList + "," + ele.value;
		}
	}

	if (chkFound != true) {
		alert("선택된 내역이 없습니다.");
		return;
	}

	if (confirm("저장하시겠습니까?") == true) {
		if (divcd == "J") {
			// 제휴몰 구매확정 후 환불
			frm.mode.value = "ipjumRefund";
			frm.asidList.value = asidList;
			frm.submit();
		} else if (divcd == "B") {
			// 고객입금 차액환불
			frm.mode.value = "ipjumDiffRefund";
			frm.asidList.value = asidList;
			frm.submit();
		} else if (divcd == "P") {
			// 상품대금 차액환불
			frm.mode.value = "prdDiffRefund";
			frm.asidList.value = asidList;
			frm.submit();
		} else if (divcd == "CB") {
			// CS서비스 - 무통장 환불(배송비)
			frm.mode.value = "csDelivRefund";
			frm.asidList.value = asidList;
			frm.submit();
		} else if (divcd == "U") {
			// 업체정산 및 고객환불
			frm.mode.value = "upcheJungsanRefund";
			frm.asidList.value = asidList;
			frm.submit();
		} else {
			alert("에러.");
			return;
		}
	}
}

function popXL() {
	var popwin = window.open("refund_check_xl_download.asp?page=1&research=on&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>&returnmethod=<%= returnmethod %>&chkGubun=<%= chkGubun %>&dategbn=<%= dategbn %>", "reActAccMonthSummary","width=1000,height=1000 scrollbars=yes resizable=yes");
	popwin.focus();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			구분 :
			<select class="select" name="divcd">
				<option value=""></option>
				<option>------</option>
				<option value="A003" <% if (divcd = "A003") then %>selected<% end if %> >환불</option>
				<option value="A007" <% if (divcd = "A007") then %>selected<% end if %> >카드취소</option>
				<option value="A008" <% if (divcd = "A008") then %>selected<% end if %> >주문취소</option>
				<option>------</option>
				<option value="A004" <% if (divcd = "A004") then %>selected<% end if %> >반품(업배)</option>
				<option value="A010" <% if (divcd = "A010") then %>selected<% end if %> >반품(텐배)</option>
			</select>
			&nbsp;
			기간 :
			<select class="select" name="dategbn">
				<option value="regdate" <%=CHKIIF(dategbn="regdate","selected","")%> >접수일</option>
				<option value="finishdate" <%=CHKIIF(dategbn="finishdate","selected","")%> >완료일</option>
			</select>
            <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			<select class="select" name="returnmethod">
				<option></option>
				<option>------</option>
				<!--
				<option value="R100" <% if (returnmethod = "R100") then %>selected<% end if %> >신용카드 취소</option>
				<option value="R120" <% if (returnmethod = "R120") then %>selected<% end if %> >신용카드 부분취소</option>
				<option value="R400" <% if (returnmethod = "R400") then %>selected<% end if %> >휴대폰결제 취소</option>
				<option value="R020" <% if (returnmethod = "R020") then %>selected<% end if %> >실시간이체 취소</option>
				<option>------</option>
				<option value="R050" <% if (returnmethod = "R050") then %>selected<% end if %> >입점몰결제 취소</option>
				<option>------</option>
				-->
				<option value="R007" <% if (returnmethod = "R007") then %>selected<% end if %> >무통장 환불</option>
				<option value="R910" <% if (returnmethod = "R910") then %>selected<% end if %> >예치금 환불</option>
				<option value="R900" <% if (returnmethod = "R900") then %>selected<% end if %> >마일리지 환급</option>
				<option value="REXC" <% if (returnmethod = "REXC") then %>selected<% end if %> >무통장/예치금/마일리지 이외 환불</option>
				<!--
				<option>------</option>
				<option value="R000" <% if (returnmethod = "R000") then %>selected<% end if %> >환불 없음</option>
				-->
			</select>
			&nbsp;
			주문번호 :
			<input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="14">
			&nbsp;
			환불액 :
			<input type="text" class="text" name="refundMin" value="<%= refundMin %>" size="10">
			~
			<input type="text" class="text" name="refundMax" value="<%= refundMax %>" size="10">
		</td>
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			환불방식 :
			<input type="checkbox" name="retR007" value="Y" <% if (retR007 = "Y") then %>checked<% end if %> > 무통장 환불
			<input type="checkbox" name="retR910" value="Y" <% if (retR910 = "Y") then %>checked<% end if %> > 예치금 환불
			<input type="checkbox" name="retR900" value="Y" <% if (retR900 = "Y") then %>checked<% end if %> > 마일리지 환급
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			검토 :
			<input type="radio" name="chkGubun" value="" <% if (chkGubun = "") then %>checked<% end if %> > 전체
			<input type="radio" name="chkGubun" value="addjung" <% if (chkGubun = "addjung") then %>checked<% end if %> > 업체추가정산(반품 등)
			<input type="radio" name="chkGubun" value="err" <% if (chkGubun = "err") then %>checked<% end if %> > 금액불일치(환불)
			<input type="radio" name="chkGubun" value="ret" <% if (chkGubun = "ret") then %>checked<% end if %> > 반품(업체추가정산)
			<input type="radio" name="chkGubun" value="etc" <% if (chkGubun = "etc") then %>checked<% end if %> > 업체기타정산
			<input type="radio" name="chkGubun" value="retbea" <% if (chkGubun = "retbea") then %>checked<% end if %> disabled> 배송비(변심반품-업배)
			<input type="radio" name="chkGubun" value="retbeaTen" <% if (chkGubun = "retbeaTen") then %>checked<% end if %> disabled> 배송비(변심반품-텐배)
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="checkbox" name="exCheckFinish" value="Y" <% if (exCheckFinish = "Y") then %>checked<% end if %> > 검토완료 내역 제외(예치금환급, 구매확정, 초과입금, 업체정산환불, CS서비스)
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

	<div align="right">
		<input type="button" class="button" value="배송비(CS)" onClick="jsSetTitle('CB');">
		&nbsp;
		<input type="button" class="button" value="초과입금 차액" onClick="jsSetTitle('B');">
		<!--
		<input type="button" class="button" value="상품대금 차액" onClick="jsSetTitle('P');" disabled>
		-->
		&nbsp;
		<input type="button" class="button" value="업체정산환불" onClick="jsSetTitle('U');">
		&nbsp;
		<input type="button" class="button" value="구매확정" onClick="jsSetTitle('J');">
		&nbsp;
		<input type="button" class="button" value="엑셀받기" onclick="popXL();">
	</div>

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmList" method="post" onSubmit="return false;">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><% = oCCSRefundCheck.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= oCCSRefundCheck.FTotalpage %></b>
			&nbsp;
			<font color="red">환불액 합계</font> : <b><%= FormatNumber(oCCSRefundCheck.FrefundSUM,0) %> 원</b>
			&nbsp;
			<font color="red">업체추가정산 합계</font> : <b><%= FormatNumber(oCCSRefundCheck.FaddjungSUM,0) %> 원</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"></td>
		<td width="70">ASID</td>
		<td width="100">주문번호</td>
		<td width="120">구분</td>
		<td width="80">사유01</td>
		<td width="80">사유02</td>
		<td width="220">제목</td>
		<td width="80">환불방식</td>
		<td width="80">취소/반품</td>
		<td width="70"><b>환불액</b></td>
		<!--
		<td width="70">반품배송비</td>
		-->
		<td width="70">업체정산</td>
		<td width="100">정산사유</td>
		<td width="70">관련입금</td>
		<td width="80">접수일</td>
		<td width="80">완료일</td>
		<td>비고</td>
	</tr>
<% if oCCSRefundCheck.FresultCount < 1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for i = 0 to oCCSRefundCheck.FResultCount - 1 %>
	<tr class="a" align="center" bgcolor="FFFFFF">
		<td><input type="checkbox" name="asid" value="<%= oCCSRefundCheck.FItemList(i).Fasid %>"></td>
		<td><a href="javascript:Cscenter_Action_List('<%= oCCSRefundCheck.FItemList(i).FOrderserial %>','','')"><%= oCCSRefundCheck.FItemList(i).Fasid %></a></td>
		<td><a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= oCCSRefundCheck.FItemList(i).FOrderserial %>')"><%= oCCSRefundCheck.FItemList(i).FOrderserial %></a></td>
		<td><%= DDotFormat(oCCSRefundCheck.FItemList(i).Fdivcdname,7) %></td>
		<td align="left" style="padding-left:5px;"><acronym title="<%= oCCSRefundCheck.FItemList(i).Fgubun01name %>"><%= DDotFormat(oCCSRefundCheck.FItemList(i).Fgubun01name,5) %></acronym></td>
		<td align="left" style="padding-left:5px;"><acronym title="<%= oCCSRefundCheck.FItemList(i).Fgubun02name %>"><%= DDotFormat(oCCSRefundCheck.FItemList(i).Fgubun02name,5) %></acronym></td>
		<td align="left" style="padding-left:5px;"><acronym title="<%= oCCSRefundCheck.FItemList(i).Ftitle %>"><%= DDotFormat(oCCSRefundCheck.FItemList(i).Ftitle,18) %></acronym></td>
		<td align="left" style="padding-left:5px;"><acronym title="<%= oCCSRefundCheck.FItemList(i).FreturnmethodName %>"><%= DDotFormat(oCCSRefundCheck.FItemList(i).FreturnmethodName,4) %></acronym></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(oCCSRefundCheck.FItemList(i).FOrgRefundRequire, 0) %></td>
		<td align="right" style="padding-right:5px;"><b><%= FormatNumber(oCCSRefundCheck.FItemList(i).Frefundresult, 0) %></b></td>
		<!--
		<td align="right" style="padding-right:5px;"><%= FormatNumber(oCCSRefundCheck.FItemList(i).Freturndeliverpay, 0) %></td>
		-->
		<td align="right" style="padding-right:5px;"><%= FormatNumber(oCCSRefundCheck.FItemList(i).Fadd_upchejungsandeliverypay, 0) %></td>
		<td align="left" style="padding-left:5px;"><acronym title="<%= oCCSRefundCheck.FItemList(i).Fadd_upchejungsancause %>"><%= DDotFormat(oCCSRefundCheck.FItemList(i).Fadd_upchejungsancause,5) %></acronym></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(oCCSRefundCheck.FItemList(i).FappPrice, 0) %></td>
		<td><acronym title="<%= oCCSRefundCheck.FItemList(i).Fregdate %>"><%= Left(oCCSRefundCheck.FItemList(i).Fregdate,10) %></acronym></td>
		<td><acronym title="<%= oCCSRefundCheck.FItemList(i).Ffinishdate %>"><%= Left(oCCSRefundCheck.FItemList(i).Ffinishdate,10) %></acronym></td>
		<td>
			<% if (oCCSRefundCheck.FItemList(i).Frefundresult <> oCCSRefundCheck.FItemList(i).FOrgRefundRequire) and (oCCSRefundCheck.FItemList(i).FOrgRefundRequire <> 0) then %>
			<font color="red">환불액 불일치</font>
			<% end if %>
		</td>
	</tr>
	<% next %>
<% end if %>
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
    		<% if oCCSRefundCheck.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= oCCSRefundCheck.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>
    		<% for i = 0 + oCCSRefundCheck.StartScrollPage to oCCSRefundCheck.FScrollCount + oCCSRefundCheck.StartScrollPage - 1 %>
    			<% if i > oCCSRefundCheck.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oCCSRefundCheck.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>

<%
set oCCSRefundCheck = Nothing
%>

<form name="frmAct" method="post" onSubmit="return false;" action="refund_check_list_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="asidList" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
