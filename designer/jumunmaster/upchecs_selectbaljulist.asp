<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_cs_baljucls.asp"-->

<SCRIPT LANGUAGE="JavaScript">
<!--

function winPrint() {
	window.print();
}

//-->
</SCRIPT>
<STYLE TYPE="text/css">
<!--
.print {page-break-before: always;font-size: 12px; color:red;}
.no {font-size: 12px; color:red;}
body {background-color:"#FFFFFF"}
-->
</STYLE>
<%

Dim cutPage :  cutPage =4
IF (session("ssBctID") ="funnyhands") then
    cutPage = 3
End IF
%>
	<input type="hidden" name="menupos" value="<%= menupos %>">

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>">
		<td width="50" bgcolor="<%= adminColor("gray") %>">액션</td>
		<td align="left">
			<input type="button" class="button" onclick="winPrint()" value="프린트하기">
			<!--
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('')" value="엑셀(주소분리)">
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('V2')" value="엑셀(주소통합)">
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('V3')" value="엑셀(업체코드)">
			-->
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('V4')" value="엑셀(CS일련번호 추가)">
			&nbsp;
			<input type=button class="button" onclick="CsvPrint()" value="CSV로 저장">
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">
			총 건수 : <font color="red"><span id="totalno"></span>건</font>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->



<%
dim i
dim ojumun
dim ix,sql
Dim listitemlist,listitem,listitemcount


listitem =  Replace(request("orderserial"), " ", "")  '' orderserial is cs detail idx
set ojumun = new CCSJumunMaster

ojumun.FRectOrderSerial = listitem
ojumun.FRectDesignerID = session("ssBctID")
ojumun.DesignerCS_SelectBaljuList

%>
<script language="JavaScript">
<!--
function ExcelPrint(iSheetType) {
    xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="upchecs_beasonglist_excel.asp";
	xlfrm.submit();
}

function CsvPrint(iSheetType){
    xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="upchecs_beasonglist_csv.asp";
	xlfrm.submit();
}


//OLD function
function ExcelGo1() {
	//var popwin = window.open('','popexcel','width=800, height=600, scrollbars=1,resizable=1');
	//xlfrm.target="popexcel";
	//popwin.location="beasonglistexcel_process.asp?orderserial=<%= listitem %>";


	xlfrm.target="_blank";
	xlfrm.action="beasonglistexcel_process.asp";
	xlfrm.submit();

}

//OLD function
function ExcelGo2() {
	//var popwin = window.open('','popexcel','width=800, height=600, scrollbars=1,resizable=1');
	//xlfrm.target="popexcel";
	//popwin.location="beasonglistexcel_process.asp?orderserial=<%= listitem %>";

	xlfrm.target="_blank";
	xlfrm.action="beasonglistexcel2_process.asp";
	xlfrm.submit();
}
//-->
</script>

<% for ix=0 to ojumun.FResultCount - 1 %>
<table class="no">
<tr>
	<td><% = ix +1 %></td>
</tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" class="a">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="22">CS주문번호</td>
		<td>주문일</td>
		<td>구매자 성명</td>
		<td>구매자 전화</td>
		<td>구매자 핸드폰</td>
		<td>구매자 email</td>
	</tr>
	<tr align="center">
		<td height="22"><%= ojumun.FMasterItemList(ix).FOrgOrderSerial %>-<%= ojumun.FMasterItemList(ix).Fcsmasteridx %></td>
		<td><%= FormatDateTime(ojumun.FMasterItemList(ix).FRegDate,2) %></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyName %></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyPhone %></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyHp %></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyemail %></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="22">수령인</td>
		<td>수령인 전화</td>
		<td>수령인 핸드폰</td>
		<td colspan="3">수령인 주소</td>
	</tr>
	<tr align="center">
		<td height="22"><%= ojumun.FMasterItemList(ix).FReqName %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqPhone %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqHp %></td>
		<td colspan="3"><%= ojumun.FMasterItemList(ix).FReqZipCode %>&nbsp;<%= ojumun.FMasterItemList(ix).FReqZipAddr %>&nbsp;<%= ojumun.FMasterItemList(ix).FReqAddress %></td>
	</tr>
<% if Not IsNULL(ojumun.FMasterItemList(ix).Freqdate) then %>
	<tr>
		<td align="center" height="22">메세지<br>서비스</td>
		<td colspan="5" align="left">
			<table border="0" cellspacing="5" cellpadding="0" class="a">
				<tr>
					<td>배송희망일 : </td>
					<td> <%= Left(CStr(ojumun.FMasterItemList(ix).Freqdate),10) %>일 <%= (ojumun.FMasterItemList(ix).Freqtime) %>시 </td>
				</tr>
				<tr>
					<td>카드/리본 : </td>
					<td> <%= (ojumun.FMasterItemList(ix).getCardribbonName) %></td>
				</tr>
				<tr>
					<td>메세지 : </td>
					<td><%= nl2br(db2html(ojumun.FMasterItemList(ix).Fmessage)) %></td>
				</tr>
				<tr>
					<td>보내는 사람 : </td>
					<td><%= (db2html(ojumun.FMasterItemList(ix).Ffromname)) %></td>
				</tr>
			</table>
		</td>
	</tr>
<% end if %>
	<tr>
		<td align="center" height="22" bgcolor="<%= adminColor("tabletop") %>">접수내용</td>
		<td colspan="5" align="center">&nbsp;<%= nl2br(db2html(ojumun.FMasterItemList(ix).FComment)) %></td>
	</tr>
</table>

<p>

<table width="100%" border="1" cellspacing="0" cellpadding="0" class="a">
	<tr align="center" height="22" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60" height="22">상품ID</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="70">판매가</td>
		<td width="50">수량</td>
	</tr>
	<tr align="center" height="22">
		<td><a href="http://www.10x10.co.kr/street/designershop.asp?itemid=<%= ojumun.FMasterItemList(ix).Fitemid %>" target="_blank"><%= ojumun.FMasterItemList(ix).Fitemid %></a></td>
		<td><%= ojumun.FMasterItemList(ix).FItemName %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemoptionName %></td>
		<td><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0) %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemNo %></td>
	</tr>
	<tr align="center">
		<td>주문제작<br>메세지</td>
		<td colspan="4" align="left">
		<% if (Not IsNULL(ojumun.FMasterItemList(ix).Frequiredetail)) and (ojumun.FMasterItemList(ix).Frequiredetail<>"") then %>
		<% if (ojumun.FMasterItemList(ix).FItemNo>1) then %>
		<% for i=0 to ojumun.FMasterItemList(ix).FItemNo-1 %>
		    [<%= i+ 1 %>번 상품 문구]
		    <%= nl2Br(splitValue(ojumun.FMasterItemList(ix).Frequiredetail,CAddDetailSpliter,i)) %>
		    <br>
		<% next %>
		<% else %>
		<%= nl2Br(ojumun.FMasterItemList(ix).Frequiredetail) %>
		<% end if %>
		<% end if %>
		</td>
	</tr>
</table>

<br>
<% if (((ix+1) mod cutPage) = 0) then %><div class="print">&nbsp;</div><% end if %>
<% next %>
<%
set ojumun = Nothing
%>
<iframe name="iiframeXL" name="iiframeXL" width="0" height="0" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>

<form name=xlfrm method=post action="">
<input type="hidden" name="orderserial" value="<%= listitem %>">
<input type="hidden" name="isall" value="">
<input type="hidden" name="SheetType" value="">
</form>
<script language='javascript'>
	totalno.innerText = "<%= ix %>";
</script>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
