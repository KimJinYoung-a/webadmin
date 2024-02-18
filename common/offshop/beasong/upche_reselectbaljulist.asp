<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.26 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim i , ojumun ,ix,sql ,iSall ,detailidxarr ,j ,TooManyOrder
	detailidxarr =  request("chkidx")
	iSall   =  request("isall")

set ojumun = new cupchebeasong_list
	ojumun.FRectdetailidxarr = detailidxarr
	ojumun.FRectIsAll       = iSall
	ojumun.FRectDesignerID  = session("ssBctID")
	ojumun.fReDesignerSelectBaljuList()

TooManyOrder = FALSE

if ojumun.FResultCount>2000 then
    TooManyOrder=true
end if
%>

<SCRIPT LANGUAGE="JavaScript">

function winPrint() {
window.print();
}

function ExcelPrint(iSheetType) {
	xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="/common/offshop/beasong/upche_dobeasonglistexcel.asp";
	xlfrm.submit();

}

function CsvPrint(iSheetType){
    xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="/common/offshop/beasong/upche_dobeasonglistCSV.asp";
	xlfrm.submit();
}

</SCRIPT>
<STYLE TYPE="text/css">
	.print {page-break-before: always;font-size: 12px; color:red;}
	.no {font-size: 12px; color:red;}
</STYLE>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">액션</td>
	<td align="left">
	    <table border="0" cellspacing="3" cellpadding="3" >
	    <tr>
	        <td><input type="button" class="button" onclick="ExcelPrint('V4')" value="엑셀 저장  (일련번호 추가)"></td>
	        <td><input type="button" class="button" onclick="ExcelPrint('')" value="엑셀파일로 저장(주소분리)"></td>
	        <td><input type="button" class="button" onclick="ExcelPrint('V2')" value="엑셀파일로 저장(주소통합)"></td>
	    </tr>
		<tr>
	        <td><input type="button" class="button" onclick="CsvPrint('')" value="            CSV 저장           "></td>
	        <td><input type="button" class="button" onclick="winPrint()" value="프린트하기"></td>
	        <td></td>
	    </tr>
	    </table>
	</td>
	<td width="100" bgcolor="<%= adminColor("gray") %>">
		총 건수 : <font color="red"><span id="totalno"></span>건</font>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<% IF (TooManyOrder) then %>
	주문 내역이 많아 내역이 표시되지 않습니다.
	<br>
	엑셀 데이터는 다운로드 가능합니다.
<% else %>
<% for ix=0 to ojumun.FResultCount - 1 %>
<table class="no">
<tr><td><% = ix +1 %></td></tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" class="a">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<td height="25">주문번호</td>
	<td>주문일</td>
	<td>수령인</td>
	<td>수령인 전화</td>
	<td>수령인 핸드폰</td>
	<td>수령인 email</td>
</tr>
<tr align="center">
	<td height="25"><%= ojumun.FItemList(ix).forderno %></td>
	<td><%= FormatDateTime(ojumun.FItemList(ix).FRegDate,2) %></td>
	<td><%= ojumun.FItemList(ix).FReqName %></td>
	<td><%= ojumun.FItemList(ix).FReqPhone %></td>
	<td><%= ojumun.FItemList(ix).FReqHp %></td>
	<td>&nbsp;</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="6">수령인 주소</td>
</tr>
<tr align="center">
	<td colspan="6"><%= ojumun.FItemList(ix).FReqZipCode %>&nbsp;<%= ojumun.FItemList(ix).FReqZipAddr %>&nbsp;<%= ojumun.FItemList(ix).FReqAddress %></td>
</tr>
<tr>
	<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">기타사항</td>
	<td colspan="5" align="center">&nbsp;<%= nl2br(db2html(ojumun.FItemList(ix).FComment)) %></td>
</tr>
</table>

<br>

<table width="100%" border="1" cellspacing="0" cellpadding="0" class="a">
<tr align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">상품ID</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td width="70">판매가</td>
	<td width="50">수량</td>
</tr>
<tr align="center">
	<td><%= ojumun.FItemList(ix).Fitemid %></td>
	<td><%= ojumun.FItemList(ix).FItemName %></td>
	<td><%= ojumun.FItemList(ix).FItemoptionName %>&nbsp;</td>
	<td><%= FormatNumber(ojumun.FItemList(ix).Fsellprice,0) %></td>
	<td><%= ojumun.FItemList(ix).FItemNo %></td>
</tr>
</table>

<br>
<% if ((ix+1) mod 4) = 0 then %><div class="print">&nbsp;</div><% end if %>
<% next %>

<% end if %>
<%
set ojumun = Nothing
%>
<iframe name="iiframeXL" name="iiframeXL" width=0 height=0 frameborder=0 scrolling=no align="center"></iframe>
<form name="xlfrm" method="post" action="">
	<input type="hidden" name="detailidxarr" value="<%= detailidxarr %>">
	<input type="hidden" name="isall" value="<%= iSall %>">
	<input type="hidden" name="SheetType" value="">
</form>
<script language='javascript'>
	totalno.innerText = "<%= ix %>";
</script>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->