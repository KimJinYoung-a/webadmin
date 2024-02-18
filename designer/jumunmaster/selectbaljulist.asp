<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/checkPartnerLog.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_baljucls.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->

<SCRIPT LANGUAGE="JavaScript">
<!--
function winPrint() {
window.print();
}

function CustExcelPrint(iSheetType) {
	xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="dobeasonglistexcelCust.asp";
	xlfrm.submit();
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
function IsInvalidOrderCharExists(s)
        dim buf, result, iid

        iid = 1
        do until iid > len(s)
                buf = mid(s, iid, cint(1))
                if (buf = ",") or (buf = " ") then
                       result = false
                elseif (buf >= "0" and buf <= "9") then
                        result = false
                else
                        IsInvalidOrderCharExists = true
                        exit function
                end if
                iid = iid + 1
        loop

        IsInvalidOrderCharExists = false
end function


Dim isCustomizeBrand

isCustomizeBrand = ((session("ssBctID") ="victoria001") or (session("ssBctID") ="thegirin"))

Dim cutPage :  cutPage =4
IF (session("ssBctID") ="funnyhands") then
    cutPage = 3
End IF

dim i
dim ojumun
dim ix,sql
Dim listitemlist,listitem,listitemcount


listitem =  Replace(request("orderserial"), " ", "")  '' orderserial is Index of Order

''2017/03/21 추가=============================================
if (IsInvalidOrderCharExists(listitem)) then
    response.write "<script>alert('올바르지 않은 문자열이 있습니다.')</script>"
    dbget.Close() : response.end
end if
''============================================================
%>
	<input type="hidden" name="menupos" value="<%= menupos %>">

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>">
		<td width="50" bgcolor="<%= adminColor("gray") %>">액션</td>
		<td align="left">
			<input type="button" class="button" onclick="winPrint()" value="프린트하기">
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('')" value="엑셀(주소분리)">
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('V2')" value="엑셀(주소통합)">
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('V3')" value="엑셀(업체코드)">
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('V4')" value="엑셀(일련번호 추가)">
			&nbsp;
			<input type=button class="button" onclick="CsvPrint()" value="CSV로 저장">
			<% if (isCustomizeBrand) then %>
		    <br><br>
		    <input type=button class="button_ing" onclick="CustExcelPrint('')" value="<%= session("ssBctID") %> 전용포멧 Excel">
		    <% end if %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">
			총 건수 : <font color="red"><span id="totalno"></span>건</font>
		</td>
	</tr>
	<!--
	<tr bgcolor="<%= adminColor("topbar") %>">
		<td colspan="10">
			엑셀파일로 저장(1)은 배송지 주소가 1,2로 나누어져 출력됩니다.<br>
			엑셀파일로 저장(2)은 배송지 주소가 1,2가 하나로 합쳐져서 출력됩니다.<br>
			사용하시는 양식에 맞게 (1) 또는 (2)를 선택하셔서 사용하십시요.
		</td>
	</tr>
	-->
</table>
<!-- 액션 끝 -->



<%

set ojumun = new CJumunMaster

ojumun.FRectOrderSerial = listitem
ojumun.FRectDesignerID = session("ssBctID")
ojumun.DesignerSelectBaljuList

dim dumitime : dumitime = Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)
dim oGift, j
set oGift = new COrderGift
%>
<script language="JavaScript">
<!--
function ExcelPrint(iSheetType) {
    xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="dobeasonglistexcel.asp?dumi=<%=dumitime%>";
	xlfrm.submit();
}

function CsvPrint(iSheetType){
    xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="dobeasonglistCSV.asp?dumi=<%=dumitime%>";
	xlfrm.submit();
}


//OLD function
function ExcelGo1() {
	//var popwin = window.open('','popexcel','width=800, height=600, scrollbars=1,resizable=1');
	//xlfrm.target="popexcel";
	//popwin.location="beasonglistexcel_process.asp?orderserial=<%= listitem %>";


	xlfrm.target="_blank";
	xlfrm.action="beasonglistexcel_process.asp?dumi=<%=dumitime%>";
	xlfrm.submit();

}

//OLD function
function ExcelGo2() {
	//var popwin = window.open('','popexcel','width=800, height=600, scrollbars=1,resizable=1');
	//xlfrm.target="popexcel";
	//popwin.location="beasonglistexcel_process.asp?orderserial=<%= listitem %>";

	xlfrm.target="_blank";
	xlfrm.action="beasonglistexcel2_process.asp?dumi=<%=dumitime%>";
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
		<td height="22">주문번호</td>
		<td>주문일</td>
		<td>구매자 성명</td>
		<td>구매자 전화</td>
		<td>구매자 핸드폰</td>
		<td>구매자 email</td>
	</tr>
	<tr align="center">
		<td height="22"><%= ojumun.FMasterItemList(ix).FOrderSerial %></td>
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
					<td> <%= Left(CStr(ojumun.FMasterItemList(ix).Freqdate),10) %>일 <%= (ojumun.FMasterItemList(ix).GetReqTimeText) %> </td>
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
		<td align="center" height="22" bgcolor="<%= adminColor("tabletop") %>">기타사항</td>
		<td colspan="5" align="center">&nbsp;<%= nl2br(db2html(ojumun.FMasterItemList(ix).FComment)) %></td>
	</tr>
	<%
	oGift.FRectOrderSerial = ojumun.FMasterItemList(ix).FOrderSerial
    oGift.FRectMakerid = session("ssBctId")
    oGift.FRectGiftDelivery = "Y"
    oGift.GetOneOrderGiftlist
	%>
	<% if (oGift.FResultCount>0) then %>
	<tr>
	    <td align="center" height="22">사은품</td>
		<td colspan="5" align="left">
		    <% for j=0 to oGift.FResultCount -1 %>
                <%= oGift.FItemList(j).GetEventConditionStr %><br>
            <% next %>
		</td>
	</tr>
	<% end if %>
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
		<td><a href="http://www.10x10.co.kr/street/designershop.asp?itemid=<%= ojumun.Fitemid %>" target="_blank"><%= ojumun.FMasterItemList(ix).Fitemid %></a></td>
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
		<%= nl2Br(Replace(ojumun.FMasterItemList(ix).Frequiredetail, CAddDetailSpliter, "")) %>
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
set oGift = Nothing
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
