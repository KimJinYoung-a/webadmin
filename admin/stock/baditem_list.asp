<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  불량오차상품등록
' History : 서동석 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->

<%

dim gubun,makerid,itemgubun,itemid,itemoption, disp
gubun       = requestCheckVar(request("gubun"),9)
makerid     = requestCheckVar(request("makerid"),32)
itemgubun   = requestCheckVar(request("itemgubun"),2)
itemid      = requestCheckVar(request("itemid"),9)
itemoption  = requestCheckVar(request("itemoption"),4)
disp        = requestCheckVar(request("disp"),9)

if gubun="" then gubun="ST"
if disp="" then disp="B"

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromdate,todate

fromdate    = requestCheckVar(request("fromdate"),10)
todate      = requestCheckVar(request("todate"),10)

if fromdate<>"" then
	yyyy1 = Left(fromdate,4)
	mm1 = Mid(fromdate,6,2)
	dd1 = Mid(fromdate,9,2)
else
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
end if

if todate<>"" then
	yyyy2 = Left(todate,4)
	mm2 = Mid(todate,6,2)
	dd2 = Mid(todate,9,2)
else
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
end if



if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromdate = CStr(DateSerial(yyyy1, mm1, dd1))
todate = CStr(DateSerial(yyyy2, mm2, dd2+1))


dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectStartDate = fromdate
osummarystock.FRectEndDate	 = todate
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID =  itemid
osummarystock.FRectItemOption =  itemoption
osummarystock.FRectmakerid = makerid
osummarystock.FRectKindDisplay = disp

osummarystock.GetDailyErrItemList

dim i, totrealerritemno, totbaditemno

totrealerritemno=0
totbaditemno =0

dim itembarcode, makername, itemname, itemoptionname

%>

<script language="JavaScript" src="/js/ttpbarcode.js"></script>
<script language='javascript'>
function PopBadItemInput(){
	var popwin = window.open('/common/pop_baditem_input.asp','pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function PopBadItemReInput(){
    var makerid = frm.makerid.value;
	var popwin = window.open('/common/pop_baditem_re_input.asp?makerid=' + makerid,'pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function PopBadItemLossInput(){
    var makerid = frm.makerid.value;
	var popwin = window.open('/common/pop_baditem_re_input.asp?makerid=' + makerid + '&actType=actloss','pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}


function DelDetail(yyyymmdd,itemgubun,itemid,itemoption){
    if (confirm('등록된 내역을 삭제 하시겠습니까?')){
        var mode = "deldetail";
        var popwin = window.open('baditem_process.asp?mode=' + mode + '&yyyymmdd=' + yyyymmdd + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption ,'baditem_process','width=100,height=100,resizable=yes,scrollbars=yes');
        popwin.focus();
    }
}

function RefreshItem(itemgubun,itemid,itemoption){
    if (confirm('새로 고침 하시겠습니까?')){
        var mode = "refreshdetail";
        var popwin = window.open('baditem_process.asp?mode=' + mode + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption ,'baditem_process','width=100,height=100,resizable=yes,scrollbars=yes');
        popwin.focus();
    }
}


function DrawReceiptPrintobj_TEC(elementid,printname){
        var objstring = "";
        var e;
        objstring = '<OBJECT name="' + elementid + '" ';
        objstring = objstring + ' classid="clsid:E76C9051-A8C4-458E-9F60-3C14DB9EECF9" ';
        objstring = objstring + ' codebase="http://billyman/Tec_dol.cab#version=1,5,0,0" ';
        objstring = objstring + ' width=0 ';
        objstring = objstring + ' height=0 ';
        objstring = objstring + ' align=center ';
        objstring = objstring + ' hspace=0 ';
        objstring = objstring + ' vspace=0 ';
        objstring = objstring + ' > ';
        objstring = objstring + ' <PARAM Name="PrinterName" Value="' + printname + '"> ';
        objstring = objstring + ' </OBJECT>';

        document.write(objstring);
}


function Baditemprint(iyyyymmdd, ibarcode, ibarcodeText, imakerid, iitemname, iitemoptionname){
	var X = 1;
	var Y = 1;
	var F = 1;

	// 사용안함
	return;

	// TEC_DO3 : 452
	if (TEC_DO3.IsDriver == 1){
           X = 1.05;
           Y = 1.05;
           F = 1.2;

			TEC_DO3.SetPaper(900,600);
			TEC_DO3.OffsetX = 20;
			TEC_DO3.OffsetY = 20;
			TEC_DO3.PrinterOpen();

			TEC_DO3.PrintText(50*X, 40*Y, "HY견고딕", 30*F, 0, 0, "등 록 일");
			TEC_DO3.PrintText(250*X, 40*Y, "HY견고딕", 30*F, 0, 0, iyyyymmdd);

			TEC_DO3.PrintText(50*X, 80*Y, "HY견고딕", 30*F, 0, 0, "브랜드ID");
			TEC_DO3.PrintText(250*X, 80*Y, "HY견고딕", 30*F, 0, 0, imakerid);

			TEC_DO3.PrintText(50*X, 120*Y, "HY견고딕", 30*F, 0, 0, "상 품 명");
			TEC_DO3.PrintText(250*X, 120*Y, "HY견고딕", 30*F, 0, 0, iitemname);
			TEC_DO3.PrintText(50*X, 160*Y, "HY견고딕", 30*F, 0, 0, "옵 션 명");
			TEC_DO3.PrintText(250*X, 160*Y, "HY견고딕", 30*F, 0, 0, iitemoptionname);

			TEC_DO3.PrintText(50*X, 200*Y, "HY견고딕", 30*F, 0, 0, "상품코드");
			TEC_DO3.PrintText(250*X, 200*Y, "HY견고딕", 50*F, 0, 0, ibarcodeText);


			TEC_DO3.PrintText(270*X, 260*Y, "TEC-BarFont Code128", 80, 0, 0, ibarcode);

			TEC_DO3.PrintText(50*X, 380*Y, "HY견고딕", 30*F, 0, 0, "불량사유");

			TEC_DO3.PrinterClose();


    }else window.status = "TEC B-452 드라이버를 설치해 주세요"
}

// DrawReceiptPrintobj_TEC("TEC_DO3","TEC B-452");

function IndexBarcodePrint(barcode, makerid, itemname, itemoptionname, customerprice, printno) {
	// /js/barcode.js 참조
	if (initTTPprinter("TTP-243_80x50", "T", "N", "                         www.10x10.co.kr                         ", "Y", "￦", "Y", 3, 0) != true) {
		alert('프린터가 설치되지 않았거나 올바른 프린터명이 아닙니다.[4]');
		return;
	}

	printTTPOneIndexBarcodeForBadItem(barcode, makerid, itemname, itemoptionname, customerprice, 1);
}

function isUInt(val) {
	var re = /^[0-9]+$/;
	return re.test(val);
}

function trimString(val) {
    return val.replace(/^\s+|\s+$/gm,'');
}

function SubmitFrm(frm) {
	frm.itemid.value = trimString(frm.itemid.value);

	if (frm.itemid.value != "") {
		if (isUInt(frm.itemid.value) != true) {
			alert("상품코드는 숫자만 가능합니다.");
			return;
		}
	}

	frm.submit();
}

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=frm method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
        	브랜드명 : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
			&nbsp;
			상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="9" size="9">
        	&nbsp;
        	검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="SubmitFrm(document.frm);">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			표시:
			<input type="radio" name="disp" value="A" <% if (disp = "A") then %>checked<% end if %>> 전체
        	<input type="radio" name="disp" value="B" <% if (disp = "B") then %>checked<% end if %>> 불량
        	<input type="radio" name="disp" value="D" <% if (disp = "D") then %>checked<% end if %>> 오차
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="불량상품등록" onclick="PopBadItemInput()" border="0" >&nbsp;&nbsp;
			<!--
        	<input type="button" class="button" value="불량상품반품" onclick="PopBadItemReInput()" border="0">&nbsp;&nbsp;
        	<input type="button" class="button" value="불량상품로스처리" onclick="PopBadItemLossInput()" border="0">
			-->
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="65">등록일</td>
		<td width="100">브랜드ID</td>
		<td width="35">거래<br>구분</td>
		<td width="35">상품<br>구분</td>
		<td width="70">상품<br>코드</td>
		<td width="40">옵션</td>
		<td>아이템명</td>
		<td>옵션</td>
		<td width="50">소비자가</td>
		<td width="30">불량</td>
		<td width="30">오차</td>
		<td width="80">등록자</td>
		<!--
		<td width="50">삭제</td>
		-->
		<td width="50">출력</td>
    </tr>
	<% for i=0 to osummarystock.FResultCount - 1 %>
	<%

    itembarcode 	= osummarystock.FItemList(i).Fitemgubun & BF_GetFormattedItemId(osummarystock.FItemList(i).FItemId) & osummarystock.FItemList(i).FItemOption

	makername		= Replace(osummarystock.FItemList(i).FMakerid, Chr(34), "")
    itemname		= Replace(osummarystock.FItemList(i).Fitemname, Chr(34), "")
    itemoptionname	= Replace(osummarystock.FItemList(i).Fitemoptionname, Chr(34), "")

    makername		= Replace(makername, "'", "")
    itemname		= Replace(itemname, "'", "")
    itemoptionname	= Replace(itemoptionname, "'", "")

	totrealerritemno = totrealerritemno + osummarystock.FItemList(i).Ferrbaditemno
	totbaditemno	 = totbaditemno + osummarystock.FItemList(i).Ferrrealcheckno
	%>
    <tr align="center" bgcolor="#FFFFFF">
		<td><%= osummarystock.FItemList(i).Fyyyymmdd %></td>
		<td><%= osummarystock.FItemList(i).Fmakerid %></td>
		<td><%= osummarystock.FItemList(i).GetMwDivName %></td>
		<td><%= osummarystock.FItemList(i).FItemgubun %></td>
		<td><%= CHKIIF(osummarystock.FItemList(i).FItemid>=1000000,format00(8,osummarystock.FItemList(i).FItemid),format00(6,osummarystock.FItemList(i).FItemid)) %></td>
		<td align="center"><%= osummarystock.FItemList(i).FItemoption %></td>
		<td align="left"><%= osummarystock.FItemList(i).FItemname %></td>
		<td><%= osummarystock.FItemList(i).FItemOptionName %></td>
		<td align="right"><%= formatnumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
		<td><%= osummarystock.FItemList(i).Ferrbaditemno %></td>
		<td><%= osummarystock.FItemList(i).Ferrrealcheckno %></td>
		<td>
			<%= osummarystock.FItemList(i).Freguser %>
			<% if Not IsNull(osummarystock.FItemList(i).Fmodiuser) and (osummarystock.FItemList(i).Fmodiuser <> osummarystock.FItemList(i).Freguser) then %>
			<br>-&gt; <%= osummarystock.FItemList(i).Fmodiuser %>
			<% end if %>
		</td>
		<!--
		<td><a href="javascript:DelDetail('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).FItemgubun %>','<%= osummarystock.FItemList(i).FItemid %>','<%= osummarystock.FItemList(i).FItemoption %>');"><img src="/images/icon_delete.gif" width="45" border="0"></a></td>
		-->
    	<td>
			<!--
			<input type="button" class="button" value="출력" onclick="Baditemprint('<%= osummarystock.FItemList(i).Fyyyymmdd %>', '<%= osummarystock.FItemList(i).FItemgubun & CHKIIF(osummarystock.FItemList(i).FItemid>=1000000,Format00(8,osummarystock.FItemList(i).FItemid),Format00(6,osummarystock.FItemList(i).FItemid)) & osummarystock.FItemList(i).FItemoption %>', '<%= osummarystock.FItemList(i).FItemgubun & "-" & CHKIIF(osummarystock.FItemList(i).FItemid>=1000000,Format00(8,osummarystock.FItemList(i).FItemid),Format00(6,osummarystock.FItemList(i).FItemid)) & "-" & osummarystock.FItemList(i).FItemoption %>', '<%= osummarystock.FItemList(i).FMakerid %>', '<%= Replace(Replace(osummarystock.FItemList(i).Fitemname,"'",""),Chr(34),"") %>', '<%= Replace(Replace(osummarystock.FItemList(i).Fitemoptionname,"'",""),Chr(34),"") %>')">
			-->
			<input type="button" class="button" value="출력" onClick="IndexBarcodePrint('<%= itembarcode %>', '<%= osummarystock.FItemList(i).Fmakerid %>', '<%= itemname %>', '<%= itemoptionname %>', '<%= osummarystock.FItemList(i).Fyyyymmdd %>', 1)">
		</td>
    </tr>
   	<% next %>
	<tr align="center" bgcolor="#FFFFFF">
	  <td>Total</td>
	  <td colspan="8"></td>
	  <td><%= totrealerritemno %></td>
	  <td><%= totbaditemno %></td>
	  <td></td>
	  <td></td>
	</tr>

</table>


<%
set osummarystock = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
