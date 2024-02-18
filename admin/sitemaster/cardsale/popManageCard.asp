<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/etc/mustPriceCls.asp"-->
<!-- #include virtual="/admin/sitemaster/cardsale/cardsaleCls.asp"-->
<%
Dim mode, idx, styleStr
idx         = request("idx")

If idx = "" Then
    mode = "I"
	styleStr = "display:none;"
Else
    mode = "U"
End If

Dim startDate, endDate, startDateTime, endDateTime, cardCode, saleType, minPrice, maxPrice, salePrice, isusing
Dim oCardSale, arrRows, lp, bannerTitle, bannerView, bgcolor, blnWeb, blnMobile, blnApp
Set oCardSale = new CCardSale
	arrRows = oCardSale.fnCardList
If mode = "U" Then
	oCardSale.FRectIdx		= idx
	oCardSale.getCardSaleOneItem

	cardCode		= oCardSale.FOneItem.FCardCode
	saleType		= oCardSale.FOneItem.FSaleType
	minPrice		= oCardSale.FOneItem.FMinPrice
	maxPrice		= oCardSale.FOneItem.FMaxPrice
	salePrice		= oCardSale.FOneItem.FSalePrice
	isusing			= oCardSale.FOneItem.FIsUsing
	startDate  	 	= LEFT(oCardSale.FOneItem.FStartDate, 10)
	endDate     	= LEFT(oCardSale.FOneItem.FEndDate, 10)
	startDateTime 	= Num2Str(hour(oCardSale.FOneItem.FStartDate),2,"0","R") & ":" & Num2Str(minute(oCardSale.FOneItem.FStartDate),2,"0","R") & ":" & Num2Str(Second(oCardSale.FOneItem.FStartDate),2,"0","R")
	endDateTime 	= Num2Str(hour(oCardSale.FOneItem.FEndDate),2,"0","R") & ":" & Num2Str(minute(oCardSale.FOneItem.FEndDate),2,"0","R") & ":" & Num2Str(Second(oCardSale.FOneItem.FEndDate),2,"0","R")
	bannerTitle		= oCardSale.FOneItem.FbannerTitle
	bannerView		= oCardSale.FOneItem.FbannerView
	bgcolor			= oCardSale.FOneItem.Fbgcolor
	blnWeb			= oCardSale.FOneItem.FblnWeb
	blnMobile		= oCardSale.FOneItem.FblnMobile
	blnApp			= oCardSale.FOneItem.FblnApp

	If saleType = "2" Then
		styleStr = ""
	Else
		styleStr = "display:none;"
	End If
End If
Set oCardSale = nothing
%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript'>
function checkDate() {
	var frm = document.frm;
	var startDate = frm.startDate.value;
	var endDate = frm.endDate.value;
	var startdate = toDate(startDate);
	var enddate = toDate(endDate);

	if (startdate > enddate) {
		alert("종료일이 시작일보다 과거날짜입니다.");
		return false;
	}
	return true;
}
function frm_check(){
	if ($("#cardCode").val() == "") {
		alert('카드를 선택 하세요');
		$("#cardCode").focus();
		return false;
	}
	if ($("#termSdt").val() == "") {
		alert('특가 시작일을 입력하세요');
		return false;
	}
	if ($("#termEdt").val() == "") {
		alert('특가 종료일을 입력하세요');
		return false;
	}
	if(!$('input:radio[name=saleType]').is(':checked')){
		alert('할인타입을 선택하세요')
		return false;
	}
	if ($("#salePrice").val() == "") {
		alert('금액을 입력하세요');
		$("#salePrice").focus();
		return false;
	}

	if ($("#minPrice").val() == "") {
		alert('최소구매금액을 입력하세요');
		$("#minPrice").focus();
		return false;
	}
	if ($("#isusing").val() == "") {
		alert('사용여부를 선택 하세요');
		$("#isusing").focus();
		return false;
	}
	if (confirm('저장 하시겠습니까?')){
		document.frm.submit();
	}
}
function numOnly(selector){
    selector.value = selector.value.replace(/[^0-9]/g,'');
}
function chkCpnType(comp){
	if (comp.value=="2"){
		$("#imxcpndiscount_tr").show();
	}else{
		$("#imxcpndiscount_tr").hide();
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="card_process.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">카드명</td>
    <td bgcolor="#FFFFFF">
        <select id="cardCode" name="cardCode" class="select">
			<option value="">-선택-</option>
	<%
		 If isArray(arrRows) Then
		 	For lp = 0 To Ubound(arrRows, 2)
	%>
			<option value="<%= arrRows(0, lp) %>" <%= Chkiif(arrRows(0, lp) = cardCode, "selected", "") %> ><%= arrRows(2, lp) %></option>
	<%
			Next
		End If
	%>
        </select>
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">기간</td>
    <td bgcolor="#FFFFFF">
        <input type="text" id="termSdt" name="startDate" readonly size="11" maxlength="10" value="<%= startDate %>" style="cursor:pointer; text-align:center;" />
        <input type="text" id="termSdtTime" name="startDateTime" size="8" maxlength="8" value="<%= startDateTime %>" style="text-align:center;" /> ~
        <input type="text" id="termEdt" name="endDate" readonly size="11" maxlength="10" value="<%= endDate %>" style="cursor:pointer; text-align:center;" />
        <input type="text" id="termEdtTime" name="endDateTime" size="8" maxlength="8" value="<%= endDateTime %>" style="text-align:center;" />
        <script type="text/javascript">
            var CAL_Start = new Calendar({
                inputField : "termSdt", trigger    : "termSdt",
                onSelect: function() {
                    var date = Calendar.intToDate(this.selection.get());
                    CAL_End.args.min = date;
                    CAL_End.redraw();
                    this.hide();
                    if(frm.startDateTime.value=="") frm.startDateTime.value='00:00:00';
                    if(frm.endDateTime.value=="") frm.endDateTime.value='23:59:59';
                    if(frm.endDate.value==""||getDayInterval(frm.startDate.value, frm.endDate.value) < 0) frm.endDate.value=frm.startDate.value;
                    doInsertDayInterval();	// 날짜 자동계산
                }, bottomBar: true, dateFormat: "%Y-%m-%d"
            });
            var CAL_End = new Calendar({
                inputField : "termEdt", trigger    : "termEdt",
                onSelect: function() {
                    var date = Calendar.intToDate(this.selection.get());
                    CAL_Start.args.max = date;
                    CAL_Start.redraw();
                    this.hide();

                    if(frm.startDate.value==""||getDayInterval(frm.startDate.value, frm.endDate.value) < 0) frm.startDate.value=frm.endDate.value;
                    doInsertDayInterval();	// 날짜 자동계산
                }, bottomBar: true, dateFormat: "%Y-%m-%d"
            });
        </script>
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">할인타입</td>
    <td bgcolor="#FFFFFF">
	<input type=text id="salePrice" name="salePrice" value="<%= salePrice %>" maxlength="7" size="10" onkeyup="numOnly(this)" onblur="numOnly(this)">
		<input type="radio" name="saleType" value="1" <%= Chkiif(saleType = "1", "checked", "") %> onClick="chkCpnType(this)">원할인
	    <input type="radio" name="saleType" value="2" <%= Chkiif(saleType = "2", "checked", "") %> onClick="chkCpnType(this)">%할인
	(금액 또는 % 할인)
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">최소구매금액</td>
    <td bgcolor="#FFFFFF">
        <input type=text id="minPrice" name="minPrice" value="<%= minPrice %>" maxlength=7 size=10  onkeyup="numOnly(this)" onblur="numOnly(this)">원 이상 구매시 사용가능(숫자)
    </td>
</tr>
<tr id="imxcpndiscount_tr" height="25" bgcolor="<%= adminColor("gray") %>" style="<%= styleStr %>">
    <td width="20%">최대할인금액</td>
    <td bgcolor="#FFFFFF">
        <input type=text name="maxPrice" value="<%= maxPrice %>" maxlength=7 size=10  onkeyup="numOnly(this)" onblur="numOnly(this)">원 할인(숫자)(ex 5% 시 10000 / 10%시 20000 / 무제한 0 입력)
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">배너 노출 문구</td>
    <td bgcolor="#FFFFFF">
        <input type=text name="bannerTitle" value="<%= bannerTitle %>" maxlength="64" size="60"><br><input type="checkbox" name="bannerView"<% if bannerView="N" then response.write " checked" %>>배너 노출 안함
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">채널</td>
    <td bgcolor="#FFFFFF">
        <input type="checkbox" class="formCheckInput" name="blnWeb" value="Y"<% if blnWeb="Y" then response.write " checked" %>> PC
		<input type="checkbox" class="formCheckInput" name="blnMobile" value="Y"<% if blnMobile="Y" then response.write " checked" %>> Mobile
		<input type="checkbox" class="formCheckInput" name="blnApp" value="Y"<% if blnApp="Y" then response.write " checked" %>> APP
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">배너 배경 컬러</td>
    <td bgcolor="#FFFFFF">
        #<input type=text name="bgcolor" value="<%= bgcolor %>" maxlength="64" size="10">
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">사용여부</td>
    <td bgcolor="#FFFFFF">
        <select name="isusing" id="isusing" class="select">
			<option value="">-선택-</option>
			<option value="Y" <%= Chkiif(isusing = "Y", "selected", "") %> >Y</option>
			<option value="N" <%= Chkiif(isusing = "N", "selected", "") %> >N</option>
        </select>
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>" align="center">
    <td bgcolor="#FFFFFF" colspan="2">
        <input type="button" value="저장" class="button" onclick="frm_check();" />
    </td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
