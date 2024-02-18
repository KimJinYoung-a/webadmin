<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/etc/mustPriceCls.asp"-->
<%
Dim mode, idx, isModify, mallid
Dim oMustPrice
mallid      = request("mallid")
isModify    = request("isModify")
idx         = request("idx")

If isModify = "" Then
    mode = "I"
Else
    mode = "U"
End If

Dim mallgubun, itemid, mustPrice, mustMargin, startDate, endDate, startDateTime, endDateTime, orgpricestartDate, orgpricestartDateTime, orgpriceendDate, orgpriceendDateTime
If mode = "U" Then
    SET oMustPrice = new CMustPrice
        oMustPrice.FRectIdx		= idx
        oMustPrice.getMustPirceOneItem

        mallgubun   = oMustPrice.FOneItem.FMallgubun
        itemid      = oMustPrice.FOneItem.FItemid
        mustPrice   = oMustPrice.FOneItem.FMustPrice
        mustMargin  = oMustPrice.FOneItem.FMustMargin
        startDate   = LEFT(oMustPrice.FOneItem.FStartDate, 10)
        endDate     = LEFT(oMustPrice.FOneItem.FEndDate, 10)
		startDateTime =	Num2Str(hour(oMustPrice.FOneItem.FStartDate),2,"0","R") & ":" & Num2Str(minute(oMustPrice.FOneItem.FStartDate),2,"0","R") & ":" & Num2Str(Second(oMustPrice.FOneItem.FStartDate),2,"0","R")
        endDateTime = Num2Str(hour(oMustPrice.FOneItem.FEndDate),2,"0","R") & ":" & Num2Str(minute(oMustPrice.FOneItem.FEndDate),2,"0","R") & ":" & Num2Str(Second(oMustPrice.FOneItem.FEndDate),2,"0","R")

        If IsNull(oMustPrice.FOneItem.FOrgpriceStartDate) Then
            orgpricestartDate   = ""
            orgpriceendDate     = ""
            orgpricestartDateTime =	""
            orgpriceendDateTime = ""
        Else
            orgpricestartDate   = LEFT(oMustPrice.FOneItem.FOrgpriceStartDate, 10)
            orgpriceendDate     = LEFT(oMustPrice.FOneItem.FOrgpriceEndDate, 10)
            orgpricestartDateTime =	Num2Str(hour(oMustPrice.FOneItem.FOrgpriceStartDate),2,"0","R") & ":" & Num2Str(minute(oMustPrice.FOneItem.FOrgpriceStartDate),2,"0","R") & ":" & Num2Str(Second(oMustPrice.FOneItem.FOrgpriceStartDate),2,"0","R")
            orgpriceendDateTime = Num2Str(hour(oMustPrice.FOneItem.FOrgpriceEndDate),2,"0","R") & ":" & Num2Str(minute(oMustPrice.FOneItem.FOrgpriceEndDate),2,"0","R") & ":" & Num2Str(Second(oMustPrice.FOneItem.FOrgpriceEndDate),2,"0","R")
        End If
    SET oMustPrice = nothing
End If
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
    if ($("#itemid").val() == "") {
        alert('상품코드를 입력하세요');
        $("#itemid").focus();
        return false;
    }
    if ($("#mustPrice").val() == "") {
        alert('특가를 입력하세요');
        $("#mustPrice").focus();
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
    if (confirm('저장 하시겠습니까?')){
        document.frm.submit();
    }
}
function numOnly(selector){
    selector.value = selector.value.replace(/[^0-9]/g,'');
}
function fnSelectMall(imallid){
    if(imallid == "nvstorefarm"){
        $("#orgpriceTr").show();
    } else {
        $("#orgpriceTr").hide();
        $("#orgpricetermSdt").val("");
        $("#orgpricetermSdtTime").val("");
        $("#orgpricetermEdt").val("");
        $("#orgpricetermEdtTime").val("");
    }
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="mustPrice_process.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">몰 구분</td>
    <td bgcolor="#FFFFFF">
        <select name="mallgubun" class="select" <%= Chkiif(mallgubun <> "", "disabled", "") %> onchange="fnSelectMall(this.value);">
            <option value="ssg" <%= Chkiif(mallgubun = "ssg", "selected", "") %> >SSG</option>
            <option value="coupang" <%= Chkiif(mallgubun = "coupang", "selected", "") %>>쿠팡</option>
            <option value="halfclub" <%= Chkiif(mallgubun = "halfclub", "selected", "") %>>하프클럽</option>
            <option value="hmall1010" <%= Chkiif(mallgubun = "hmall1010", "selected", "") %>>HMall</option>
            <option value="auction1010" <%= Chkiif(mallgubun = "auction1010", "selected", "") %>>옥션</option>
            <option value="shintvshopping" <%= CHKiif(mallgubun="shintvshopping","selected","") %> >신세계TV쇼핑</option>
            <option value="wetoo1300k" <%= CHKiif(mallgubun="wetoo1300k","selected","") %> >1300k</option>
            <option value="ezwel" <%= Chkiif(mallgubun = "ezwel", "selected", "") %>>이지웰페어</option>
            <option value="gmarket1010" <%= Chkiif(mallgubun = "gmarket1010", "selected", "") %>>G마켓</option>
            <option value="gsshop" <%= Chkiif(mallgubun = "gsshop", "selected", "") %>>GSShop</option>
            <option value="interpark" <%= Chkiif(mallgubun = "interpark", "selected", "") %>>인터파크</option>
            <option value="nvstorefarm" <%= Chkiif(mallgubun = "nvstorefarm", "selected", "") %>>스토어팜</option>
            <option value="Mylittlewhoopee" <%= Chkiif(mallgubun = "Mylittlewhoopee", "selected", "") %>>스토어팜 캣앤독</option>
            <option value="nvstoregift" <%= CHKiif(mallgubun="nvstoregift","selected","") %> >스토어팜 선물하기</option>
            <option value="WMP" <%= Chkiif(mallgubun = "WMP", "selected", "") %>>위메프</option>
            <option value="11st1010" <%= Chkiif(mallgubun = "11st1010", "selected", "") %>>11번가</option>
            <option value="lotteCom" <%= Chkiif(mallgubun = "lotteCom", "selected", "") %>>롯데닷컴</option>
            <option value="lotteimall" <%= Chkiif(mallgubun = "lotteimall", "selected", "") %>>롯데아이몰</option>
            <option value="lotteon" <%= Chkiif(mallgubun = "lotteon", "selected", "") %>>롯데On</option>
            <option value="skstoa" <%= CHKiif(mallgubun="skstoa","selected","") %> >SKSTOA</option>
            <option value="cjmall" <%= Chkiif(mallgubun = "cjmall", "selected", "") %>>CJMall</option>
            <option value="lfmall" <%= Chkiif(mallgubun = "lfmall", "selected", "") %>>LFmall</option>
            <option value="sabangnet" <%= Chkiif(mallgubun = "sabangnet", "selected", "") %>>사방넷</option>
            <option value="kakaogift" <%= Chkiif(mallgubun = "kakaogift", "selected", "") %>>카카오기프트</option>
            <option value="kakaostore" <%= Chkiif(mallgubun = "kakaostore", "selected", "") %>>카카오톡스토어</option>
            <option value="boribori1010" <%= Chkiif(mallgubun = "boribori1010", "selected", "") %>>보리보리</option>
            <option value="wconcept1010" <%= Chkiif(mallgubun = "wconcept1010", "selected", "") %>>W컨셉</option>
            <option value="benepia1010" <%= Chkiif(mallgubun = "benepia1010", "selected", "") %>>베네피아</option>
        </select>
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">상품코드</td>
    <td bgcolor="#FFFFFF">
        <textarea rows="2" cols="20" name="itemid" id="itemid" <%= Chkiif(itemid <> "", "disabled", "") %> ><%= itemid %></textarea>
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">특가</td>
    <td bgcolor="#FFFFFF">
        <input type="text" name="mustPrice" id="mustPrice" value="<%= mustPrice %>" onkeyup="numOnly(this)" onblur="numOnly(this)" />
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">특가시 마진</td>
    <td bgcolor="#FFFFFF">
        <input type="text" name="mustMargin" id="mustMargin" value="<%= mustMargin %>" size="3" />%
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

<tr height="25" bgcolor="<%= adminColor("gray") %>" id="orgpriceTr" <%= CHKIIF(mallgubun="nvstorefarm", "", "style='display:none;'") %> >
    <td width="20%">정상가 판매기간</td>
    <td bgcolor="#FFFFFF">
        <input type="text" id="orgpricetermSdt" name="orgpricestartDate" readonly size="11" maxlength="10" value="<%= orgpricestartDate %>" style="cursor:pointer; text-align:center;" />
        <input type="text" id="orgpricetermSdtTime" name="orgpricestartDateTime" size="8" maxlength="8" value="<%= orgpricestartDateTime %>" style="text-align:center;" /> ~
        <input type="text" id="orgpricetermEdt" name="orgpriceendDate" readonly size="11" maxlength="10" value="<%= orgpriceendDate %>" style="cursor:pointer; text-align:center;" />
        <input type="text" id="orgpricetermEdtTime" name="orgpriceendDateTime" size="8" maxlength="8" value="<%= orgpriceendDateTime %>" style="text-align:center;" />
        <script type="text/javascript">
            var CAL_Start = new Calendar({
                inputField : "orgpricetermSdt", trigger    : "orgpricetermSdt",
                onSelect: function() {
                    var date = Calendar.intToDate(this.selection.get());
                    CAL_End.args.min = date;
                    CAL_End.redraw();
                    this.hide();
                    if(frm.orgpricestartDateTime.value=="") frm.orgpricestartDateTime.value='00:00:00';
                    if(frm.orgpriceendDateTime.value=="") frm.orgpriceendDateTime.value='23:59:59';
                    if(frm.orgpriceendDate.value==""||getDayInterval(frm.orgpricestartDate.value, frm.orgpriceendDate.value) < 0) frm.orgpriceendDate.value=frm.orgpricestartDate.value;
                    doInsertDayInterval();	// 날짜 자동계산
                }, bottomBar: true, dateFormat: "%Y-%m-%d"
            });
            var CAL_End = new Calendar({
                inputField : "orgpricetermEdt", trigger    : "orgpricetermEdt",
                onSelect: function() {
                    var date = Calendar.intToDate(this.selection.get());
                    CAL_Start.args.max = date;
                    CAL_Start.redraw();
                    this.hide();

                    if(frm.orgpricestartDate.value==""||getDayInterval(frm.orgpricestartDate.value, frm.orgpriceendDate.value) < 0) frm.orgpricestartDate.value=frm.orgpriceendDate.value;
                    doInsertDayInterval();	// 날짜 자동계산
                }, bottomBar: true, dateFormat: "%Y-%m-%d"
            });
        </script>
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
