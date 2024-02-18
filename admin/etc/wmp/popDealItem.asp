<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/etc/wmp/wmpCls.asp"-->
<%
Dim mode, idx, isModify
Dim oDealItem
isModify    = request("isModify")
idx         = request("idx")

If isModify = "" Then
    mode = "I"
Else
    mode = "U"
End If

Dim itemid, startDate, endDate, startDateTime, endDateTime, newItemname, itemname, limitCount
If mode = "U" Then
    SET oDealItem = new CWmp
        oDealItem.FRectIdx		= idx
        oDealItem.getDealOneItem

        itemid      = oDealItem.FOneItem.FItemid
        itemname    = oDealItem.FOneItem.FItemname
		newItemname = oDealItem.FOneItem.FNewItemname
        limitCount  = oDealItem.FOneItem.FLimitCount
        startDate   = LEFT(oDealItem.FOneItem.FStartDate, 10)
        endDate     = LEFT(oDealItem.FOneItem.FEndDate, 10)
		startDateTime =	Num2Str(hour(oDealItem.FOneItem.FStartDate),2,"0","R") & ":" & Num2Str(minute(oDealItem.FOneItem.FStartDate),2,"0","R") & ":" & Num2Str(Second(oDealItem.FOneItem.FStartDate),2,"0","R")
        endDateTime = Num2Str(hour(oDealItem.FOneItem.FEndDate),2,"0","R") & ":" & Num2Str(minute(oDealItem.FOneItem.FEndDate),2,"0","R") & ":" & Num2Str(Second(oDealItem.FOneItem.FEndDate),2,"0","R")
    SET oDealItem = nothing
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
    if ($("#newItemname").val() == "") {
        alert('변경상품명을 입력하세요');
        $("#newItemname").focus();
        return false;
    }
    if ($("#termSdt").val() == "") {
        alert('시작일을 입력하세요');
        return false;
    }
    if ($("#termEdt").val() == "") {
        alert('종료일을 입력하세요');
        return false;
    }
    if (confirm('저장 하시겠습니까?')){
        document.frm.submit();
    }
}

function numOnly(selector){
    selector.value = selector.value.replace(/[^0-9]/g,'');
}

function jsCheckView(){
    if ($("#itemid").val() == "") {
        alert('상품코드를 입력하세요');
        $("#itemid").focus();
        return false;
    }
	$.ajax({
		type: "POST",
		url: "/admin/etc/wmp/dealoptionajax.asp?itemid=" + $("#itemid").val(),
		dataType: "text",
		async: false,
		success : function(result){
		    $("#jsCheck").empty().html(result);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}
</script>
<form name="frm" method="post" action="procDealItem.asp" onsubmit="return false;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">상품코드</td>
    <td bgcolor="#FFFFFF">
		<input type="text" id="itemid" name="itemid" size="10" maxlength="10" <%= Chkiif(itemid <> "", "disabled", "") %> value="<%= itemid %>" />
    <% If mode = "I" Then %>
		<input type="button" class="button" value="확인" onclick=jsCheckView(); >
    <% End If %>
    </td>
</tr>
<tr id="itemnameTr" height="25" bgcolor="<%= adminColor("gray") %>" <%= Chkiif(mode="I", "style='display:none'", "")  %> >
    <td width="20%">상품명</td>
    <td bgcolor="#FFFFFF">
		<input type="text" id="itemname" name="itemname" size="60" disabled value="<%= itemname %>" />
    </td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>">
    <td width="20%">변경상품명</td>
    <td bgcolor="#FFFFFF">
		<input type="text" id="newItemname" name="newItemname" size="60" value="<%= newItemname %>" />
    </td>
</tr>
<tr id="limitCountTr" height="25" bgcolor="<%= adminColor("gray") %>" <%= Chkiif(mode="I", "style='display:none'", "")  %>>
    <td width="20%">재고</td>
    <td colspan="2" width="80%" align="LEFT" bgcolor="#FFFFFF">
        <input type="text" name="limitCount" id="limitCount" size="3" value="<%= limitCount %>">
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
<tr height="25" bgcolor="<%= adminColor("gray") %>" align="center">
    <td bgcolor="#FFFFFF" colspan="3">
        <input type="button" value="저장" class="button" onclick="frm_check();" />
    </td>
</tr>
</table>
</form>
<div id="jsCheck"></div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
