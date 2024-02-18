<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/shintvshopping/shintvshoppingCls.asp"-->
<%
Dim page, oShintvshopping, i, idx, oShintvshoppingMaster, misusing, research
page		= request("page")
idx			= request("idx")
misusing	= request("isusing")
research	= request("research")

If (research = "") Then
	misusing = "Y"
End If
If page = "" Then page = 1

Dim startDate, endDate, shipCostCode, isusing, startDateTime, endDateTime
isusing = "Y"
If idx <> "" Then
	SET oShintvshoppingMaster = new CShintvshopping
		oShintvshoppingMaster.FRectIdx = idx
		oShintvshoppingMaster.getShintvshoppingshipCostCodeItemOneItem
        startDate		= LEFT(oShintvshoppingMaster.FOneItem.FStartDate, 10)
        endDate			= LEFT(oShintvshoppingMaster.FOneItem.FEndDate, 10)
		startDateTime	= Num2Str(hour(oShintvshoppingMaster.FOneItem.FStartDate),2,"0","R") & ":" & Num2Str(minute(oShintvshoppingMaster.FOneItem.FStartDate),2,"0","R") & ":" & Num2Str(Second(oShintvshoppingMaster.FOneItem.FStartDate),2,"0","R")
        endDateTime		= Num2Str(hour(oShintvshoppingMaster.FOneItem.FEndDate),2,"0","R") & ":" & Num2Str(minute(oShintvshoppingMaster.FOneItem.FEndDate),2,"0","R") & ":" & Num2Str(Second(oShintvshoppingMaster.FOneItem.FEndDate),2,"0","R")
		shipCostCode	= oShintvshoppingMaster.FOneItem.FShipCostCode
		isusing			= oShintvshoppingMaster.FOneItem.FIsusing
	SET oShintvshoppingMaster = nothing
End If

Set oShintvshopping = new CShintvshopping
	oShintvshopping.FCurrPage					= page
	oShintvshopping.FPageSize					= 50
	oShintvshopping.FRectIsusing				= misusing
	oShintvshopping.getShintvshoppingBeasongCodeItemList
%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function popBeasongDetail(v){
	var popdetail=window.open('/admin/etc/shintvshopping/popshintvshoppingBeasongCodeItemDetail.asp?midx='+v,'popMarginDetail','width=700,height=300,scrollbars=yes,resizable=yes');
	popdetail.focus();
}
function fnSaveCode(){
    if ($("#termSdt").val() == "") {
        alert('시작일을 입력하세요');
        return false;
    }
    if ($("#termEdt").val() == "") {
        alert('종료일을 입력하세요');
        return false;
    }
    if ($("#shipCostCode").val() == "") {
        alert('배송비코드를 입력하세요');
        $("#shipCostCode").focus();
        return false;
    }
    if (confirm('저장 하시겠습니까?')){
        document.frmSave.target = "xLink";
        document.frmSave.submit();
    }
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSave" method="post" action="procShintvshoppingBeasongCode.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="itemMaster">
<input type="hidden" name="idx" value="<%= idx %>">
<tr align="LEFT" bgcolor="#FFFFFF" >
	<td colspan="2" bgcolor="<%= adminColor("tabletop") %>">
		기본배송비코드는 <strong><font color="red">B01</font></strong>(5만원이상 무료배송) 입니다.</br>
		기본배송비코드외의 배송비정책을 사용시 등록하는 기능입니다.</br>
		미리 신세계TV쇼핑에서 배송비코드를 채번 후 사용해야 동작하게 됩니다.
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">기간</td>
    <td bgcolor="#FFFFFF" align="LEFT">
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
                    if(frmSave.startDateTime.value=="") frmSave.startDateTime.value='00:00:00';
                    if(frmSave.endDateTime.value=="") frmSave.endDateTime.value='23:59:59';
                    if(frmSave.endDate.value==""||getDayInterval(frmSave.startDate.value, frmSave.endDate.value) < 0) frmSave.endDate.value=frmSave.startDate.value;
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

                    if(frmSave.startDate.value==""||getDayInterval(frmSave.startDate.value, frmSave.endDate.value) < 0) frmSave.startDate.value=frmSave.endDate.value;
                    doInsertDayInterval();	// 날짜 자동계산
                }, bottomBar: true, dateFormat: "%Y-%m-%d"
            });
        </script>
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">배송비코드</td>
	<td align="LEFT">
		<input type="text" id="shipCostCode" size="3" name="shipCostCode" value="<%= shipCostCode %>">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
	<td align="LEFT">
		<input type="radio" name="isusing" value="Y" <%= Chkiif(isusing="Y", "checked", "") %>>Y
		<input type="radio" name="isusing" value="N" <%= Chkiif(isusing="N", "checked", "") %> >N
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2">
		<input type="button" class="button" value="저장" onclick="fnSaveCode();">
	</td>
</tr>
</form>
</table>

<br />
<hr style="border:solid 3px;" />
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="3" bgcolor="<%= adminColor("tabletop") %>">기간별 배송비코드 리스트</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		사용여부 :
		<select class="select" name="isusing">
			<option value="">전체</option>
			<option value="Y" <%= Chkiif(misusing="Y", "selected", "") %> >Y</option>
			<option value="N" <%= Chkiif(misusing="N", "selected", "") %>>N</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<br />
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		검색결과 : <b><%= FormatNumber(oShintvshopping.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oShintvshopping.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>기간</td>
    <td width="100">배송비코드</td>
	<td width="100">사용여부</td>
	<td width="100">등록일</td>
	<td width="100">관리</td>
</tr>
<% For i=0 to oShintvshopping.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td style="cursor:pointer;" onclick="popBeasongDetail('<%= oShintvshopping.FItemList(i).FIdx %>');">
		<%= FormatDate(oShintvshopping.FItemList(i).FStartDate,"0000-00-00 00:00:00") %> ~ <%= FormatDate(oShintvshopping.FItemList(i).FEndDate,"0000-00-00 00:00:00") %>
	</td>
	<td><%= oShintvshopping.FItemList(i).FShipCostCode %></td>
	<td><%= oShintvshopping.FItemList(i).FIsusing %></td>
	<td><%= LEFT(oShintvshopping.FItemList(i).FRegDate, 10) %></td>
	<td><input type="button" class="button" value="수정" onclick="javascript:location.href='/admin/etc/shintvshopping/popShintvshoppingDeliveryCodeItem.asp?idx=<%= oShintvshopping.FItemList(i).FIdx %>';"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oShintvshopping.HasPreScroll then %>
		<a href="javascript:goPage('<%= oShintvshopping.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oShintvshopping.StartScrollPage to oShintvshopping.FScrollCount + oShintvshopping.StartScrollPage - 1 %>
    		<% if i>oShintvshopping.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oShintvshopping.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set oShintvshopping = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
