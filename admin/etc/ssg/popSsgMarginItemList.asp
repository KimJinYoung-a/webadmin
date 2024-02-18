<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Dim page, oSsg, i, idx, oSsgMaster, misusing, research, mallid
page		= request("page")
idx			= request("idx")
mallid		= request("mallid")
misusing	= request("isusing")
research	= request("research")

If (research = "") Then
	misusing = "Y"
End If
If page = "" Then page = 1

Dim startDate, endDate, margin, isusing
isusing = "Y"
If idx <> "" Then
	SET oSsgMaster = new Cssg
		oSsgMaster.FRectIdx = idx
		oSsgMaster.FRectMallGubun = mallid
		oSsgMaster.getSsgMarginItemOneItem

		startDate = oSsgMaster.FOneItem.FStartDate
		endDate = oSsgMaster.FOneItem.FEndDate
		margin 	= oSsgMaster.FOneItem.FMargin
		isusing = oSsgMaster.FOneItem.FIsusing
	SET oSsgMaster = nothing
End If

Set oSsg = new Cssg
	oSsg.FCurrPage					= page
	oSsg.FPageSize					= 50
	oSsg.FRectMallGubun				= mallid
	oSsg.FRectIsusing				= misusing
	oSsg.getssgMarginItemList
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
function popMarginDetail(v){
	var popdetail=window.open('/admin/etc/ssg/popSsgMarginItemDetail.asp?mallid=<%= mallid %>&midx='+v,'popMarginDetail','width=700,height=300,scrollbars=yes,resizable=yes');
	popdetail.focus();
}
function fnSaveMargin(){
    if ($("#termSdt").val() == "") {
        alert('시작일을 입력하세요');
        return false;
    }
    if ($("#termEdt").val() == "") {
        alert('종료일을 입력하세요');
        return false;
    }
    if ($("#margin").val() == "") {
        alert('마진을 입력하세요');
        $("#margin").focus();
        return false;
    }
    if (confirm('저장 하시겠습니까?')){
        document.frmSave.target = "xLink";
        document.frmSave.submit();
    }
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSave" method="post" action="procSsgMargin.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="itemMaster">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2" bgcolor="<%= adminColor("tabletop") %>">+ 기간별 마진 등록 및 수정(상품)</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">기간</td>
	<td align="LEFT">
        <input type="text" id="termSdt" name="startDate" readonly size="11" maxlength="10" value="<%= startDate %>" style="cursor:pointer; text-align:center;" /> ~
        <input type="text" id="termEdt" name="endDate" readonly size="11" maxlength="10" value="<%= endDate %>" style="cursor:pointer; text-align:center;" />
        <script type="text/javascript">
            var CAL_Start = new Calendar({
                inputField : "termSdt", trigger    : "termSdt",
                onSelect: function() {
                    var date = Calendar.intToDate(this.selection.get());
                    CAL_End.args.min = date;
                    CAL_End.redraw();
                    this.hide();

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
<tr align="center" bgcolor="#FFFFFF" >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">마진</td>
	<td align="LEFT">
		<input type="text" id="margin" size="3" name="margin" value="<%= margin %>">%
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
		<input type="button" class="button" value="저장" onclick="fnSaveMargin();">
	</td>
</tr>
</form>
</table>

<br /><br />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="3" bgcolor="<%= adminColor("tabletop") %>">기간별 마진 리스트</td>
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
		검색결과 : <b><%= FormatNumber(oSsg.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oSsg.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>기간</td>
    <td width="100">적용마진</td>
	<td width="100">사용여부</td>
	<td width="100">등록일</td>
	<td width="100">관리</td>
</tr>
<% For i=0 to oSsg.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td style="cursor:pointer;" onclick="popMarginDetail('<%= oSsg.FItemList(i).FIdx %>');"><%= oSsg.FItemList(i).FStartDate %> ~ <%= oSsg.FItemList(i).FEndDate %></td>
	<td><%= oSsg.FItemList(i).FMargin %>%</td>
	<td><%= oSsg.FItemList(i).FIsusing %></td>
	<td><%= LEFT(oSsg.FItemList(i).FRegDate, 10) %></td>
	<td><input type="button" class="button" value="수정" onclick="javascript:location.href='/admin/etc/ssg/popssgMarginItemList.asp?idx=<%= oSsg.FItemList(i).FIdx %>&mallid=<%= mallid %>';"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oSsg.HasPreScroll then %>
		<a href="javascript:goPage('<%= oSsg.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oSsg.StartScrollPage to oSsg.FScrollCount + oSsg.StartScrollPage - 1 %>
    		<% if i>oSsg.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oSsg.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set oSsg = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
