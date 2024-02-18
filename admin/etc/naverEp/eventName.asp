<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/common/commonCls.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<%
Dim page, oEvt, i, idx, oEvtOne, mallid, misusing, research, eventName, gubun
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
	SET oEvtOne = new epShop
		oEvtOne.FRectIdx = idx
		oEvtOne.FRectMallGubun = mallid
		oEvtOne.getEventStringOneItem

		gubun		= oEvtOne.FOneItem.FGubun
		startDate	= LEFT(oEvtOne.FOneItem.FStartDate, 10)
		endDate		= LEFT(oEvtOne.FOneItem.FEndDate, 10)
		eventName	= oEvtOne.FOneItem.FEventName
		isusing		= oEvtOne.FOneItem.FIsusing
	SET oEvtOne = nothing
End If

Set oEvt = new epShop
	oEvt.FCurrPage					= page
	oEvt.FPageSize					= 50
	oEvt.FRectMallGubun				= mallid
	oEvt.FRectIsusing				= misusing
	oEvt.getEventStringList
%>
<!-- #include virtual="/admin/etc/potal/inc_naverHead.asp" -->
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
function fnSaveMargin(){
	if ($("#gubun").val() == "2"){
		if ($("#termSdt").val() == "") {
			alert('시작일을 입력하세요');
			return false;
		}
		if ($("#termEdt").val() == "") {
			alert('종료일을 입력하세요');
			return false;
		}
	}
    if ($("#eventName").val() == "") {
		alert('이벤트문구를 입력하세요');
		$("#eventName").focus();
		return false;
    }
    if (confirm('저장 하시겠습니까?')){
		if ($("#idx").val() == "") {
			$("#mode").val("I");
		}else{
			$("#mode").val("U");
		}
        document.frmSave.target = "xLink";
        document.frmSave.submit();
    }
}
function fnViewTr(v){
	if(v == 1 || v ==''){
		$("#DateTr").hide();
		$("#isUsingTr").hide();
		if(v==''){
			$("#eventNameTr").hide();
		}else{
			$("#eventNameTr").show();
		}
	}else{
		$("#DateTr").show();
		$("#isUsingTr").show();
		$("#eventNameTr").show();
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSave" method="post" action="procEpShopEvent.asp" onsubmit="return false;">
<input type="hidden" name="mode" id="mode" value="">
<input type="hidden" name="idx" id="idx" value="<%= idx %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2" bgcolor="<%= adminColor("tabletop") %>">+ 기간별 이벤트문구 등록 및 수정</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" id="DateTr" <%= Chkiif(gubun <> "2", "style='display:none;'", "") %> >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">기간</td>
	<td align="LEFT">
        <input type="text" id="termSdt" name="startDate" readonly size="11" maxlength="10" value="<%= startDate %>" style="cursor:pointer; text-align:center;" />00:00:00 ~
        <input type="text" id="termEdt" name="endDate" readonly size="11" maxlength="10" value="<%= endDate %>" style="cursor:pointer; text-align:center;" />23:59:59
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
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">구분</td>
	<td align="LEFT">
	<%
		If idx <> "" Then
			Select Case gubun
				Case "1"		response.write "기본"
				Case "2"		response.write "기간별"
			End Select
			response.write "<input type='hidden' name='gubun' value='"& gubun &"'> "
		Else
	%>
		<select class="select" id="gubun" name="gubun" onchange="fnViewTr(this.value);">
			<option value="">-선택-</option>
			<option value="1">기본</option>
			<option value="2">기간별</option>
		</select>
	<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" id="eventNameTr" <%= Chkiif(idx = "", "style=""display:none;""", "") %> >
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">이벤트문구</td>
	<td align="LEFT">
		<input type="text" class="text" id="eventName" size="100" name="eventName" value="<%= eventName %>">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" id="isUsingTr" <%= Chkiif(gubun <> "2", "style='display:none;'", "") %>>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
	<td align="LEFT">
		<input type="radio" name="isusing" value="Y" <%= Chkiif(isusing="Y", "checked", "") %>>Y
		<input type="radio" name="isusing" value="N" <%= Chkiif(isusing="N", "checked", "") %> >N
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2">
		<input type="button" class="button" value="처음으로" onclick="location.replace('/admin/etc/naverEp/eventName.asp?menupos=<%=menupos%>&mallid=nvshop');">
		<input type="button" class="button" value="저장" onclick="fnSaveMargin();">
	</td>
</tr>
</form>
</table>

<br /><br />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="mallid" value="<%= mallid %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="3" bgcolor="<%= adminColor("tabletop") %>">기간별 이벤트문구 리스트</td>
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
	<td colspan="6">
		검색결과 : <b><%= FormatNumber(oEvt.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oEvt.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">기간</td>
	<td width="100">구분</td>
    <td>이벤트문구</td>
	<td width="100">사용여부</td>
	<td width="100">등록일</td>
	<td width="100">관리</td>
</tr>
<% For i=0 to oEvt.FResultCount - 1 %>
<tr align="center" bgcolor="<%= Chkiif(oEvt.FItemList(i).FGubun="1", "YELLOW", "#FFFFFF") %>">
	<td>
	<%
		If oEvt.FItemList(i).FGubun <> "1" Then
			response.write LEFT(oEvt.FItemList(i).FStartDate, 10) &" ~ "&  LEFT(oEvt.FItemList(i).FEndDate, 10)
		End If
	%>
	</td>
	<td>
	<%
		Select Case oEvt.FItemList(i).FGubun
			Case "1"	response.write "기본"
			Case "2"	response.write "기간별"
		End Select
	%>
	</td>
	<td><%= oEvt.FItemList(i).FEventName %></td>
	<td><%= oEvt.FItemList(i).FIsusing %></td>
	<td><%= LEFT(oEvt.FItemList(i).FRegDate, 10) %></td>
	<td><input type="button" class="button" value="수정" onclick="javascript:location.href='/admin/etc/naverEp/eventName.asp?menupos=<%=menupos%>&idx=<%= oEvt.FItemList(i).FIdx %>&mallid=<%= mallid %>';"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oEvt.HasPreScroll then %>
		<a href="javascript:goPage('<%= oEvt.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oEvt.StartScrollPage to oEvt.FScrollCount + oEvt.StartScrollPage - 1 %>
    		<% if i>oEvt.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oEvt.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set oEvt = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->