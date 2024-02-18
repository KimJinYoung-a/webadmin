<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/que/queItemCls.asp"-->
<%
Dim itemid, startdate, enddate, mode
itemid  	= request("itemid")
startdate  	= request("startdate")
enddate  	= request("enddate")
mode		= request("mode")

If mode = "I" Then
	If itemid<>"" then
		Dim iA, arrTemp, arrItemid
		itemid = replace(itemid,",",chr(10))
		itemid = replace(itemid,chr(13),"")
		arrTemp = Split(itemid,chr(10))
		iA = 0
		Do While iA <= ubound(arrTemp)
			If Trim(arrTemp(iA))<>"" then
				If Not(isNumeric(trim(arrTemp(iA)))) then
					Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
					dbget.close()	:	response.End
				Else
					arrItemid = arrItemid & trim(arrTemp(iA)) & ","
				End If
			End If
			iA = iA + 1
		Loop
		itemid = left(arrItemid,len(arrItemid)-1)
	End If

	'insert db_temp.dbo.tbl_tmp_ScheduleSplit (title,itemid ,startdate,enddate) values('수정금지','아이템아디','시작기간','종료기간')
	Dim i, spItemid, strSql
	spItemid = Split(itemid, ",")
	For i=0 to Ubound(spItemid)
		strSql = ""
		strSql = strSql & " IF NOT EXISTS (SELECT TOP 1 itemid FROM  db_temp.dbo.tbl_tmp_ScheduleSplit WHERE itemid = '" & spItemid(i) & "' and startdate = '"& startdate &"' and enddate = '"& enddate &"' ) "
		strSql = strSql & " BEGIN "
		strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_tmp_ScheduleSplit (title, itemid, startdate, enddate) VALUES ('수정금지','" & spItemid(i) & "', '"& startdate &"', '"& enddate &"')"
		strSql = strSql & " END "
		dbget.execute strSql
	Next
	response.redirect("/admin/etc/que/notinsche.asp")
End If
%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script>
function goPage(pg){
    frms.page.value = pg;
    frms.submit();
}
function frmSubmit(){
	if($("#itemid").val() == ''){
		alert("상품코드를 입력하세요");
		$("#itemid").focus();
		return;
	}

	if($("#termSdt").val() == ''){
		alert("시작일을 입력하세요");
		$("#termSdt").focus();
		return;
	}

	if($("#termEdt").val() == ''){
		alert("종료일을 입력하세요");
		$("#termEdt").focus();
		return;
	}

	if(confirm("저장하시겠습니까?")){
		document.frm.submit();
	}
}
</script>



<form name="frms" method="get" action="">
<input type= "hidden" name="page" value="<%= page %>">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type= "hidden" name="mode" value="I">
<tr align="center" bgcolor="#FFFFFF" >
	<td>
		<h2>가격 수정 홀딩 임시페이지</h2>
	</td>
</tr>
</table>
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="">
<input type= "hidden" name="mode" value="I">
<tr align="center" bgcolor="#FFFFFF" >
	<td>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
	</td>
	<td>
		시작일 :
        <input type="text" id="termSdt" name="startDate" readonly size="11" maxlength="10" value="<%= startDate %>" style="cursor:pointer; text-align:center;" />
		종료일 :
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
	<td colspan="2">
		<input type="button" class="button_s" value="저장" onClick="frmSubmit();">
	</td>
</tr>
</form>
</table>
<br />
<h3>리스트</h3>
<%
Dim oOutmall, j, page, pagesize
page 		= request("page")
pagesize	= request("pagesize")

If page = "" Then page = 1
If pagesize = "" Then pagesize = 100

Set oOutmall = new COutmall
	oOutmall.FPageSize 			= pagesize
	oOutmall.FCurrPage			= page
	oOutmall.getInboundNotScheduleitemList
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
		검색결과 : <b><%= FormatNumber(oOutmall.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oOutmall.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상품코드</td>
	<td>제목</td>
	<td>시작일</td>
	<td>종료일</td>
	<td>등록일</td>
</tr>
<% For i = 0 To oOutmall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= oOutmall.FItemlist(i).FItemid %></td>
	<td><%= oOutmall.FItemlist(i).FTitle %></td>
	<td><%= oOutmall.FItemlist(i).FStartdate %></td>
	<td><%= oOutmall.FItemlist(i).FEnddate %></td>
	<td><%= oOutmall.FItemlist(i).FRegdate %></td>
</tr>
<%
	Next
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14" align="center">
	<% If oOutmall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oOutmall.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oOutmall.StartScrollPage To oOutmall.FScrollCount + oOutmall.StartScrollPage - 1 %>
		<% If i>oOutmall.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oOutmall.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</table>
<% Set oOutmall = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->