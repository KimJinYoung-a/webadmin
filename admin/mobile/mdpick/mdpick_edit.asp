<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : 모바일 mdpick
' History : 2013.12.17 한용민
'###############################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incsessionadmin.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/mobile/mdpick_cls.asp" -->

<%
dim idx, itemid, isusing, orderno, regdate, lastdate, regadminid, lastadminid, menupos, startdate, enddate
	idx = request("idx")
	menupos = request("menupos")

dim omdpick, i
set omdpick = new cmdpick
	omdpick.frectidx = idx
	
	if idx <> "" then
		omdpick.getmdpick_one()
		
		if omdpick.ftotalcount > 0 then
			idx = omdpick.FOneItem.fidx
			itemid = omdpick.FOneItem.fitemid
			isusing = omdpick.FOneItem.fisusing
			orderno = omdpick.FOneItem.forderno
			regdate = omdpick.FOneItem.fregdate
			lastdate = omdpick.FOneItem.flastdate
			regadminid = omdpick.FOneItem.fregadminid
			lastadminid = omdpick.FOneItem.flastadminid
			startdate = omdpick.FOneItem.fstartdate
			enddate = omdpick.FOneItem.fenddate
		end if
	end if
	
if orderno="" then orderno=99
if isusing="" then isusing="Y"
%>

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">

function mdpickproc(){
	if (frm.orderno.value==''){
		alert('정렬순위를 입력해 주세요.');
		frm.orderno.focus();
		return;
	}
    //if (frm.startdate.value.length!=10){
    //    alert('시작일을 입력  하세요.');
    //    return;
    //}
    //if (frm.enddate.value.length!=10){
    //    alert('종료일을 입력  하세요.');
    //    return;
    //}
    //var vstartdate = new Date(frm.startdate.value.substr(0,4), (1*frm.startdate.value.substr(5,2))-1, frm.startdate.value.substr(8,2));
    //var venddate = new Date(frm.enddate.value.substr(0,4), (1*frm.enddate.value.substr(5,2))-1, frm.enddate.value.substr(8,2));
    
    //if (vstartdate>venddate){
    //    alert('종료일이 시작일보다 빠르면 안됩니다.');
    //    return;
    //}    
	if (!IsDouble(frm.orderno.value)){
		alert('정렬순위는 숫자만 가능합니다.');
		frm.orderno.focus();
		return;
	}
	if (frm.isusing.value==''){
		alert('사용여부를 선택해 주세요.');
		frm.isusing.focus();
		return;
	}
	
	frm.submit();	
}
	
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ MDPICK 등록
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post" action="/admin/mobile/mdpick/mdpick_process.asp">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="mdpickedit">
<input type="hidden" name="idx" value="<%=idx%>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>상품코드</b><br></td>
	<td bgcolor="#FFFFFF">
		<%= itemid %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>사용여부</b><br></td>
	<td bgcolor="#FFFFFF">
		<% drawSelectBoxisusingYN "isusing", isusing, "" %>
	</td>
</tr>
<!--<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>반영시작일</b><br></td>
	<td bgcolor="#FFFFFF">
        <input id="startdate" name="startdate" value="<%= Left(startdate,10) %>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
        <input type="text" name="dummy0" value="00:00:00" size="8" readonly class="text_ro" />
	    <script type="text/javascript">
		var CAL_Start = new Calendar({
			inputField : "startdate",
			trigger    : "startdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
	</td>
</tr>-->
<!--<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>반영시작일</b><br></td>
	<td bgcolor="#FFFFFF">
        <input id="enddate" name="enddate" value="<%= Left(enddate,10) %>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
        <input type="text" name="dummy1" value="23:59:59" size="8" readonly class="text_ro" />
	    <script type="text/javascript">
		var CAL_End = new Calendar({
			inputField : "enddate",
			trigger    : "enddate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
	</td>
</tr>-->
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>정렬순위</b><br></td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="orderno" value="<%= orderno %>" size=3 maxlength=3>
	</td>
</tr>
<% if lastadminid<>"" then %>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center"><b>최근수정</b><br></td>
		<td bgcolor="#FFFFFF">
			<%= lastdate %>
			<Br>(<%= lastadminid %>)
		</td>
	</tr>
<% end if %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" value="저장" onclick="mdpickproc();" class="button">
	</td>
</tr>
</form>	
</table>

<%
set omdpick = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
