<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 데이터분석 영업이슈
' History : 2016.01.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_salesissue_cls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
dim i, menupos, salesidx, department_id,startdate,enddate,title,comment,reguserid,regdate,isusing, startdatetime, enddatetime
dim username
	menupos = requestCheckVar(request("menupos"),10)
	salesidx = requestCheckVar(request("salesidx"),10)

dim csales
SET csales = New cdataanalysis_salesissue
	csales.frectsalesidx = salesidx

	if salesidx<>"" then
		csales.getdataanalysis_salesissue_oneitem()

		if csales.FtotalCount > 0 then
			department_id = csales.FOneItem.fdepartment_id
			startdate = left(csales.FOneItem.fstartdate,10)
				startdatetime=Right(csales.FOneItem.fstartdate,8)
			enddate = left(csales.FOneItem.fenddate,10)
				enddatetime=Right(csales.FOneItem.fenddate,8)
			title = csales.FOneItem.ftitle
			comment = csales.FOneItem.fcomment
			reguserid = csales.FOneItem.freguserid
			regdate = csales.FOneItem.fregdate
			isusing = csales.FOneItem.fisusing
			username = csales.FOneItem.fusername
		end if
	end if

if startdate="" then
	if startdate = "" or isnull(startdate) then startdate = date()
end if
if enddate="" then
	if enddate = "" or isnull(enddate) then enddate = dateadd("d", +7, date())
end if
if startdatetime="" then startdatetime="00:00:00"
if enddatetime="" then enddatetime="23:59:59"
if reguserid="" then reguserid=session("ssBctId")
if isusing="" then isusing="Y"
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function jssave() {
	if (frm.department_id.value==''){
		alert('부서를 선택해 주세요.');
		return false;
	}
	if (frm.startdate.value==''){
		alert('시작일을 입력해 주세요.');
		frm.startdate.focus();
		return false;
	}
	if (frm.enddate.value==''){
		alert('종료일을 입력해 주세요.');
		frm.enddate.focus();
		return false;
	}
	if (frm.title.value==''){
		alert('프로젝트명을 입력해 주세요.');
		frm.title.focus();
		return false;
	}
	if (frm.comment.value==''){
		alert('설명(목적/결과)을 입력해 주세요.');
		frm.comment.focus();
		return false;
	}
	if (frm.isusing.value==''){
		alert('사용여부를 선택해 주세요.');
		return false;
	}

	if(confirm("정보를 저장 하시겠습니까?")) {
		//frm.target="ifproc";
		frm.mode.value="salesissuereg";
		frm.action="/admin/dataanalysis/salesissue/salesissue_process.asp";
		frm.submit();
	}
}

</script>

<form name="frm" method="post" action="" style="margin:0;">
<input type="hidden" name="mode" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%=adminColor("sky")%>"><b>번호</b><br></td>
	<td>
		<%= salesidx %><input type="hidden" name="salesidx" value="<%= salesidx %>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%=adminColor("sky")%>"><b>부서</b><br></td>
	<td>
		<%= drawSelectBoxDepartmentALL("department_id", department_id) %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%=adminColor("sky")%>"><b>시작일</b><br></td>
	<td>
		<input id='startdate' name='startdate' value='<%= startdate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		<input type="text" class="text_ro" name="startdatetime" value="<%= startdatetime %>" size="8" maxlength="8" />
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%=adminColor("sky")%>"><b>종료일</b><br></td>
	<td>
		<input id='enddate' name='enddate' value='<%= enddate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		<input type="text" class="text_ro" name="enddatetime" value="<%= enddatetime %>" size="8" maxlength="8" />
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%=adminColor("sky")%>"><b>프로젝트명</b><br></td>
	<td>
		<input type="text" name="title" value="<%= title %>" size=60>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%=adminColor("sky")%>"><b>설명(목적/결과)</b><br></td>
	<td>
		<textarea name="comment" rows="5" cols="60"><%= comment %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%=adminColor("sky")%>"><b>등록</b><br></td>
	<td>
		<%= regdate %>
		<br><%= username %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%=adminColor("sky")%>"><b>사용여부</b><br></td>
	<td>
		<select name="isusing" class="select">
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>사용</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>사용안함</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan="2">
		<input type="button" onClick="jssave();" value="저장" class="button">
	</td>
</tr>
</table>
</form>

<iframe id="ifproc" name="ifproc" width=0 height=0></iframe>

<script type="text/javascript">
	var CAL_Start = new Calendar({
		inputField : "startdate", trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	var CAL_End = new Calendar({
		inputField : "enddate", trigger    : "enddate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
</script>

<%
set csales=nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->