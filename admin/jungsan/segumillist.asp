<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 전자세금계산서 발급 관리
' History : 2017.09.21 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/jungsan/electaxcls.asp"-->
<%
Dim yyyy, year_from, year_to, i
yyyy = request("yyyy")

If yyyy	= "" Then yyyy	= Year(now)
year_from = "2020"
year_to = Year(now) + 20

Dim ojungsan, arrList
set ojungsan = new CElecTaxReg
	ojungsan.FRectyyyy = yyyy
	arrList = ojungsan.fnTaxdateList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function frmCheck(idx, ymd){
	var taxdate;
	taxdate = $("#a"+idx).val();
	if (taxdate == '' ){
		alert('날짜를 입력하세요');
		$("#a"+idx).focus();
		return;
	}

	if (confirm(''+ymd+' 데이터를 저장 하시겠습니까?')){
		$("#taxdate").val(taxdate);
		$("#idx").val(idx);
		document.frmSvArr.target = "xLink";
		document.frmSvArr.action = "/admin/jungsan/proc_segumil.asp"
		document.frmSvArr.submit();
	}
}
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
</script>
<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		연월 : 
		<select name="yyyy" class="select">
		<% For i = year_from to year_to %>
			<option value="<%= i %>" <%= Chkiif(CInt(yyyy) = i, "selected", "") %>><%= i %></option>
		<% Next %>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
</form>
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">정산월</td>
	<td width="200">작업연월</td>
	<td>발급기준일</td>
	<td width="150">등록자ID</td>
	<td width="100">관리</td>
</tr>
<%
	if isArray(arrList) then
		for i = 0 to ubound(arrList,2) 
%>
<tr align="center" bgcolor="#FFFFFF" height="50">
	<td><%= LEFT(Dateadd("m",-1,arrList(1,i)), 7) %></td>
	<td><%=arrList(1,i)%></td>
	<td align="center">
		<input type="text" size="13" id="a<%= arrList(0,i) %>"  name="a<%= arrList(0,i) %>" value="<%=arrList(2,i)%>" onClick="jsPopCal('a<%= arrList(0,i) %>');" style="cursor:hand;">
	</td>
	<td align="center"><%= arrList(3,i) %></td>
	<td width="100"><input type="button" class="button_s" value="저장" onclick="frmCheck('<%=arrList(0,i)%>', '<%=arrList(1,i)%>')"></td>
</tr>
<%
		Next
	End If
%>
</table>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" id="idx" name="idx" value="">
<input type="hidden" id="taxdate" name="taxdate" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="10"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->