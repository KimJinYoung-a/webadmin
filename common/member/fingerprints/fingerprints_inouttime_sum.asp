<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 지문인식 근태관리
' Hieditor : 2011.03.22 한용민 생성
'            2012.02.15 허진원 - 미니달력 교체
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/member/fingerprints/fingerprints_cls.asp" -->

<%
Dim ofingerprints,i, part_sn ,sDt ,eDt ,empno
	sDt = requestCheckVar(request("sDt"),10)
	eDt = requestCheckVar(request("eDt"),10)
	empno = requestCheckVar(request("empno"),32)
	part_sn = requestCheckVar(request("part_sn"),10)
	menupos = requestCheckVar(request("menupos"),10)

if sDt = "" then sDt = date
if eDt = "" then eDt = date

'/관리자가 아닐경우 본인 파트만 확인 가능
if getlevel_sn("",session("ssBctID")) >= "3" then
	part_sn = getpart_sn("",session("ssBctID"))
end if

set ofingerprints = new cfingerprints_list
	ofingerprints.frectpart_sn = part_sn
	ofingerprints.frectempno = empno
	ofingerprints.FrectSDate = sDt
	ofingerprints.FrectEDate = dateadd("d",1,eDt)	
	ofingerprints.ffingerprints_sum()
	
%>

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		사원번호 : <input type="text" name="empno" value="<%=empno%>" size=16 maxlength=16>
		부서:
		<%=printPartOption("part_sn", part_sn)%>
		기간 : 
		<input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "sDt", trigger    : "sDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "eDt", trigger    : "eDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">	
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">		
	</td>
	<td align="right">			
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ofingerprints.FTotalCount %></b> ※총 3,000건 까지 검색 됩니다
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">	
	<td>출근기준일</td>
	<td>사원번호</td>		
	<td>성명</td>
	<td>출근시간</td>
	<td>퇴근시간</td>
	<td>총근무시간</td>
	<td>출퇴근(분)</td>
	<td>외출(분)</td>
	<td>외출횟수</td>
</tr>
<% if ofingerprints.FresultCount>0 then %>
<% for i=0 to ofingerprints.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#ffffff';>
	<td>
		<%= ofingerprints.FItemList(i).fyyyymmdd %>
	</td>
	<td>
		<%= ofingerprints.FItemList(i).fempno %>
	</td>
	<td>
		<%= ofingerprints.FItemList(i).fusername %>
	</td>
	<td>
		<%= FormatDate(ofingerprints.FItemList(i).fInTime,"0000-00-00 00:00:00") %>
	</td>
		
	<td>
		<%
		if ofingerprints.FItemList(i).fOutTime <> "1900-01-01" then
			response.write FormatDate(ofingerprints.FItemList(i).fOutTime,"0000-00-00 00:00:00")
		end if
		%>		
	</td>
	<td>		
		<%= minutechagehour(ofingerprints.FItemList(i).fworkmin - ofingerprints.FItemList(i).fexmin) %>
	</td>
		
	<td>
		<%= ofingerprints.FItemList(i).fworkmin %>
	</td>
	<td>
		<%= ofingerprints.FItemList(i).fexmin %>
	</td>
		
	<td>
		<%= ofingerprints.FItemList(i).freinCNT %>/<%= ofingerprints.FItemList(i).freoutCNT %>
	</td>
</tr>   
</form>
<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set ofingerprints = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->