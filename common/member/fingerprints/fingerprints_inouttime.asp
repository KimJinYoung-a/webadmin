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
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%

if (session("ssAdminPsn") = "10") and (session("ssBctId") <> "boyishP") then
	'// CS팀장님 요청사항, 2015-04-08
	response.write  "권한이 없습니다. - 시스템팀 문의 " ''eastone
	dbget.close() : response.end
end if

Dim ofingerprints,i,page ,part_sn ,sDt ,eDt , dispadmin ,SearchString ,SearchKey
dim department_id, inc_subdepartment
dim IsShowMyDepartmentOnly : IsShowMyDepartmentOnly = False

	sDt = requestCheckVar(request("sDt"),10)
	eDt = requestCheckVar(request("eDt"),10)
	part_sn = requestCheckVar(request("part_sn"),10)
	menupos = requestCheckVar(request("menupos"),10)
	SearchKey = requestCheckvar(Request("SearchKey"),1)
	SearchString = requestCheckvar(Request("SearchString"),32)
	page = requestCheckVar(request("page"),10)
	department_id = requestCheckvar(Request("department_id"),10)
	inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)

	if page = "" then page = 1
	if sDt = "" then sDt = DateSerial(Year(date()), month(date()),1)
	if eDt = "" then eDt = date
	dispadmin = false

'// 점장이나 파트장 이상일경우 등록 권한 부여
if (C_ManagerUpJob or C_ADMIN_AUTH) then
	dispadmin = true
end if

'/관리자일경우 등록 권한 부여
if getlevel_sn("",session("ssBctID")) <= "2" then
	dispadmin = true

'/관리자가 아니며
else
	'/경영관리팀이 아닐경우 파트 박아 넣음
	if Not(C_MngPart or C_PSMngPart) then
		'part_sn = getpart_sn("",session("ssBctID"))

		IsShowMyDepartmentOnly = True
		if (department_id = "") then
			department_id = GetUserDepartmentID("",session("ssBctID"))
		end if

	'/경영관리팀일경우 등록 권한 부여
	else
		dispadmin = true
	end if
end if

set ofingerprints = new cfingerprints_list
	ofingerprints.FPageSize = 20
	ofingerprints.FCurrPage = page
	ofingerprints.FSearchType 	= searchKey
	ofingerprints.FSearchText 	= searchString
	ofingerprints.frectpart_sn = part_sn
	ofingerprints.FrectSDate = sDt
	ofingerprints.FrectEDate = eDt
	ofingerprints.Fdepartment_id 		= department_id
	ofingerprints.Finc_subdepartment 	= inc_subdepartment

	ofingerprints.ffingerprints_list()
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>

	function fingerprintssum(empno,part_sn,sDt,eDt){
		var fingerprintssum = window.open('/common/member/fingerprints/fingerprints_inouttime_sum.asp?empno='+empno+'&part_sn='+part_sn+'&sDt='+sDt+'&eDt='+eDt,'fingerprintssum','width=1024,height=768,scrollbars=yes,resizable=yes');
		fingerprintssum.focus();
	}

	function fingerprintsedit(idx){
		var fingerprintsedit = window.open('/common/member/fingerprints/fingerprints_inouttime_edit.asp?idx='+idx,'fingerprintsedit','width=600,height=400,scrollbars=yes,resizable=yes');
		fingerprintsedit.focus();
	}

	function frmsubmit(page){
		frm.page.value = page;
		frm.submit();
	}

	//코드 등록 & 수정
	//function popPosCodeManage(){
	//    var popPosCodeManage = window.open('/common/member/fingerprints/fingerprints_poscode.asp','popPosCodeManage','width=800,height=600,scrollbars=yes,resizable=yes');
	//    popPosCodeManage.focus();
	//}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%=page%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		부서NEW:
		<% if (IsShowMyDepartmentOnly = True) then %>
			<%= drawSelectBoxMyDepartment(session("ssBctId"), "department_id", department_id) %>
		<% else %>
			<%= drawSelectBoxDepartment("department_id", department_id) %>
		<% end if %>
		<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > 하위 부서직원 제외
		&nbsp;
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
		검색:
		<select name="SearchKey" class="select">
			<option value="">::구분::</option>
			<option value="1" >사번</option>
			<option value="2">이름</option>
		</select>
		<input type="text" class="text" name="SearchString" size="16" value="<%=SearchString%>">
		<script language="javascript">document.frm.SearchKey.value="<%= SearchKey %>";</script>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:frmsubmit('');">
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
		<!--<input type="button" onclick="fingerprintssum('','<%'=part_sn%>','<%'=sDt%>','<%'= eDt %>');" value="통계보기" class="button">-->
	</td>
	<td align="right">
		<% 'if C_ADMIN_AUTH then %>
		<% '<input type="button" value="코드관리" class="button" onClick="popPosCodeManage();"> %>
		<% 'end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ofingerprints.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ofingerprints.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>사원번호</td>
	<td>부서 (포스구분)</td>
	<td>성명</td>
	<td>상태</td>
	<td>출근일</td>
	<td>수정일</td>
	<td>비고</td>
</tr>
<% if ofingerprints.FresultCount>0 then %>
<% for i=0 to ofingerprints.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= ofingerprints.FItemList(i).fidx %>
	</td>
	<td>
		<a href="javascript:fingerprintssum('<%= ofingerprints.FItemList(i).fempno %>','','<%=sDt%>','<%= eDt %>');" onfocus="this.blur();"><%= ofingerprints.FItemList(i).fempno %></a>
	</td>
	<td>
		<a href="javascript:fingerprintssum('','<%= ofingerprints.FItemList(i).fpart_sn %>','<%=sDt%>','<%= eDt %>');" onfocus="this.blur();">
		<%= ofingerprints.FItemList(i).FdepartmentNameFull %> (<%= ofingerprints.FItemList(i).fplaceiname %>)</a>
	</td>
	<td>
		<%= ofingerprints.FItemList(i).fusername %>
	</td>
	<td>
		<%= ofingerprints.FItemList(i).finoutTypeName %>
	</td>
	<td>
		<%= FormatDate(ofingerprints.FItemList(i).finoutTime,"0000-00-00 00:00:00") %>
	</td>
	<td>
		<%= chkIIF(ofingerprints.FItemList(i).flastedituserid<>"",ofingerprints.FItemList(i).flastedituserid & "<br />","") %>
		<%= FormatDate(ofingerprints.FItemList(i).flasteditupdate,"0000-00-00 00:00:00") %>
	</td>
	<td>
		<% if dispadmin then %>
			<input type="button" onclick="fingerprintsedit('<%= ofingerprints.FItemList(i).fidx %>');" value="수정" class="button">
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ofingerprints.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit(<%= ofingerprints.StartScrollPage-1 %>);">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ofingerprints.StartScrollPage to ofingerprints.StartScrollPage + ofingerprints.FScrollCount - 1 %>
			<% if (i > ofingerprints.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ofingerprints.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit(<%= i %>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ofingerprints.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit(<%= i %>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
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
