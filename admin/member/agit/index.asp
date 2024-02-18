<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  아지트 포인트 리스트
' History : 2017.2.20 정윤정 생성 
'           2018.03.26 허진원 - 선택적 직급 표시
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
	Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
	Dim SearchKey, SearchString, StateDiv, posit_sn, research, orderby  
	dim department_id, inc_subdepartment 
	dim intLoop, arrList ,clsagit
	dim sYYYY
	iCurrPage =requestCheckvar(request("iC"),10)
	SearchKey =requestCheckvar(request("SearchKey"),1)
	SearchString =requestCheckvar(request("SearchString"),60)
	StateDiv =requestCheckvar(request("StateDiv"),1)
	posit_sn =requestCheckvar(request("posit_sn"),10)
	research =requestCheckvar(request("research"),1)
	orderby =requestCheckvar(request("orderby"),1)
	inc_subdepartment =requestCheckvar(request("inc_subdepartment"),1)
	department_id =requestCheckvar(request("department_id"),10)
	sYYYY=requestCheckvar(request("selY"),4)
	if sYYYY="" then sYYYY = year(date())
	iPageSize = 50
	if iCurrPage ="" then iCurrPage =1
	set clsagit	= new CAgitPoint
		clsagit.FCurrPage 		= iCurrPage
		clsagit.FPageSize 		= iPageSize		
		clsagit.FRectposit_sn = posit_sn
		clsagit.FRectSearchKey= SearchKey    
		clsagit.FRectSearchString  =SearchString 
		clsagit.Fdepartment_id=   department_id  
		clsagit.Finc_subdepartment =inc_subdepartment
		clsagit.FRectStateDiv = StateDiv 
		clsagit.FRectYYYY = sYYYY
		arrList = clsagit.fnAgitGetList
		iTotCnt = clsagit.FTotCnt 
set clsagit	= nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수


%>
<script type="text/javascript">
	//전체 등록
	function jsSetYearPoint(){
	 	if (confirm("내년도 아지트 이용 포인트가 생성됩니다. 전체 포인트를 등록하시겠습니까?") ) { 
		document.frmPrc.submit();
	}
	}
	
	//미등록자 등록
	function jsSetMonthPoint(){
		var winP = window.open("popRegAgit.asp","popP","width=1000, height=800,scrollbars=yes,resizable=yes");
		winP.focus;
	}
	
		// 사용자 수정/삭제
	function jsModMember(empno)
	{
		var w = window.open("/admin/member/tenbyten/pop_member_modify.asp?menupos=<%=menupos%>&sEPN="+empno,"popMem","width=700,height=600,scrollbars=yes,resizeable=yes");
		w.focus();
	}
	
	function jsDetail(empno,syyyy,smm, eyyyy,emm){
		var w = window.open("/admin/member/agit/uselist.asp?menupos=<%=menupos%>&SearchKey=3&SearchString="+empno+"&selSY="+syyyy+"&selSM="+smm+"&selEY="+eyyyy+"&selEM="+emm,"popAgit","");
		w.focus();
	}
</script>
<form name="frmPrc" method="post" action="/admin/member/Agit/procAgit.asp">	
	<input type="hidden" name="hidM" value="A">
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="4" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			부서NEW:
			<%= drawSelectBoxDepartmentALL("department_id", department_id) %>
			<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > 하위 부서직원 제외
		</td>

		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr> 
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left"> 
			검색:
			<select name="SearchKey" class="select">
				<option value="">::구분::</option>
				<option value="1" >아이디</option>
				<option value="2">사용자명</option>
				<option value="3">사번</option>
			</select>
			<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
				&nbsp;
		  	재직여부:
			<select name="StateDiv" class="select">
				<option value="">전체</option>
				<option value="Y">재직</option>
				<option value="N">퇴사</option>
			</select>
			<% if C_PSMngPart or C_ADMIN_AUTH then %>
			&nbsp;
			<%=printPositOptionIN90("posit_sn", posit_sn)%>
			<% end if %>
		&nbsp;기간:
		<%dim i 
		%>
		<select name="selY" class="select">
			<%for i=year(dateadd("yyyy",1,date())) to 2017 step-1%>
			<option value="<%=i%>"><%=i%></option>
			<%next%>
		</select>
			<script language="javascript">
				document.frm.StateDiv.value="<%= StateDiv %>";
				document.frm.SearchKey.value="<%= SearchKey %>"; 
				document.frm.selY.value ="<%=sYYYY%>";
			</script> 
		</td>
	</tr>	
</table>
</form>
<!-- 검색 끝 -->


<!-- 액션 시작 -->
<%
'// 로그인정보(등급)에 따라 기본 부서 설정(관리자 이상:1 및 시스템팀:7 경영관리팀:8 제외)
if (session("ssAdminLsn")<=1 or session("ssAdminPsn")=7 or  C_PSMngPart or C_ADMIN_AUTH) then
%>

<p>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			[
			관리자 :
				<input type="button" class="button" value="포인트등록" onClick="javascript:jsSetMonthPoint();">
				<input type="button" class="button" value="전체포인트등록(년1회)" onClick="javascript:jsSetYearPoint()">
			]	
		</td> 
	</tr>
</table> 
<% end if %>

<!-- 액션 끝 -->
<p>

<!-- 상단 띠 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			검색결과 : <b><%=iTotCnt%></b>
			&nbsp;
			페이지 : <b><%= iCurrPage %> / <%=iTotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>idx</td>
		<td>사번</td>
		<td>ID</td>
		<td>이름</td>
		<td>입사일</td>
		<td>부서</td>
		<% if C_PSMngPart or C_ADMIN_AUTH then %><td>직급</td><% end if %>
		<td>사용가능기간</td>
		<td>총 포인트</td>
		<td>사용 포인트</td>
		<td>잔여 포인트</td>		
		<td>재직여부</td>
		<td>사용가능</td>
		<td>등록자</td> 
	</tr>
	<% dim isusing, ndate
	if isArray(arrList) THEN
		ndate = Cstr(date())
			For intLoop = 0 To UBound(arrList,2)
			IF arrList(8,intLoop)>=ndate then '사용가능여부
				isusing ="Y"
			ELSE
				isusing ="N"
			END IF	
		%>  
	<tr bgcolor=<%if isusing="Y" then%>"#ffffff"<%else%>"#EFEFEF"<%END IF%> height="30">
		<td align="center"><%=arrList(0,intLoop)%></td>
		<td align="center"><a href="javascript:jsModMember('<%=arrList(1,intLoop)%>')"><%=arrList(1,intLoop)%></a></td>
		<td align="center"><%=arrList(2,intLoop)%></td>
		<td align="center"><%=arrList(3,intLoop)%></td>
		<td align="center"><%=arrList(4,intLoop)%></td>
		<td><%=arrList(5,intLoop)%></td>
		<% if C_PSMngPart or C_ADMIN_AUTH then %><td align="center"><%=arrList(6,intLoop)%></td><% end if %>
		<td align="center"><%=arrList(7,intLoop)%>~<%=arrList(8,intLoop)%></td>
		<td align="center"><%=formatnumber(arrList(9,intLoop),0)%></td>
		<td align="center"><a href="javascript:jsDetail('<%=arrList(1,intLoop)%>','<%=year(arrList(7,intLoop))%>','<%=month(arrList(7,intLoop))%>','<%=year(arrList(8,intLoop))%>','<%=month(arrList(8,intLoop))%>');"><%=formatnumber(arrList(10,intLoop),0)%></a></td> 
		<td align="center"><%=formatnumber(arrList(9,intLoop)-arrList(10,intLoop),0)%></td> 
		<td align="center"><%=arrList(11,intLoop)%></td>
		<td align="center"><%=isusing%></td>
		<td align="center"><%=arrList(12,intLoop)%></td>  		
		 
	</tr>
	<% next %>
	<% else %>
	<tr bgcolor="#ffffff">
		<td colspan="14" align="center">등록된 내역이 존재하지 않습니다.</td>
	</tr>
	<%end if%>
</table>
<!-- 페이징처리 --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
 <!-- #include virtual="/lib/db/dbclose.asp" -->