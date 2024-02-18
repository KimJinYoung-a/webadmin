<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<%
dim clsNonagit, arrList, intLoop
dim iTotCnt
Dim SearchKey, SearchString, StateDiv, posit_sn, research, orderby  
dim department_id, inc_subdepartment 
dim menupos
	SearchKey =requestCheckvar(request("SearchKey"),1)
	SearchString =requestCheckvar(request("SearchString"),60)
	StateDiv =requestCheckvar(request("StateDiv"),1)
	posit_sn =requestCheckvar(request("posit_sn"),10)
	inc_subdepartment =requestCheckvar(request("inc_subdepartment"),1)
	department_id =requestCheckvar(request("department_id"),10)
	
 set clsNonagit = new CAgitPoint
 	 	clsNonagit.FRectposit_sn = posit_sn
		clsNonagit.FRectSearchKey= SearchKey    
		clsNonagit.FRectSearchString  =SearchString 
		clsNonagit.Fdepartment_id=   department_id  
		clsNonagit.Finc_subdepartment =inc_subdepartment
		clsNonagit.FRectStateDiv = StateDiv 
		arrList = clsNonagit.fnGetNonRegagitList
		iTotCnt = clsNonagit.FTotCnt 
 set clsNonagit = nothing
%>
<script type="text/javascript">
	// 사용자 수정/삭제
	function jsModMember(empno)
	{
		var w = window.open("/admin/member/tenbyten/pop_member_modify.asp?menupos=<%=menupos%>&sEPN="+empno,"popMem","width=700,height=600,scrollbars=yes,resizeable=yes");
		w.focus();
	}
	
	
function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function checkThis(comp){
    AnCheckClick(comp)
}

	//선택등록
function jsSetMonthagit(){
	if(confirm("선택된 사번의 포인트를 등록하시겠습니까?")){
		document.frmagit.submit();
	}
}
	</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
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
			재직여부:
			<select name="StateDiv" class="select">
				<option value="">전체</option>
				<option value="Y">재직</option>
				<option value="N">퇴사</option>
			</select>
			&nbsp;
			<%=printPositOptionIN90("posit_sn", posit_sn)%>
			&nbsp;
			검색:
			<select name="SearchKey" class="select">
				<option value="">::구분::</option>
				<option value="1" >아이디</option>
				<option value="2">사용자명</option>
				<option value="3">사번</option>
			</select>
			<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
		  
			<script language="javascript">
				document.frm.StateDiv.value="<%= StateDiv %>";
				document.frm.SearchKey.value="<%= SearchKey %>"; 
			</script> 
		</td>
	</tr> 
	</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<%
'// 로그인정보(등급)에 따라 기본 부서 설정(관리자 이상:1 및 시스템팀:7  )
if (session("ssAdminLsn")<=1 or session("ssAdminPsn")=7  or C_PSMngPart or C_ADMIN_AUTH) then
%>

<p>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left"> 
				<input type="button" class="button" value="선택 포인트등록" onClick="javascript:jsSetMonthagit();"> 
		</td> 
	</tr>
</table> 
<% end if %>

<!-- 액션 끝 -->
<p>
	<form name="frmagit" method="post" action="procAgit.asp">
		<input type="hidden" name="hidM" value="I">
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<input type="hidden" name="addp" value=""> 
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			검색결과 : <b><%=iTotCnt%></b> 
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td><input type="checkbox" name="chkAll" onClick="CkeckAll(this)"></td>
		<td>사번</td>
		<td>이름</td>
		<td>입사일</td>
		<td>부서</td>
		<td>직급</td>
		<td>사용가능기간</td>
		<td>총 포인트</td>	
		 
	</tr>
	<% dim sday, eday, dday, totpoint, monpoint, orgpoint
	IF isArray(arrList) then
		 	sday = date()
			eday = year(date())&"-12-31" 	
			orgpoint = 20 
			 
			For intLoop = 0 To uBound(arrList,2)  
				 totpoint= int((orgpoint/12)*(cint(datediff("m",arrList(3,intLoop),eday))))
				 if day(arrList(3,intLoop))=1 then
				 	totpoint =  int((orgpoint/12)*(cint(datediff("m",arrList(3,intLoop),eday))+1))
					end if
		%> 
	<tr align="center" bgcolor="#ffffff" height="30">
		<td><input type="checkbox" name="chki" value="<%=arrList(0,intLoop)%>" onClick="checkThis(this)"></td>
		<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%=arrList(0,intLoop)%></a></td>
		<td><%=arrList(2,intLoop)%></td>
		<td><%=arrList(3,intLoop)%></td>
		<td><%=arrList(4,intLoop)%></td>
		<td><%=arrList(5,intLoop)%></td>
		<td><%=sday%>~<%=eday%><input type="hidden" name="sDay" value="<%=sday%>"><input type="hidden" name="eDay" value="<%=eday%>"></td>
		<td><%=formatnumber(totpoint,0)%><input type="hidden" name="totp" value="<%=totpoint%>"></td>	
	 
	</tr>
	<%	Next
		END IF%>
</table>
 </form>
<!-- #include virtual="/lib/db/dbclose.asp" -->