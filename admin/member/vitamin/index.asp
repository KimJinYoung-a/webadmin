<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  비타민 리스트
' History : 2017.2.20 정윤정 생성 
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
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVitaminCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
	Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
	Dim SearchKey, SearchString, StateDiv, posit_sn, research, orderby  
	dim department_id, inc_subdepartment 
	dim intLoop, arrList ,clsvm
	dim sYYYY
	
	iCurrPage =requestCheckvar(request("iC"),10)
	SearchKey =requestCheckvar(request("SearchKey"),1)
	SearchString =requestCheckvar(request("SearchString"),60)
	StateDiv =requestCheckvar(request("StateDiv"),1)
	posit_sn =requestCheckvar(request("posit_sn"),10)
	research =requestCheckvar(request("research"),1)
	orderby =requestCheckvar(request("orderby"),2)
	inc_subdepartment =requestCheckvar(request("inc_subdepartment"),1)
	department_id =requestCheckvar(request("department_id"),10)
	sYYYY=requestCheckvar(request("selY"),4)
	iPageSize=requestCheckvar(request("selPS"),10)
	
	if sYYYY="" then sYYYY = year(date())
	if orderby="" then orderby = "CD"	
	
	if (iPageSize = "") then
			iPageSize = 50
	end if 
	
	if iCurrPage ="" then iCurrPage =1
	set clsvm	= new Cvitamin
		clsvm.FCurrPage 		= iCurrPage
		clsvm.FPageSize 		= iPageSize		
		clsvm.FRectposit_sn = posit_sn
		clsvm.FRectSearchKey= SearchKey    
		clsvm.FRectSearchString  =SearchString 
		clsvm.Fdepartment_id=   department_id  
		clsvm.Finc_subdepartment =inc_subdepartment
		clsvm.FRectStateDiv = StateDiv 
		clsvm.FRectYYYY = sYYYY
		clsvm.FRectOrderby = orderby
		arrList = clsvm.fnvitaminGetList
		iTotCnt = clsvm.FTotCnt 
set clsvm	= nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수


%>
<script type="text/javascript">
	//전체 등록
	function jsSetYearVM(){
	 	if (confirm("내년도 비타민이 생성됩니다. 전체 비타민을 등록하시겠습니까?") == true) { 
		document.frmPrc.submit();
	}
	}
	
	//미등록자 등록
	function jsSetMonthVM(){
		var winVM = window.open("popRegVitamin.asp?menupos=<%=menupos%>","popVM","width=1000, height=800,scrollbars=yes,resizable=yes");
		winVM.focus;
	}
	
		// 사용자 수정/삭제
	function jsModMember(empno)
	{
		var w = window.open("/admin/member/tenbyten/pop_member_modify.asp?menupos=<%=menupos%>&sEPN="+empno,"popMem","width=700,height=600,scrollbars=yes,resizeable=yes");
		w.focus();
	}
	
	//리스트 정렬
	 function jsSort(sValue,i){  
	  
	 	document.frm.orderby.value= sValue; 
	 	 
		   if (-1 < eval("document.all.img"+i).src.indexOf("_alpha")){
	        document.frm.orderby.value= sValue+"D";  
	    }else if (-1 < eval("document.all.img"+i).src.indexOf("_bot")){
	     		document.frm.orderby.value= sValue+"A";  
	    }else{
	       document.frm.orderby.value= sValue+"D";  
	    } 
	    
	   
		 document.frm.submit();
	}

	// 비타민 수정
	function jsModVitamin(idx) {
		var w = window.open("popModifyVitamin.asp?idx="+idx+"&menupos=<%=menupos%>","popVitamin","width=500,height=200,scrollbars=yes,resizeable=yes");
		w.focus();
	}
</script>
<form name="frmPrc" method="post" action="/admin/member/vitamin/procVitamin.asp">	
	<input type="hidden" name="hidM" value="A">
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="orderby" value="<%=orderby%>">
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
			&nbsp;
			<%=printPositOptionIN90("posit_sn", posit_sn)%>
			&nbsp;기간:
		<%dim i%>
		<select name="selY" class="select">
			<%for i=year(date()) to 2017 step-1%>
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
	<tr>
		<td bgcolor="#FFFFFF" >	표시갯수:
			<select class="select" name="selPS"> 
				<option value="50" <% if (iPageSize = "50") then %>selected<% end if %> >50 개</option>
				<option value="100" <% if (iPageSize = "100") then %>selected<% end if %> >100 개</option>
				<option value="500" <% if (iPageSize = "500") then %>selected<% end if %> >500 개</option>
			</select></td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->


<!-- 액션 시작 -->
<%
'// 로그인정보(등급)에 따라 기본 부서 설정(관리자 이상:1 및 시스템팀:7 C_PSMngPart:인사팀 외 개발자)
if (session("ssAdminLsn")<=1 or session("ssAdminPsn")=7 or  C_PSMngPart or C_ADMIN_AUTH) then
%>

<p>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			[
			관리자 :
				<input type="button" class="button" value="비타민등록" onClick="javascript:jsSetMonthVM();">
				<input type="button" class="button" value="전체비타민등록(년1회)" onClick="javascript:jsSetYearVM()">
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
		<td onClick="javascript:jsSort('C','1');" style="cursor:pointer;"><b>사번</b> <img src="/images/list_lineup<%IF orderby="CD" THEN%>_bot<%ELSEIF orderby="CA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
		<td onClick="javascript:jsSort('N','2');" style="cursor:pointer;"><b>이름</b> <img src="/images/list_lineup<%IF orderby="ND" THEN%>_bot<%ELSEIF orderby="NA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
		<td>입사일</td>
		<td>부서</td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td>직급</td><% end if %>
		<td>사용가능기간</td>
		<td>총 비타민</td>
		<td>사용비타민</td>
		<td>잔여 비타민</td>		
		<td>재직여부</td>
		<td>사용가능</td>
		<td>등록자</td>
		<td></td>
	</tr>
	<% dim isusing
	if isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
			IF arrList(7,intLoop) <= date() and arrList(8,intLoop)>=date() then '사용여부
				isusing ="Y"
			ELSE
				isusing ="N"
			END IF	
		%>  
	<tr bgcolor=<%if isusing="Y" then%>"#ffffff"<%else%>"#EFEFEF"<%END IF%> height="30">
		<td align="center"><%=arrList(0,intLoop)%></td>
		<td align="center"><a href="javascript:jsModMember('<%=arrList(1,intLoop)%>')"><%=arrList(1,intLoop)%></a></td>
		<td align="center"><%=arrList(3,intLoop)%></td>
		<td align="center"><%=arrList(4,intLoop)%></td>
		<td><%=arrList(5,intLoop)%></td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td align="center"><%=arrList(6,intLoop)%></td><% end if %>
		<td align="center"><%=formatdate(arrList(7,intLoop),"0000-00-00")%>~<%=formatdate(arrList(8,intLoop),"0000-00-00")%></td>
		<td align="right"><a href="javascript:jsModVitamin('<%=arrList(0,intLoop)%>')"><%=formatnumber(arrList(9,intLoop),0)%></a></td>
		<td align="right"><%=formatnumber(arrList(10,intLoop),0)%></td> 
		<td align="right"><%=formatnumber(arrList(9,intLoop)-arrList(10,intLoop),0)%></td> 
		<td align="center"><%=arrList(11,intLoop)%></td>
		<td align="center"><%=isusing%></td>
		<td align="center"><%=arrList(12,intLoop)%></td>  		
		<td></td>
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