<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  비타민 상세리스트
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
	Dim SearchKey, SearchString, StateDiv, posit_sn, research, orderby  ,sStatus
	dim department_id, inc_subdepartment 
	dim intLoop, arrList ,clsvm
	dim vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, i,sDateType
	
	iCurrPage =requestCheckvar(request("iC"),10)
	SearchKey =requestCheckvar(request("SearchKey"),1)
	SearchString =requestCheckvar(request("SearchString"),60)
	StateDiv =requestCheckvar(request("StateDiv"),1)
	posit_sn =requestCheckvar(request("posit_sn"),10)
	research =requestCheckvar(request("research"),1)
	orderby =requestCheckvar(request("orderby"),1)
	inc_subdepartment =requestCheckvar(request("inc_subdepartment"),1)
	department_id =requestCheckvar(request("department_id"),10)
	sStatus=requestCheckvar(request("selStatus"),1)

	sDateType=requestCheckvar(request("selT"),1)
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(request("sday"),"1")
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	
	iPageSize = 50
	if iCurrPage ="" then iCurrPage =1
	if sDateType ="" then sDateType="1"
	set clsvm	= new Cvitamin
		clsvm.FCurrPage 		= iCurrPage
		clsvm.FPageSize 		= iPageSize		
		clsvm.FRectposit_sn = posit_sn
		clsvm.FRectSearchKey= SearchKey    
		clsvm.FRectSearchString  =SearchString 
		clsvm.Fdepartment_id=   department_id  
		clsvm.Finc_subdepartment =inc_subdepartment
		clsvm.FRectStateDiv = StateDiv 
		clsvm.FRectDateType	 = sDateType
		clsvm.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
		clsvm.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
		clsvm.FRectStatus  = sStatus
		arrList = clsvm.fnGetDetailList
		iTotCnt = clsvm.FTotCnt 
set clsvm	= nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수


%>
<script type="text/javascript">
	 
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

function jsdelVM(){
	var chk = 0;
	var obj = document.frmL.chki;
	 
	if(typeof(obj.length)=="undefined"){ 
		if(obj.checked){
			chk++; 
		}		
	}else{  
		for (var i = 0; i < obj.length; i++){
			if (obj[i].checked){
						chk++;
	 		}
		}
	}
	 
	if (chk==0)
	{ 
		alert("비타민을 하나이상 선택해주세요.");
		return;
	}
		 
	if(confirm("선택한 비타민을 삭제하시겠습니까?")){
		document.frmL.hidM.value="D";
		document.frmL.submit();
	}
}

function jspayVM(){
	var chk = 0;
	var obj = document.frmL.chki;
	 
	if(typeof(obj.length)=="undefined"){ 
		if(obj.checked){
			chk++; 
		}		
	}else{ 
		for (var i = 0; i < obj.length; i++){
			if (obj[i].checked){
						chk++;
	 		}
		}
	}
	 
	if (chk==0)
	{ 
		alert("비타민을 하나이상 선택해주세요.");
		return;
	}
	if(confirm("선택한 비타민을 지급하시겠습니까?\n선택한 비타민중 승인완료된 내역만 지급완료로 변경됩니다.")){
		document.frmL.hidM.value="P";
		document.frmL.submit();
	}
}

 function jsViewEapp(iridx){	  	 
	   var winVME =window.open("/admin/approval/eapp/modeapp.asp?iridx="+iridx,"popVM","width=880, height=600,scrollbars=yes, resizable=yes");
	   winVME.focus();
	 }
</script>
<form name="frmPrc" method="post" action="/admin/member/vitamin/procVitamin.asp">	
	<input type="hidden" name="hidM" value="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
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
			
			&nbsp;
			<select name="selT" class="select">
				<option value="1">신청일</option>
				<option value="2">지급일</option>
			</select>
			<%
					'### 년
					Response.Write "<select name=""syear"" class=""select"">"
					For i=Year(now) To 2017 Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 월
					Response.Write "<select name=""smonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 일
					Response.Write "<select name=""sday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;~&nbsp;"

					'#############################

					'### 년
					Response.Write "<select name=""eyear"" class=""select"">"
					For i=Year(now) To 2017 Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 월
					Response.Write "<select name=""emonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 일
					Response.Write "<select name=""eday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>"

 
				%>
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
		  상태:
		  <select name="selStatus" class="select">
		  	<option value="">전체</option>
		  	<option value="8">신청(품의서 작성전)</option>
		  	<option value="0">승인대기</option>
		  	<option value="1">승인완료</option>
		  	<option value="7">지급완료</option>
		  </select>
		  	&nbsp; 
		  재직여부:
			<select name="StateDiv" class="select">
				<option value="">전체</option>
				<option value="Y">재직</option>
				<option value="N">퇴사</option>
			</select>
			&nbsp;
			<%=printPositOptionIN90("posit_sn", posit_sn)%>
		
			<script language="javascript">
				document.frm.StateDiv.value="<%= StateDiv %>";
				document.frm.SearchKey.value="<%= SearchKey %>"; 
				document.frm.selStatus.value="<%= sStatus %>"; 
				document.frm.selT.value = "<%=sDateType%>";
			</script> 
		</td>
	</tr> 
	</form>
</table>
<!-- 검색 끝 -->


<!-- 액션 시작 -->
<%
'// 로그인정보(등급)에 따라 기본 부서 설정(관리자 이상:1 및 시스템팀:7 경영관리팀:8 제외)
if (session("ssAdminLsn")<=1 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8 or C_PSMngPart or C_ADMIN_AUTH) then
%>

<p>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			[
			관리자 :
				<input type="button" class="button" value="선택 삭제" onClick="javascript:jsdelVM();">
				
				<input type="button" class="button" value="선택 지급" onClick="javascript:jspayVM();" <%if sStatus <>"1" then%>disabled<%end if%>>
			
			]	
			<p>+검색조건에서 상태값을 [승인완료]로 변경시 [선택 지급] 버튼이 활성화됩니다.</p>
		</td> 
	</tr>
</table> 
<% end if %>

<!-- 액션 끝 -->
<p>

<!-- 상단 띠 시작 -->
<form name="frmL" method="post" action="procVitamin.asp">
	<input type="hidden" name="hidM" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			검색결과 : <b><%=iTotCnt%></b>
			&nbsp;
			페이지 : <b><%= iCurrPage %> / <%=iTotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td><input type="checkbox" name="chkAll" onClick="CkeckAll(this)"></td> 
		<td>idx</td>
		<td>사번</td>
		<td>이름</td>
		<td>입사일</td>
		<td>부서</td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td>직급</td><% end if %>
		<td>신청금액</td> 
		<td>신청일</td> 
		<td>지급일</td>		
		<td>상태</td> 
		<td>품의서No</td>
	</tr>
	<% dim isusing
	if isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
			
		%>  
	<tr align="center" bgcolor="#ffffff" height="30">
		<td><input type="checkbox" name="chki" id="chki" value="<%=arrList(0,intLoop)%>" onClick="checkThis(this)" <%if arrList(4,intLoop) =7 then%>disabled<%end if%>></td>
		<td><%=arrList(0,intLoop)%></td>
		<td><%=arrList(5,intLoop)%></td>
		<td><%=arrList(7,intLoop)%></td>
		<td><%=arrList(8,intLoop)%></td>
		<td><%=arrList(9,intLoop)%></td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td><%=arrList(10,intLoop)%></td><% end if %>
		<td align="right"><%=formatnumber(arrList(1,intLoop),0)%></td> 
		<td><%=formatdate(arrList(2,intLoop),"0000-00-00")%></td>
		<td><%=arrList(3,intLoop)%></td>		
		<td><%=fnStatusDesc(arrList(4,intLoop),arrList(11,intLoop))%></td> 
		<td><a href="javascript:jsViewEapp(<%=arrList(11,intLoop)%>);"><%=arrList(11,intLoop)%></a></td>
	</tr>
	<% next %>
	<% else %>
	<tr bgcolor="#ffffff">
		<td colspan="14" align="center">등록된 내역이 존재하지 않습니다.</td>
	</tr>
	<%end if%>
</table>
</form>
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