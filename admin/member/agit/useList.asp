<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  아지트 신청내역 리스트
' History : 2017.2.20 정윤정 생성 
'           2018.03.26 허진원 속초 추가/ 직급 표시 선택 제거
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
	dim sSYYYY,sSMM,sEYYYY,sEMM
	dim blnipkum, blnreturnkey, blnusing, blnRefund,chkTerm,AreaDiv
	
	iCurrPage =requestCheckvar(request("iC"),10)
	SearchKey =requestCheckvar(request("SearchKey"),1)
	SearchString =requestCheckvar(request("SearchString"),60)
	StateDiv =requestCheckvar(request("StateDiv"),1)
	posit_sn =requestCheckvar(request("posit_sn"),10)
	research =requestCheckvar(request("research"),1)
	orderby =requestCheckvar(request("orderby"),1)
	inc_subdepartment =requestCheckvar(request("inc_subdepartment"),1)
	department_id =requestCheckvar(request("department_id"),10)
	sSYYYY=requestCheckvar(request("selSY"),4)
	sSMM=requestCheckvar(request("selSM"),2)
	sEYYYY=requestCheckvar(request("selEY"),4)
	sEMM=requestCheckvar(request("selEM"),2)
	blnipkum = requestCheckvar(request("selipkum"),1)
	blnreturnkey = requestCheckvar(request("selRK"),1)
	blnusing = requestCheckvar(request("selUse"),1)
	blnRefund = requestCheckvar(request("selre"),1)
	chkTerm    =requestCheckvar(request("chkTerm"),3)
	 AreaDiv =requestCheckvar(request("AreaDiv"),1) 
	 
	if sSYYYY="" then sSYYYY = year(date())
	if sSMM="" then sSMM = month(date())
	if sEYYYY="" then sEYYYY = year(date())
	if sEMM="" then sEMM = month(date())
	iPageSize = 50
	if iCurrPage ="" then iCurrPage =1
		
	if research =""	 then
		 chkTerm ="on" 
		 blnusing =	"Y"
	end if
'	
'	dim strParm
' strParm = "ic="&iCurrPage&"&department_id="&department_id&"&inc_subdepartment="&inc_subdepartment&"&SearchKey="&SearchKey&"&SearchString="&SearchString&"&StateDiv="&StateDiv
' strParm = strParm&"&posit_sn="&posit_sn&"&selSY="&selSY&"&selSM="&selSM&"&selEY="&selEY&"&selEM="&selEM&"&AreaDiv="&AreaDiv&""
 
	set clsagit	= new CAgitUse
		clsagit.FCurrPage 		= iCurrPage
		clsagit.FPageSize 		= iPageSize		
		clsagit.FRectposit_sn = posit_sn
		clsagit.FRectSearchKey= SearchKey    
		clsagit.FRectSearchString  =SearchString 
		clsagit.Fdepartment_id=   department_id  
		clsagit.Finc_subdepartment =inc_subdepartment
		clsagit.FRectStateDiv = StateDiv 
		clsagit.FRectAreadiv = AreaDiv
		clsagit.FRectSYYYYMM = sSYYYY&"-"&Format00(2,sSMM)
		clsagit.FRectEYYYYMM = sEYYYY&"-"&Format00(2,sEMM)
		clsagit.FRectIpkum 			= blnipkum
		clsagit.FRectreturnkey 	= blnreturnkey
		clsagit.FRectUsing 			= blnusing
		clsagit.FRectRefund 		= blnRefund
		clsagit.FRectChkTerm    = chkTerm
		arrList = clsagit.FnAgitUseList
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

	// 아지트 안내 문자 관리
	function jsModInfoSMS() {
		var w = window.open("popAgitInfoSms.asp","popAgtSms","width=500,height=500,scrollbars=yes,resizeable=yes");
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
 
 function jsChkIdx(ival){
 	if (typeof(document.frmList.chki.length)=="undefined"){
 		document.frmList.chki.checked=true;
 	}else{
 		document.frmList.chki[ival].checked=true;
	}
 }
 
 function jsModBook(){
 	if(confirm("선택내용을 저장하시겠습니까?")){
 		document.frmList.target="ifmProc";
 		document.frmList.submit();
 	}
 }
 
 function jsNewBook(){
 	var p = window.open("/admin/member/tenbyten/agit/pop_tenbyten_Agit_Edit_admin.asp","popNAgit","width=700,height=700,scrollbars=yes,resize=yes");
		p.focus();
 }
 
 function jsViewBook(idx){
 	var p = window.open("/admin/member/tenbyten/agit/pop_tenbyten_Agit_Edit_admin.asp?idx="+idx,"popNAgit","width=700,height=700,scrollbars=yes,resize=yes");
		p.focus();
 }
 
 
//아지트 패널티 관리 오픈
function jsPopPenalty(){
	var winP = window.open("popAgitPenaltyList.asp","popP","width=1000, height=800,scrollbars=yes,resizable=yes");
	winP.focus;
}
</script>
<iframe id="ifmProc" name="ifmProc" src="about:blank" width="0" height="0" frameborder="0"></iframe>

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
			<input type="button" class="button_s" value="검색" onClick="javascript:	document.frmList.target=self;document.frm.submit();">
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
		&nbsp;
		<input type="checkbox" name="chkTerm" <%if chkTerm ="on" then%>checked<%end if%>>
		이용기간:
		<%dim i%> 
		<select name="selSY" class="select">
			<%for i=year(dateadd("yyyy",1,date())) to 2017 step-1%>
			<option value="<%=i%>"><%=i%></option>
			<%next%>
		</select>
		<select name="selSM" class="select">
			<%for i=1 to 12%>
			<option value="<%=i%>"><%=i%></option>
			<%next%>
		</select>
		~ 
		<select name="selEY" class="select">
			<%for i=year(dateadd("yyyy",1,date())) to 2017 step-1%>
			<option value="<%=i%>"><%=i%></option>
			<%next%>
		</select>
		<select name="selEM" class="select">
			<%for i=1 to 12%>
			<option value="<%=i%>"><%=i%></option>
			<%next%>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
		<td align="left"> 
			아지트 : 
			<select name="AreaDiv" class="select">
				<option value="">전체</option>
				<option value="1">제주도</option>
				<!--<option value="2">양평</option>-->
				<option value="3">속초</option>
			</select>
		&nbsp;입금여부:
		<select name="selipkum" class="select">
			<option value="">전체</option>
			<option value="1">Y</option>
			<option value="0">N</option>
		</select>
		&nbsp;키반납여부:
		<select name="selRK" class="select">
			<option value="">전체</option>
			<option value="1">Y</option>
			<option value="0">N</option>
		</select>
		&nbsp;신청상태:
		<select name="selUse" class="select">
			<option value="">전체</option>
			<option value="Y">Y</option>
			<option value="N">N</option>
		</select>
		&nbsp;환불여부:
		<select name="selre" class="select">
			<option value="">전체</option>
			<option value="Y">Y</option>
			<option value="N">N</option>
		</select>
			<script language="javascript">
				document.frm.StateDiv.value="<%= StateDiv %>";
				document.frm.SearchKey.value="<%= SearchKey %>"; 
				document.frm.selSY.value ="<%=sSYYYY%>";
				document.frm.selSM.value ="<%=sSMM%>";
				document.frm.selEY.value ="<%=sEYYYY%>";
				document.frm.selEM.value ="<%=sEMM%>";
				document.frm.selipkum.value ="<%=blnipkum%>";
				document.frm.selRK.value ="<%=blnreturnkey%>";
				document.frm.selUse.value ="<%=blnusing%>";
				document.frm.selre.value ="<%=blnRefund%>";
				document.frm.AreaDiv.value ="<%=areadiv%>";
			</script> 
		</td>
	</tr>	
</table>
</form>
<!-- 검색 끝 -->


<!-- 액션 시작 -->
 
<p>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<% if C_PSMngPart or C_ADMIN_AUTH then %><input type="button" class="button"  value="아지트 안내문자 관리" onClick="jsModInfoSMS();"><% end if %>
		</td>
		<td align="right">
			<input type="button" class="button"  value="선택 내용 저장" onClick="jsModBook();">
			<input type="button" class="button"  value="관리자 신규등록" onClick="jsNewBook();">
			<input type="button" class="button" value="패널티 관리" onClick="javascript:jsPopPenalty()">
		</td> 
	</tr>
</table> 
 

<!-- 액션 끝 -->
<p>

<!-- 상단 띠 시작 -->
<form name="frmList" method="post" action="/admin/member/Agit/procAgit.asp">
	<input type="hidden" name="hidM" value="M">
	<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="19">
			검색결과 : <b><%=iTotCnt%></b>
			&nbsp;
			페이지 : <b><%= iCurrPage %> / <%=iTotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td><input type="checkbox" name="chkAll" onClick="CkeckAll(this)"></td>
		<td>idx</td>
		<td>사번</td>
		<td>ID</td>
		<td>이름</td>
		<td>입사일</td>
		<td>부서</td>
		<% if C_PSMngPart or C_ADMIN_AUTH then %><td>직급</td><% end if %>
		<td>재직</td>
		<td>아지트</td>
		<td>이용기간</td>
		<td>이용인원</td>
		<td>이용포인트</td>
		<td>이용금액</td>		
		<td>입금</td>
		<td>키반납</td>
		<td>신청상태</td>  
		<td>등록일</td> 
		<td>패널티</td> 
	</tr> 
	<% dim isusing, ndate
	if isArray(arrList) THEN
		ndate = Cstr(date())
			For intLoop = 0 To UBound(arrList,2)
			 
		%>  
	<tr bgcolor=<%if arrList(18,intLoop) ="Y" then%>"#ffffff"<%else%>"#EFEFEF"<%END IF%> height="30">
		<td><input type="checkbox" name="chki"   value="<%=arrList(0,intLoop)%>" onClick="checkThis(this)"></td>
		<td align="center"><a href="javascript:jsViewBook('<%=arrList(0,intLoop)%>');"><%=arrList(0,intLoop)%></a></td>
		<td align="center"><%=arrList(1,intLoop)%></td>
		<td align="center"><%=arrList(2,intLoop)%></td>
		<td align="center"><%=arrList(3,intLoop)%></td>
		<td align="center"><%=arrList(4,intLoop)%></td>
		<td><%=arrList(5,intLoop)%></td>
		<% if C_PSMngPart or C_ADMIN_AUTH then %><td align="center"><%=arrList(6,intLoop)%></td><% end if %>
		<td align="center"><%=arrList(7,intLoop)%></td>
		<td align="center"><% Select Case arrList(8,intLoop): Case "1" %>제주도<%: Case "2" %>양평<%: Case "3" %>속초<%:end Select %></td>
		<td align="center"><%=formatdate(arrList(9,intLoop),"0000.00.00-00:00") %> (<%=FnWeekName(DatePart("w", arrList(9,intLoop)))%>)~<%=formatdate(arrList(10,intLoop),"0000.00.00-00:00") %>(<%=FnWeekName(DatePart("w", arrList(10,intLoop)))%>)</td>
		<td align="center"><%=arrList(11,intLoop)%></td>
		<td align="center"><%=arrList(12,intLoop)%></td>
		<td align="center"><%=formatnumber(arrList(14,intLoop),0)%></td>  
		<td align="center"> 
			<input type="radio" name="rdoin<%=arrList(0,intLoop)%>" value="0" <%if arrList(15,intLoop) =0 then%>checked<%end if%> onClick="jsChkIdx(<%=intLoop%>);"><font color="blue">입금전</font> 
			<input type="radio" name="rdoin<%=arrList(0,intLoop)%>" value="1" <%if arrList(15,intLoop) = 1 then%>checked<%end if%> onClick="jsChkIdx(<%=intLoop%>);"><font color="red">입금완료</font>			
			<input type="radio" name="rdoin<%=arrList(0,intLoop)%>" value="9" <%if arrList(15,intLoop) = 9 then%>checked<%end if%> onClick="jsChkIdx(<%=intLoop%>);"><font color="gray">환불</font>			
		</td>
		<td align="center"> 
			<input type="radio" name="rdorek<%=arrList(0,intLoop)%>" value="1" <%if arrList(17,intLoop) then%>checked<%end if%> onClick="jsChkIdx(<%=intLoop%>);"><font color="blue">Y</font> 
			<input type="radio" name="rdorek<%=arrList(0,intLoop)%>" value="0" <%if not arrList(17,intLoop)  then%>checked<%end if%> onClick="jsChkIdx(<%=intLoop%>);"><font color="red">N</font>			
	 </td>
		<td align="center"><%if arrList(18,intLoop) ="Y" then%><font color="blue">Y</font><%else%><font color="red">N</font><%end if%></td>
	 
		<td align="center"><%=formatdate(arrList(24,intLoop),"0000-00-00") %></td>
		<td align="center"><%if arrList(21,intLoop)>0 then%>
			<%=arrList(22,intLoop)%>~<%=arrList(23,intLoop)%>
			<%end if%>
			
			</td>  		
		 
	</tr>
	<% next %>
	<% else %>
	<tr bgcolor="#ffffff">
		<td colspan="20" align="center">등록된 내역이 존재하지 않습니다.</td>
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
 