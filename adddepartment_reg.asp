<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 관리부서 추가
' Hieditor : 2017.08.22 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAddDepCls.asp"-->
<%
dim omember , i , empno
dim clsAddDep, arrList, intLoop
	empno = requestcheckvar(request("empno"),32)

if empno = "" then
	response.write "<script language='javascript'>"
	response.write " 	alert('사원번호가 없습니다');"
	response.write "	self.close();"
	response.write "</script>"
	response.end
end if

set clsAddDep = new CAddDep
  clsAddDep.Fempno = empno
  arrList = clsAddDep.fnGetAddDepList
set clsAddDep = nothing 
 
%>

<script language="javascript">
	
	//부서추가
	function jsAdddep(){
		if (frm.department_id.value==''){
			alert('부서를 선택해주세요');
			frm.department_id.focus();
			return;
		}
		
		frm.action='/common/offshop/member/adddepartment_process.asp';
		frm.mode.value='A';
		frm.submit();
	}

	//삭제
	function del(empno,shopid){
		if(confirm("삭제 하시겠습니까??") == true) {		
			location.href='/common/offshop/member/adddepartment_process.asp?empno='+empno+'&shopid='+shopid+'&mode=del';
		} else {
			return;
		}	
	}
	
	//담당매장지정
	function shopfirstchange(empno,shopid){
		if(confirm("선택하신 매장을 대표담당매장으로 변경 하시겠습니까??") == true) {		
			location.href='/common/offshop/member/shopuser_process.asp?empno='+empno+'&shopid='+shopid+'&mode=shopfirstchange';
		} else {
			return;
		}	
	}
		
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="post">
<input type="hidden" name="mode">
<input type="hidden" name="empno" value="<%=empno%>">
<tr>
	<td align="left">
		추가할 부서:<%= drawSelectBoxDepartment("department_id", "") %>
		<input type="button" onclick="jsAdddep();" value="추가" class="button">
	</td>
	<td align="right">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->
<br>
 

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= omember.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>사번</td>
	<td>아이디</td>	
	<td>부서</td> 
	<td>비고</td>
</tr>
<% if isArray(arrList) then %>
	
<% for intLoop=0 to arrList(2,intLoop)  %>

<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background="#ffffff";>
	<td align="center">
		<%= arrList(0,intLoop) %>
	</td>
	<td align="center">
		<%= arrList(1,intLoop) %>
	</td>	
	<td align="center">
				<%= arrList(3,intLoop) %>
	</td> 
	<td align="center">
		<input type="button" onclick=" " value="담당매장지정" class="button">
		<input type="button" onclick=" ;" value="삭제" class="button">
	</td>	
</tr>   
<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>


<%
set omember = nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->