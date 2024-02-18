<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 부서권한 추가
' Hieditor : 2017.08.23 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAddDepCls.asp"-->
<%
dim  empno
dim clsadddep, arrList ,intLoop
	empno = requestcheckvar(request("sEPN"),32)
 

if empno = "" then
	response.write "<script language='javascript'>"
	response.write " 	alert('사원번호가 없습니다');"
'	response.write "	self.close();"
	response.write "</script>"
'	response.end
end if

set clsadddep  = new CAddDep
 clsadddep.Fempno = empno
 arrList = clsadddep.fnGetAddDepList
set clsadddep =  nothing	
	
%>

<script language="javascript">
	
	//매장추가
	function jsAdddep(){
		if (frm.department_id.value==''){
			alert('부서를 선택해주세요');
			frm.department_id.focus();
			return;
		}
		
		frm.action='adddep_proc.asp';
		frm.mode.value='A';
		frm.submit();
	}

	//삭제
	function del(empno,depid){
		if(confirm("삭제 하시겠습니까??") == true) {		
		frm.action='adddep_proc.asp';
		frm.mode.value='D';
		frm.depid.value = depid;
		frm.submit();
		} else {
			return;
		}	
	}
	  
		
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="post">
<input type="hidden" name="mode">
<input type="hidden" name="sEPN" value="<%=empno%>">
<input type="hidden" name="depid" value="">
<tr>
	<td align="left">
		추가 부서NEW:
			<%= drawSelectBoxDepartmentALL("department_id", "") %>
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
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>사번</td>
	<td>아이디</td>	
	<td>부서</td>
	<td>비고</td>
</tr>
<% if  isArray(arrList) then %>
	
<% for intLoop=0 to  ubound(arrList,2) %>

<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background="#FFFFFF";>
	<td align="center">
		<%= arrList(0,intLoop) %>
	</td>
	<td align="center">
		<%= arrList(1,intLoop)%>
	</td>	
	<td align="center">
		<%= arrList(3,intLoop) %>
	</td> 
	<td align="center">
		<input type="button" onclick="del('<%=arrList(0,intLoop)%>','<%=arrList(2,intLoop)%>');" value="삭제" class="button">
	</td>	
</tr>   
<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

 

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->