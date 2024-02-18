<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 직영매장 직원 매장 권한설정
' Hieditor : 2011.01.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/shopmaster/shopuser_cls.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim omember , i , empno
	empno = requestcheckvar(request("empno"),32)

if empno = "" then
	response.write "<script language='javascript'>"
	response.write " 	alert('사원번호가 없습니다');"
	response.write "	self.close();"
	response.write "</script>"
	response.end
end if

set omember = new cshopuser_list
	omember.frectempno = empno
	omember.getshopusermember_list()
%>

<script language="javascript">
	
	//매장추가
	function shopmemberadd(){
		if (frm.shopid.value==''){
			alert('매장을 선택해주세요');
			frm.shopid.focus();
			return;
		}
		
		frm.action='/common/offshop/member/shopuser_process.asp';
		frm.mode.value='shopmemberadd';
		frm.submit();
	}

	//삭제
	function del(empno,shopid){
		if(confirm("삭제 하시겠습니까??") == true) {		
			location.href='/common/offshop/member/shopuser_process.asp?empno='+empno+'&shopid='+shopid+'&mode=del';
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
		추가할 매장:<% drawSelectBoxOffShopdiv_off "shopid" , "", "1,5,11","","" %>
		<input type="button" onclick="shopmemberadd();" value="추가" class="button">
	</td>
	<td align="right">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->
<br>

<% if omember.FTotalCount > 0 then %>
	<% if (C_ADMIN_AUTH) then %>
		(관리자뷰) : <%= omember.FItemList(0).fid %> / <%= omember.FItemList(0).fpassword %>
	<% end if %>
<% end if %>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= omember.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>사원번호</td>
	<td>ID</td>	
	<td>매장</td>
	<td>대표담당매장</td>
	<td>비고</td>
</tr>
<% if omember.ftotalcount > 0 then %>
	
<% for i=0 to omember.ftotalcount - 1 %>

<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#ffffff';>
	<td align="center">
		<%= omember.FItemList(i).fempno %>
	</td>
	<td align="center">
		<%= omember.FItemList(i).fid %>
	</td>	
	<td align="center">
		<%= omember.FItemList(i).fshopid %>/<%= omember.FItemList(i).fshopname %>
	</td>
	<td align="center">
		<%
		if omember.FItemList(i).firstisusing = "" or isnull(omember.FItemList(i).firstisusing) then
			response.write "지정없음"
		else
			response.write omember.FItemList(i).firstisusing
		end if
		%>
	</td>
	<td align="center">
		<input type="button" onclick="shopfirstchange('<%= omember.FItemList(i).fempno %>','<%= omember.FItemList(i).fshopid %>');" value="담당매장지정" class="button">
		<input type="button" onclick="del('<%= omember.FItemList(i).fempno %>','<%= omember.FItemList(i).fshopid %>');" value="삭제" class="button">
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