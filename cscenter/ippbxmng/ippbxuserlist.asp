<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터 IPP 사용자연결
' Hieditor : 2015.05.27 이상구 생성
'			 2021.04.09 한용민 수정(아웃소싱 위탁업체 권한 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_ippbxusercls.asp" -->
<%
dim userid, i, occscenterippbxuser

set occscenterippbxuser = new CCSCenterIppbxUser
	occscenterippbxuser.FPageSize = 50
	occscenterippbxuser.FCurrPage = 1
	occscenterippbxuser.GetCSCenterIppbxUserList

%>
<script type='text/javascript'>

function ModifyIppbxInfo(frm){
	if (confirm("수정하시겠습니까?") == true) {
		frm.submit();
	}
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		아이디 : <input type="text" class="text" name="userid" value="<%= userid %>">
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onclick="document.frm.submit()">
	</td>
</tr>
</table>
</form>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		검색결과 : <b><%= occscenterippbxuser.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td >내선번호</td>
	<td >어드민아이디</td>
	<td >사용여부</td>
	<td >수정일</td>
	<td>비고</td>
</tr>
<% if occscenterippbxuser.FTotalCount > 0 then %>
	<% for i = 0 to (occscenterippbxuser.FResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF" height="25">
    	<form name="frm<%= i %>" method="post" action="/cscenter/ippbxmng/ippbxuser_process.asp" style="margin:0px;">
    	<input type="hidden" class="text" name="menupos" value="<%= menupos %>">
    	<input type="hidden" class="text" name="localcallno" value="<%= occscenterippbxuser.FItemList(i).Flocalcallno %>">
        <td><%= occscenterippbxuser.FItemList(i).Flocalcallno %></td>
        <td><input type="text" class="text" name="userid" value="<%= occscenterippbxuser.FItemList(i).Fuserid %>"></td>
        <td>
			<select name="useyn" class="select">
				<option value="Y" <% if (occscenterippbxuser.FItemList(i).Fuseyn = "Y") then %>selected<% end if %>>사용함
				<option value="N" <% if (occscenterippbxuser.FItemList(i).Fuseyn = "N") then %>selected<% end if %>>사용안함
			</select>
        </td>
        <td><%= occscenterippbxuser.FItemList(i).Flastupdate %></td>
        <td>
			<%
			' 위탁업체 팀장 빼고는 위탁업체 일반직원은 수정권한 없음.
			if C_CSUser or C_ADMIN_AUTH then
			%>
				<% if C_CSOutsourcingUser then %>
					<% if C_CSOutsourcingPowerUser then %>
						<input type="button" class="button" value="수정" onClick="ModifyIppbxInfo(frm<%= i %>)">
					<% end if %>
				<% else %>	
					<input type="button" class="button" value="수정" onClick="ModifyIppbxInfo(frm<%= i %>)">
				<% end if %>
			<% end if %>
        </td>
        </form>
    </tr>
	<% next %>
<% else %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="25" colspan="10">검색결과가 없습니다.</td>
    </tr>
<% end if %>
</table>

<%
set occscenterippbxuser = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->