<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 어드민 USB 인증
' History : 2008.09.25 한용민 생성 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/admin_keyclass.asp" -->

<script language="javascript">

	function getsubmit(){

	if(!frm_edit.key_idx.value){
		alert("인증kEY를 입력해주세요");
		frm_edit.key_idx.focus();
		return false;
	}

		frm_edit.mode.value = 'new';	
		frm_edit.submit();
	}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_edit" action="/admin/member/admin_keyprocess.asp" method="get">
	<input type="hidden" name="mode">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">		
		<td align="center">인증KEY</td>
		<td align="center">Team</td>	
		<td align="center">사용자</td>	
		<td align="center">상세사용자</td>		
		<td align="center">사용여부</td>
    </tr>
	<tr align="center" bgcolor="ffffff">		
		<td align="center"><input type="text" size=30 name="key_idx"></td>
		<td align="center">
			<select name="teamname">
				<option value="">선택</option>
				<option value="CEO" >CEO</option>
				<option value="SYSTEM">SYSTEM</option>
				<option value="ONLINE">ONLINE</option>
				<option value="MARKETING">MARKETING</option>
				<option value="MD">MD</option>
				<option value="WD">WD</option>
				<option value="물류">물류</option>
				<option value="OFFLINE">OFFLINE</option>
				<option value="CS">CS</option>
				<option value="ITHINKSO">ITHINKSO</option>														
				<option value="경영">경영</option>
				<option value="FINGERS">FINGERS</option>
				<option value="패션">패션사업팀</option>				
			</select>		
		</td>	
		<td align="center"><input type="text" name="username"></td>	
		<td align="center"><input type="text" name="username_detail"></td>		
		<td align="center">
			<select name="del_isusing">
				<option value="Y">사용</option>
				<option value="N">삭제</option>
			</select>		
		</td>
    </tr>  
</form>
	<tr align="center" bgcolor="ffffff">		
		<td align="left" colspan=5><input type="button" class="button" value="저장" onclick="getsubmit();"></td>
    </tr>	      
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
