<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.28 한용민 생성
'	Description : 오거나이저
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<script language="javascript">

	function getsubmit(){

	if(!frm_edit.isusing.value){
		alert("사용여부를 선택해주세요");
		frm_edit.isusing.focus();
		return false;
	}

	if(!frm_edit.itemid.value){
		alert("상품코드 입력 해주세요");
		frm_edit.itemid.focus();
		return false;
	}
		
		frm_edit.mode.value = 'new';	
		frm_edit.mode_type.value = 'keyword';
		frm_edit.submit();
	}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_edit" action="/admin/organizer/alpha_process.asp" method="get">
	<input type="hidden" name="mode">
	<input type="hidden" name="mode_type">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">		
		<td align="center">상품코드</td>
		<td align="center">사용여부</td>
    </tr>
	<tr align="center" bgcolor="ffffff">		
				<td align="center">
					<input type="text" size=30 name="itemid" >
				</td>	
				<td align="center">
					<select name="isusing" >
						<option value="" >선택</option>
						<option value="Y" >Y</option>
						<option value="N" >N</option>
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
