<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.09 한용민 생성
'	Description : 다이어리스토리
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/admin_keyclass.asp" -->

<script language="javascript">

	function getsubmit(){

	if(!frm_edit.option_value.value){
		alert("옵션명을 입력해주세요");
		frm_edit.key_idx.focus();
		return false;
	}

	if(!frm_edit.type.value){
		alert("타입을 선택해주세요");
		frm_edit.type.focus();
		return false;
	}
		
		frm_edit.mode.value = 'new';	
		frm_edit.mode_type.value = 'keyword';
		frm_edit.submit();
	}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_edit" action="/admin/diary2009/option/option_reg.asp" method="get">
	<input type="hidden" name="mode">
	<input type="hidden" name="mode_type">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">		
		<td align="center">옵션명</td>
		<td align="center">정렬순위</td>
		<td align="center">타입</td>
		<td align="center">사용여부</td>
    </tr>
	<tr align="center" bgcolor="ffffff">		
				<td align="center">
					<input type="text" size=30 name="option_value" >
				</td>	
				<td align="center"><input type="text" size=10 name="option_order" ></td>
				<td align="center">
					<select name="type" >
						<option value="" >선택</option>
						<option value="style" >style</option>
						<option value="color" >color</option>
						<option value="concept" >concept</option>							
						<option value="size" >size</option>							
						<option value="form" >form</option>							
						<option value="material" >material</option>
					</select>
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
