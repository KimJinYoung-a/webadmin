<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  마일리지 구분 
' History : 2007.10.23 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mileage/mileage_class.asp"-->
<%
dim jukyocd , menupos
	jukyocd = request("jukyocd")
	menupos = request("menupos")

dim omileage , i
set omileage = new Cmileagelist
	omileage.frectjukyocd = request("jukyocd")
	omileage.fmileage_add()
%>

<script language="javascript">
	function reg(){
	if(document.frm.jukyocd.value == ""){
	alert('마일리지 코드번호를 입력하세요');
	document.frm.jukyocd.focus();
	}
	else if(document.frm.jukyoname.value == ""){
	alert('코드명을 입력하세요');
	document.frm.jukyoname.focus();
	}
	else if(document.frm.isusing.value == ""){
	alert('상태를 입력하세요');
	document.frm.isusing.focus();
	}
	else
	{
	document.frm.action = "mileage_edit_process.asp";
	document.frm.submit();
	}
	}
</script>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>마일리지관리 수정</strong></font>
		</td>
	</tr>

	<form name="frm" action="" method="get">
	<input type="hidden" name="realjukyocd" value="<%= omileage.flist(0).fjukyocd %>"
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">
			마일리지코드 번호
		</td>
		<td align="center" bgcolor=#FFFFFF>
			<input type="text" class="text" name="jukyocd" value="<%= omileage.flist(0).fjukyocd %>" size="20" maxlength="20">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">
			코드명
		</td>
		<td align="center" bgcolor=#FFFFFF>
			<input type="text" class="text" name="jukyoname" value="<%= omileage.flist(0).fjukyoname %>" size="20" maxlength="20">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">
			상태
		</td>
		<td align="center" bgcolor=#FFFFFF>
			<select class="select" name="isusing">
				<option>선택</option>
				<option value="Y" <% if ucase(omileage.flist(0).fisusing) = "Y" then response.write "selected" %>>사용</option>
				<option value="N" <% if ucase(omileage.flist(0).fisusing) = "N" then response.write "selected" %>>사용안함</option>
			</select>
		</td>
	</tr>				
	</form>
	
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<input type="button" class="button" value="저장" onclick="javascript:reg();">
			&nbsp;
			<input type="button" class="button" value="닫기" onclick="javascript:window.close();">
		</td>
	</tr>
</table>	

		
<!--<script language="javascript">
	opener.location.reload();
	self.close();
</script>-->

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
