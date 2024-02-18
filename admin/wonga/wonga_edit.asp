<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  월간원가보고서 그룹 카테고리 데이타 등록
' History : 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/wonga/wonga_month_class.asp"-->

<% 
dim menupos,gubun,category,field,yyyymm,chulgocount
	chulgocount = request("chulgocount")
	menupos = request("menupos")
	gubun = request("groupname")
	category = cint(request("category"))
	field = cint(request("field"))
	yyyymm = request("yyyymm")

dim owongamonth_re,i
	set owongamonth_re = new Cwongalist
		owongamonth_re.frectgubun = Request("groupname")
		owongamonth_re.fwongamonth_add()
%>

<script language="javascript">
function form_submit(){
	frmreg.action = "/admin/wonga/wonga_edit_process.asp";
	frmreg.submit();	
}
</script>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>수정</strong> / 기준값이란? 이루고자 하는 목표 달성치를 말합니다.</font>
			</td>			
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td><br></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!--표 헤드끝-->

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">	
	<form name="frmreg" method="post" action="">
	
		<tr bgcolor=ffffff>
			<td align="center">
				카테고리명 : 
			</td>
			<td colspan="5">
			<input type="hidden" name="yyyymm" value="<%= yyyymm %>">
			<input type="hidden" name="chulgocount" value="<%= chulgocount %>">
			<input type="hidden" name="category" size="20" maxlength="20" value="<%= category %>">
			<input type="text" name="category_box_0" size="20" maxlength="20" value="<%= frectcategoryname(category,0) %>">
			<input type="hidden" name="groupname" size="20" maxlength="20" value="<%= gubun %>"></td>
			
		</tr>

		<tr bgcolor=ffffff>
			<td align="center">필드명 : </td>
			<td><input type="text" name="field_box_0" size="20" maxlength="20" value="<%= frectfieldname(category,field) %>">
				<input type="hidden" name="field" value="<%= cstr(field) %>"></td>
		</tr>
		<tr bgcolor=ffffff>	
			<td align="center">기준값 : </td>
			<td><input type="text" name="gijun_box_0" size="20" maxlength="20" value="<%= frectgijunvalue(category,field) %>"></td>
		</tr>
		<tr bgcolor=ffffff>	
			<td align="center">값 : </td>
			<td><input type="text" name="value_box_0" size="20" maxlength="20" value="<%= frectfieldvalue(category,field) %>"></td>
		</tr>			
		
</form>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="left"><br><input type="button" value="수정하기" onclick="form_submit();"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->