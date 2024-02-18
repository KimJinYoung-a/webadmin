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
dim menupos,gubun,category_add
	menupos = request("menupos")
	gubun = request("gubunbox")		'기존그룹값이 잇을경우 선택을 위하여 구분값을 받아 온다.
	category_add = request("category_add_box")
	
	if category_add = "" then
		category_add = 1
	end if	
	
dim owongamonth_re,i
set owongamonth_re = new Cwongalist
	owongamonth_re.frectgubun = Request("gubunbox")
	owongamonth_re.fwongamonth_add()	
%>	
<%	
'###########################################################	그룹명 셀렉트박스
Sub DrawUserGubun(gubunbox,gubunid)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str
	
	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = "select groupname from"
	userquery = userquery & " db_datamart.dbo.tbl_month_wonga_category"
	userquery = userquery & " group by groupname"
	userquery = userquery & " order by groupname asc"
	db3_rsget.Open userquery, db3_dbget, 1
	
	response.write "<select onChange=javascript:check_gubun(this); name='" & gubunbox & "' "  '검색하고자하는 것을 셀렉트 네임으로 하고
	if gubunid <> "" then					'구분값이 있으면 선택을 못하도록 disabled
		response.write "disabled"
	end if	
	response.write ">"		
	response.write "<option value=''"							'옵션의 값이 없으면
		if gubunid ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">기존사용구분 선택</option>"								'선택이란 단어가 나오도록.

	if not db3_rsget.EOF then
		do until db3_rsget.EOF
			if Lcase(gubunid) = Lcase(db3_rsget("groupname")) then 	'검색될 이름과 db에 저장된 이름을 비교해서 맞다면, //
				tem_str = " selected"								'// 검색어로 선택
			end if
			response.write "<option value='" & db3_rsget("groupname") & "' " & tem_str & ">" & db2html(db3_rsget("groupname")) & "</option>"
			tem_str = ""				'db3_rsget에 gubunid 선택하고 검색할 값으로 선택
			db3_rsget.movenext
		loop
	end if
	response.write "</select>"
db3_rsget.close		
End Sub
%>

<script language="javascript">

<!-- 구분 검색시작-->
function check_gubun(frm)
{
	document.frmreg.groupname.value = "";
	document.frmreg.groupname.value = document.frmreg.gubunbox.value;
	document.frmreg.submit();
}
<!-- 구분 검색 끝-->

function form_submit(){
	if (document.frmreg.groupname.value=="")
	{
		alert('구분을 입력하세요');
		document.frmreg.groupname.focus();		 
	}
	else if (document.frmreg.yyyy.value=="")
	{
		alert('년도를 입력하세요');
		document.frmreg.yyyy.focus();		 
	}
	else if (document.frmreg.mm.value=="")
	{
		alert('달을 입력하세요');
		document.frmreg.mm.focus();		 
	}
	else if (document.frmreg.count.value=="")
	{
		alert('총 수량을 입력하세요');
		document.frmreg.count.focus();		 
	}
	else
	{
		frmreg.action = "/admin/wonga/wonga_add_category_data_process.asp";
		frmreg.submit();
	}	
}

<!-- 년도와 달 중복체크 시작-->
function yyyymmcheck(){
	if (document.frmreg.groupname.value=="")
	{
		alert('구분을 입력하세요');
		document.frmreg.groupname.focus();		 
	}
	else if (document.frmreg.yyyy.value=="")
	{
		alert('년도를 입력하세요');
		document.frmreg.yyyy.focus();		 
	}
	else if (document.frmreg.mm.value=="")
	{
		alert('달을 입력하세요');
		document.frmreg.mm.focus();		 
	}
	else
	{
		var yyyy = frmreg.yyyy.value;
		var mm = frmreg.mm.value;
		var groupname = frmreg.groupname.value;
		var popup = window.open('/admin/wonga/wonga_yyyymm_check.asp?yyyy='+yyyy+'&mm='+mm+'&groupname='+groupname,'yyyymmcheckpopup','width=1,height=1,scrollbars=yes,resizable=yes');
		popup.focus();	
	}	
}
<!-- 년도와 달 중복체크 끝-->

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
			<font color="red"><strong>그룹 , 카테고리 데이타 등록</strong> / 기준값이란? 이루고자 하는 목표 달성치를 말합니다.</font>
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

<% if owongamonth_re.ftotalcount = 0 then 
'##################################################################################################################	그룹 없음 시작	
%>

<table width="100%" border="0" class="a" cellpadding="1" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#ffffff><form name="frmreg" method="post" action="">
		<td align="center">
			사용구분(필수) : 
		</td>
		<td colspan="3">
			<% DrawUserGubun "gubunbox", gubun %> &nbsp;&nbsp;&nbsp;<input type="hidden" name="gubun_submit" value="<%= gubun %>">
			선택그룹 : <input type="text" name="groupname" size="20" maxlength="20" disabled> (ex: 물류)
		</td>
	</tr>
</form>
</table>


<%
'##################################################################################################################	기존그룹 시작
else %>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#ffffff><form name="frmreg" method="post" action="">
		<td align="center">
			사용구분(필수) : 
		</td>
		<td colspan="5">
			<% DrawUserGubun "gubunbox", gubun %> &nbsp;&nbsp;&nbsp;<input type="hidden" name="gubun_submit" value="<%= gubun %>">
			새로입력 : <input type="text" name="groupname" size="20" maxlength="20" value="<%= gubun %>" disabled> (ex: 물류)
		</td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align="center">
			년,달 입력(필수) : 
		</td>
		<td colspan="5">
			<input type="text" name="yyyy" size="4" maxlength="4"> 
			<input type="text" name="mm" size="2" maxlength="2"> (ex: 2007 , 01)
			<input type="button" name="checkbutton" value="중복체크(필수)" onclick="yyyymmcheck();">
		</td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align="center">
			총 수량(필수) :  
		</td>
		<td colspan="5">
			<input type="text" name="count" size="20" maxlength="20"> ex: 물류 총 출고수량 &nbsp;&nbsp;&nbsp;
			<font color="red">계산값 = 물류(2007년01월)카테고리1(필드1)값 / 총수량(물류총출고량)</font>
		</td>
	</tr>
</table>	
		<%
	dim sql ,ftotalcount
		sql = "select"
		sql = sql & " category"
		sql = sql & " from db_datamart.dbo.tbl_month_wonga_category"
		sql = sql & " where 1=1 and groupname= '"& gubun &"' and category_isusing='y'"
		sql = sql & " group by category" 	
	db3_rsget.open sql,db3_dbget,1
	'response.write sql&"<br>"	
	ftotalcount = db3_rsget.recordcount
	db3_rsget.close
	%>	
<br>	
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">	
	<% for i = 0 to ftotalcount - 1 %>
		<tr bgcolor=ffffff>
			<td align="center">
				카테고리명<%= i %> (필수) : 
			</td>
			<td colspan="5"> <%= frectcategoryname(i,0) %>
			<input type="hidden" name="category_box_0" size="20" maxlength="20" value="<%= frectcategoryname(i,0) %>">
			<input type="hidden" name="groupname" size="20" maxlength="20" value="<%= gubun %>"></td>
			
		</tr>
		<%
		dim sql1 ,ffieldcount ,t
		sql1 = "select field"
		sql1 = sql1 & " from db_datamart.dbo.tbl_month_wonga_category"
		sql1 = sql1 & " where 1=1 and groupname= '"& gubun &"' and category_isusing='y' and category='"& i &"'"
	
		db3_rsget.open sql1,db3_dbget,1
		'response.write sql1&"<br>"	
		ffieldcount = db3_rsget.recordcount
		db3_rsget.close
		%>
		<% for t = 0 to ffieldcount -1 %>
			<tr bgcolor=ffffff>
				<td align="center">필드명 : </td>
				<td><input type="hidden" name="field_box_0" size="20" maxlength="20" value="<%= frectfieldname(i,t) %>"> <%= frectfieldname(i,t) %></td>
				<td align="center">기준값 : </td>
				<td><input type="text" name="gijun_box_0" size="20" maxlength="20" value="<%= frectgijunvalue(i,t) %>"></td>
				<td>값 : </td>
				<td><input type="text" name="value_box_0" size="20" maxlength="20" value=""></td>
			</tr>			
		<% next %>
	<% next %>
</form>
</table>
<% end if %>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right"><br><input type="button" value="저장하기" onclick="form_submit();">&nbsp;
        	<input type="button" value="닫기" onclick="javascript:window.close();"></td>
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