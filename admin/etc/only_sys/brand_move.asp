<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->

<%
	Dim cBrand, vQuery, vMakerID, vItemID, vNewMakerID, vBrandName, vMoveItemCnt, vChange
	vMakerID = requestCheckVar(Request("makerid"),100)
	vItemID = Request("itemid")
	vNewMakerID = requestCheckVar(Request("newmakerid"),100)
	vChange = requestCheckVar(Request("change"),1)
	
	
	If vNewMakerID <> "" Then
		Set cBrand = new cOnlySys
		cBrand.FNewMakerID = vNewMakerID
		cBrand.FMakerID = vMakerID
		cBrand.FItemID = vItemID
		cBrand.fnBrandCont
		
		vBrandName = cBrand.FBrandName
		vMoveItemCnt = cBrand.FMoveItemCnt
		Set cBrand = Nothing
	End IF
		
	vQuery = ""
	If vMakerID <> "" Then
		vQuery = vQuery & "select * from db_item.dbo.tbl_item where makerid = '" & vMakerID & "'" & vbCrLf & vbCrLf
		vQuery = vQuery & "select * from db_item.dbo.tbl_item where makerid = '" & vNewMakerID & "'" & vbCrLf & vbCrLf
	End IF
	If vItemID <> "" Then
		vQuery = vQuery & "select * from db_item.dbo.tbl_item where itemid IN(" & vItemID & ")" & vbCrLf & vbCrLf
	End IF
	vQuery = vQuery & "select * from db_user.dbo.tbl_user_c where userid = '" & vNewMakerID & "'" & vbCrLf & vbCrLf
	
	vQuery = vQuery & "--update db_item.dbo.tbl_item" & vbCrLf
	vQuery = vQuery & "set" & vbCrLf
	vQuery = vQuery & "makerid = '" & vNewMakerID & "', brandname = '" & vBrandName & "', lastupdate = getdate()" & vbCrLf
	vQuery = vQuery & "where itemid in" & vbCrLf
	vQuery = vQuery & "(" & vbCrLf
	If vMakerID <> "" Then
		vQuery = vQuery & "	select itemid from db_item.dbo.tbl_item" & vbCrLf
		vQuery = vQuery & "	where" & vbCrLf
		vQuery = vQuery & "	makerid = '" & vMakerID & "'" & vbCrLf
	End IF
	If vItemID <> "" Then
		vQuery = vQuery & vItemID & vbCrLf
	End IF
	vQuery = vQuery & ")" & vbCrLf
%>

<script language="javascript">
function jsbrandSearch()
{
	if(frm1.makerid.value == "" && frm1.itemid.value == "")
	{
		alert("이동할 브랜드(WHERE makerid = '')\n또는 이동할 상품코드(WHERE itemid IN (''))를\n입력하세요.");
		return;
	}
	if(frm1.makerid.value != "" && frm1.itemid.value != "")
	{
		alert("이동할 브랜드(WHERE makerid = '')\n또는 이동할 상품코드(WHERE itemid IN (''))를\n하나만 입력하세요.");
		return;
	}
	if(frm1.newmakerid.value == "")
	{
		alert("1 또는 2 가 이동되어야 할 브랜드(SET makerid = '')를 입력하세요.");
		frm1.newmakerid.focus();
		return;
	}
	frm1.submit();
}
function jsBrandUpdate()
{
	if(frm1.makerid.value == "" && frm1.itemid.value == "")
	{
		alert("이동할 브랜드(WHERE makerid = '')\n또는 이동할 상품코드(WHERE itemid IN (''))를\n입력하세요.");
		return;
	}
	if(frm1.makerid.value != "" && frm1.itemid.value != "")
	{
		alert("이동할 브랜드(WHERE makerid = '')\n또는 이동할 상품코드(WHERE itemid IN (''))를\n하나만 입력하세요.");
		return;
	}
	if(frm1.newmakerid.value == "")
	{
		alert("1 또는 2 가 이동되어야 할 브랜드(SET makerid = '')를 입력하세요.");
		frm1.newmakerid.focus();
		return;
	}
	
	if(confirm("이대로 진행하시겠습니까?") == true) {
		frm1.method = "post";
		frm1.action = "brand_move_proc.asp";
		frm1.submit();
	} else {
		return;
	}
}
</script>

<br>
[<a href="brand_move.asp"><font size="5" color="blue"><strong><u>온라인</u></strong></font></a>]&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
[<a href="brand_move_offline.asp"><font size="3">오프라인</font></a>]
<br><br>

<table class="a">
<tr>
	<td>
		<form name="frm1" action="<%=CurrURL%>" method="get">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td colspan="2">
				<font color="blue">1</font>. 이동할 브랜드(WHERE makerid = '') : <input type="text" name="makerid" value="<%=vMakerID%>" maxlength="50">&nbsp;또는(둘 중 하나만)<br>
				<font color="blue">2</font>. 이동할 상품코드(WHERE itemid IN ('')) : <textarea name="itemid" cols="30" rows="7"><%=vItemID%></textarea><br>
				※ 입력 예 (123456 또는 123456,234567,345678 또는 123456, 234567, 345678)<br><br>
				<font color="blue">3</font>. <font color="blue"><b>1</b></font> 또는 <font color="blue"><b>2</b></font> 가 이동되어야 할 브랜드(SET makerid = '') : 
				<input type="text" name="newmakerid" value="<%=vNewMakerID%>" maxlength="50"><br>
				<input type="button" class="button" value="검      색" onClick="jsbrandSearch()">
			</td>
		</tr>
		</table>
		<% If vMakerID <> "" OR vItemID <> "" Then %>
		<br><br>
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td colspan="2">
				이동되어야 할 브랜드명(SET brandname = '') : <input type="text" name="brandname" value="<%=vBrandName%>" readonly><br>
				<% If vChange = "o" Then %><font color="red">이동완료.</font><% Else %>총 <%=vMoveItemCnt%>개 상품<% End If %> <input type="button" value="바로변경하기" onClick="jsBrandUpdate()">
			</td>
		</tr>
		</table>
		<% End If %>
		</form>
	</td>
</tr>
</table>

<% If vMakerID <> "" OR vItemID <> "" Then %>
<br><br>* 쿼리구문<br>
<textarea name="" cols="100" rows="17"><%=vQuery%></textarea>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->