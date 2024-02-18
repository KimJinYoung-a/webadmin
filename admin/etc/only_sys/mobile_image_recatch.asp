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
	Dim vQuery, cItem, vItemID, arrList, intLoop
	vItemID = Replace(requestCheckVar(Request("itemid"),300)," ","")
	
	If vItemID <> "" Then
		Set cItem = new cOnlySys
		cItem.FItemID = vItemID

		arrList = cItem.fnItemDetail
		Set cItem = Nothing
	End IF
	
	
	vQuery = vQuery & "select * from db_etcmall.dbo.tbl_outmall_API_Que" & vbCrLf
	vQuery = vQuery & "where itemid in (" & vItemID & ")" & vbCrLf & vbCrLf
	
	vQuery = vQuery & "--insert into db_etcmall.dbo.tbl_outmall_API_Que" & vbCrLf
	vQuery = vQuery & "select 'appDTL','EDIT',itemid,1100,GETDATE(),NULL,NULL,NULL,NULL,'" & session("ssBctId") & "'" & vbCrLf
	vQuery = vQuery & "from db_item.dbo.tbl_item" & vbCrLf
	vQuery = vQuery & "where itemid in (" & vItemID & ")" & vbCrLf
%>

<script language="javascript">
function jsItemSearch()
{
	if(frm1.itemid.value == "")
	{
		alert("상품코드 필수입니다.");
		frm1.itemid.focus();
		return;
	}
	frm1.submit();
}
function jsItemUpdate()
{
	if(frm1.itemid.value == "")
	{
		alert("itemid값이 필요합니다.");
		frm1.itemid.focus();
		return;
	}
	
	if(confirm("이대로 진행하시겠습니까?") == true) {
		frm1.method = "post";
		frm1.action = "mobile_image_recatch_proc.asp";
		frm1.submit();
	} else {
		return;
	}
}
</script>

<table class="a">
<tr>
	<td>
		<form name="frm1" action="<%=CurrURL%>" method="get">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td>
				ItemID : <input type="text" name="itemid" value="<%=vItemID%>" maxlength="32" size="100">&nbsp;
				<input type="button" class="button" value="검 색" onClick="jsItemSearch()"> * ,쉼표로 여러개 입력
			</td>
		</tr>
		<% If vItemID <> "" Then %>
			<tr>
				<td>
				<input type="button" value="다시캡쳐하기" onClick="jsItemUpdate()"></td>
			</tr>
		<% End If %>
		</table>
		</form>
		<% IF isArray(arrList) THEN %>
		<br>
		[db_item].[dbo].[tbl_Item]
		<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>ItemID</td>
			<td>ItemName</td>
		  	<td>PC링크</td>
		  	<td>모웹링크</td>
		  	<td>모앱링크</td>
		  	<td>lastupdate</td>
		</tr>
			<% For intLoop =0 To UBound(arrList,2) %>
				<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
					<td><%=arrList(0,intLoop)%></td>
					<td><%=arrList(1,intLoop)%></td>
					<td><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank">[PC링크]</a></td>
					<td><a href="http://m.10x10.co.kr/category/category_itemprd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank">[모웹링크]</a></td>
					<td><a href="http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank">[모앱링크]</a></td>
					<td><%=arrList(5,intLoop)%></td>
				</tr>
			<% Next %>
		</table>
		<% End If %>
	</td>
</tr>
</table>


<% If vItemID <> "" Then %>
<br><br>* 쿼리구문<br>
<textarea name="" cols="100" rows="15"><%=vQuery%></textarea>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->