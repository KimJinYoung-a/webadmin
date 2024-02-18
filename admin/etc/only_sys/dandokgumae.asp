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
	Dim vQuery, cItem, vItemID, arrList, intLoop, vIsDandok, vIsSunChak
	vItemID = Replace(requestCheckVar(Request("itemid"),300)," ","")
	vIsDandok = NullFillWith(requestCheckVar(Request("dandok"),1),"o")
	vIsSunChak = requestCheckVar(Request("sunchak"),1)
	
	If vItemID <> "" Then
		Set cItem = new cOnlySys
		cItem.FItemID = vItemID

		arrList = cItem.fnItemDetail
		Set cItem = Nothing
	End IF
	
	vQuery = vQuery & "-- 단독구매상품(공동구매 등) 설정" & vbCrLf
	vQuery = vQuery & "select reserveItemTp, availPayType, * from db_item.dbo.tbl_item" & vbCrLf
	vQuery = vQuery & "where itemid in (" & vItemID & ")" & vbCrLf & vbCrLf
	
	vQuery = vQuery & "-- 신구분 지정 (0-일반, 1-단독구매)" & vbCrLf
	vQuery = vQuery & "-- 선착순결제(실시간/즉시) 상품 설정 (availPayType- 8:Just1Day ,9:선착순결제)" & vbCrLf
	vQuery = vQuery & "--update db_item.dbo.tbl_item" & vbCrLf
	vQuery = vQuery & "set reserveItemTp='1'" & vbCrLf
	vQuery = vQuery & "--, availPayType='9'" & vbCrLf
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
		frm1.action = "dandokgumae_proc.asp";
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
				ItemID : <input type="text" name="itemid" value="<%=vItemID%>" size="100">&nbsp;
				<input type="button" class="button" value="검 색" onClick="jsItemSearch()"> * ,쉼표로 여러개 입력
			</td>
		</tr>
		<% If vItemID <> "" Then %>
			<tr>
				<td><br>
				<input type="checkbox" name="dandok" value="o" <%=CHKIIF(vIsDandok="o","checked","")%>>단독구매설정(reserveItemTp=1)&nbsp;&nbsp;&nbsp;
				<input type="checkbox" name="sunchak" value="o" <%=CHKIIF(vIsSunChak="o","checked","")%>>선착순결제(실시간/즉시)설정(availPayType=9)
				<input type="button" value="바로변경하기" onClick="jsItemUpdate()"></td>
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
		  	<td>reserveItemTp</td>
		  	<td>availPayType</td>
		  	<td>RegDate</td>
		  	<td>lastupdate</td>
		</tr>
			<% For intLoop =0 To UBound(arrList,2) %>
				<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
					<td><%=arrList(0,intLoop)%></td>
					<td><%=arrList(1,intLoop)%></td>
					<td><%=arrList(2,intLoop)%></td>
					<td><%=arrList(3,intLoop)%></td>
					<td><%=arrList(4,intLoop)%></td>
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