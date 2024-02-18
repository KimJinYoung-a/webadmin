<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%
Dim vAction
vAction = request("action")

If vAction = "update" Then
	Call Proc()
End IF

dim onesongjang,i
dim page

page = request("page")

if page="" then page=1

set onesongjang = new CEventsBeasong
onesongjang.FPageSize = 100
onesongjang.FCurrPage = page
onesongjang.getMomoBeasongInfoList
%>
<script language='javascript'>
function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function saveMiChulgo(iid){
	var count = 0;
	var num=document.getElementsByName("id").length
	var tmp = "";

	for(i=0; i<num; i++)
	{
		if(document.getElementsByName("id")[i].checked == true)
		{
			count +=1;
			tmp = document.getElementsByName("id")[i].value;
		}
	}
	if(count==0)
	{
		alert("재출력할 송장번호를 선택해 주세요.");
		return;
	}

    if (confirm('송장 재출력시 사용하는 메뉴 입니다.\n\n[주문번호 '+tmp+']\n\n전환 하시겠습니까?')){
        frmSubmit.action.value = "update";
        frmSubmit.orderid.value = tmp;
        frmSubmit.submit();
    }
}

</script>

<form name="frmSubmit" action="momosongjanglist.asp" method="post">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="action" value="">
<input type="hidden" name="orderid" value="">
</form>

※ 선택한것 송장 재출력 : <input class="button" type="button" value="재출력하기" onClick="saveMiChulgo()">

<form name="frm" method="get" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
</form>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>주문번호</td>
	<td width="70">아이디</td>
	<td width="60">고객명</td>
	<td width="60">수령인</td>
	<td>상품명</td>
	<td width="80">주문등록일</td>
	<td width="100">운송장번호</td>
	<td width="70">출고일</td>
</tr>
<% for i=0 to onesongjang.FResultCount-1 %>
<tr bgcolor="#FFFFFF" >
	<td align="center"><%= onesongjang.FItemList(i).Fid %></td>
	<td align="center"><%= onesongjang.FItemList(i).FUserId %></td>
	<td align="center"><%= onesongjang.FItemList(i).Fusername %></td>
	<td align="center"><%= onesongjang.FItemList(i).FReqName %></td>
	<td align="center"><%= onesongjang.FItemList(i).Fprizetitle %></td>
	<td align="center"><%= Left(onesongjang.FItemList(i).Fregdate,10) %></td>
	<td align="center"><input type="radio" name="id" value="<%= onesongjang.FItemList(i).Fid %>"><%= onesongjang.FItemList(i).FSongjangNo %></td>
	<td align="center"><% = Left(onesongjang.FItemList(i).Fsenddate,10) %></td>
</tr>
<% next %>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if onesongjang.HasPreScroll then %>
				<a href="javascript:NextPage('<%= onesongjang.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + onesongjang.StartScrollPage to onesongjang.FScrollCount + onesongjang.StartScrollPage - 1 %>
				<% if i>onesongjang.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>
		
			<% if onesongjang.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>

<%
set onesongjang = Nothing


Function Proc()
	Dim vAction, vOrderID
	vAction = request("action")
	vOrderID = request("orderid")
	
	If vAction = "update" AND vOrderID <> "" Then
		dbget.Execute "UPDATE [db_momo].[dbo].[tbl_momo_order] SET outputdate = null WHERE orderid = '" & vOrderID & "'"
	End If
	
	dbget.close()
	Response.Write "<script>alert('저장 되었습니다.');location.href='momosongjanglist.asp?menupos="&request("menupos")&"';</script>"
	Response.End
End Function
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->