<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<%
dim research,userid, fixtype, linktype, poscode, validdate
dim page, vItemID, vUseYN

	page    = request("page")
	vItemID = Request("itemid")
	vUseYN  = Request("useyn")

if page = "" then page = 1

	dim cMomoBonusCoinList, PageSize , ttpgsz , CurrPage, i
	page = requestCheckVar(request("page"),9)

	IF CurrPage = "" then CurrPage=1
	if page = "" then page = 1
	

	'### 내가 사용 코인 내역
	set cMomoBonusCoinList = new ClsMomoCoin
	cMomoBonusCoinList.FPageSize = 5
	cMomoBonusCoinList.FCurrPage = page
	cMomoBonusCoinList.FItem_List
%>

<script language="javascript">
function checkform()
{
	if(frm1.itemid.value == "")
	{
		alert('Itemid 를 입력해주세요.');
		frm1.itemid.focus();
		return false;
	}
	if (!frm1.useyn[0].checked && !frm1.useyn[1].checked)
	{
		alert("사용여부를 선택하세요.")
		return false;
	}
}

function goCoinEdit(itemid,useyn)
{
	location.href = "pop_prod_list.asp?itemid="+itemid+"&useyn="+useyn+"";
}
</script>

<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm1" action="pop_prod_proc.asp" method="post" onSubmit="return checkform(this);">
<input type="hidden" name="gb" value="<% If vItemID <> "" Then Response.Write "u" End If %>">
<tr height="25" bgcolor="FFFFFF">
	<td align="center" width="150">Itemid : <input type="text" name="itemid" value="<%=vItemID%>" size="10"></td>
	<td align="center" width="200">
		사용 유무 : <input type="radio" name="useyn" value="y" <% If vUseYN = "y" Then Response.Write "checked" End If %>>Y&nbsp;
		<input type="radio" name="useyn" value="n" <% If vUseYN = "n" Then Response.Write "checked" End If %>>N
	</td>
	<td align="center" width="50"><input type="submit" value="저장"></td>
</tr>
</form>
</table>
<br>

<!-- 리스트 시작 -->
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if cMomoBonusCoinList.FResultCount > 0 then %> 
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= cMomoBonusCoinList.FTotalCount %></b>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(tip. [ ] 안에 마우스를 대고 더블클릭하시면 코드번호가 선택이 됩니다. 그 후 바로 ctrl+c 키를 누르면 복사가 됩니다.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td align="center" width="60"></td>
	    <td align="center" width="100%">Item</td>
	    <td align="center" width="40">useYN</td>
	    <td align="center" width="40"></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to cMomoBonusCoinList.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF">	
	    <td align="center"><img src="<%= cMomoBonusCoinList.FItemList(i).fimagesmall %>"></td>
	    <td>[<%= cMomoBonusCoinList.FItemList(i).fitem %>]<%= cMomoBonusCoinList.FItemList(i).fitemname %>
	    	<br>옵션 : <%=OptionList(cMomoBonusCoinList.FItemList(i).fitem)%>
	    </td>
	    <td align="center"><%= cMomoBonusCoinList.FItemList(i).fuseyn %></td>
		<td align="center"><input type="button" value="수정" onClick="javascript:goCoinEdit('<%= cMomoBonusCoinList.FItemList(i).fitem %>','<%= cMomoBonusCoinList.FItemList(i).fuseyn %>');"></td>
	</tr>
	<% next %>
    </tr>   
    
<% else %>

	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if cMomoBonusCoinList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= cMomoBonusCoinList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cMomoBonusCoinList.StartScrollPage to cMomoBonusCoinList.StartScrollPage + cMomoBonusCoinList.FScrollCount - 1 %>
				<% if (i > cMomoBonusCoinList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cMomoBonusCoinList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cMomoBonusCoinList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
	set cMomoBonusCoinList = nothing	
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->