<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모
' Hieditor : 2009.11.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
	Dim cCoinMng, vMngIdx, vIdx, vType, vItem, vOption, vUseYN, vItemDesc
	vMngIdx = Request("mng_idx")
	vIdx = Request("idx")
	
	dim page
	page = request("page")

	dim ttpgsz , CurrPage, i
	CurrPage = requestCheckVar(request("cpg"),9)

	IF CurrPage = "" then CurrPage=1
	if page = "" then page = 1
	
	set cCoinMng = new ClsMomoCoin
	If vIdx <> "" Then
		cCoinMng.FIdx = vIdx
		cCoinMng.FCoinMngItemView
		
		vMngIdx = cCoinMng.FOneItem.fmng_idx
		vIdx = cCoinMng.FOneItem.fidx
		vType = cCoinMng.FOneItem.ftype
		vItem = cCoinMng.FOneItem.fitem
		vOption = cCoinMng.FOneItem.foption
		vUseYN = cCoinMng.FOneItem.fuseyn
		vItemDesc = cCoinMng.FOneItem.fitem_desc
	End If
	
	cCoinMng.FPageSize = 10
	cCoinMng.FCurrPage = page
	cCoinMng.FMngIdx = vMngIdx
	cCoinMng.FCoinMngItemList
%>
<script language="javascript">
function checkform()
{
	if(frm.type.value == "")
	{
		alert('아이템 type을 선택하세요.');
		frm.type.focus();
		return false;
	}
	if(frm.item.value == "")
	{
		alert('사은품No 또는 쿠폰idx 를 입력해주세요.');
		frm.item.focus();
		return false;
	}
	if (!frm.useyn[0].checked && !frm.useyn[1].checked)
	{
		alert("사용여부를 선택하세요.")
		return false;
	}
	if(frm.item_desc.value == "" || frm.item_desc.value == "해당 아이템 설명")
	{
		alert('해당 아이템 설명을 입력해주세요.');
		clean();
		return false;
	}
}

function clean()
{
	if(frm.item_desc.value == "" || frm.item_desc.value == "해당 아이템 설명")
	{
		frm.item_desc.value = "";
		frm.item_desc.focus();
		return false;
	}
}

function buttonchange(gubun)
{
	c.style.display = "none";
	i.style.display = "none";
	i1.style.display = "none";
	
	if(gubun == "c")
	{
		c.style.display = "block";
	}
	else if(gubun == "i" || gubun == "s")
	{
		i.style.display = "block";
		i1.style.display = "block";
	}
}

function popWindow(gubun)
{
	if(gubun == "c")
	{
		window.open('/admin/sitemaster/couponlist.asp','coupon','width=700,height=700');
	}
	else if(gubun == "i" || gubun == "s")
	{
		window.open('pop_prod_list.asp','item','width=800,height=500');
	}
}
</script>

<form name="frm" method="post" action="coin_manage_item_proc.asp" onSubmit="return checkform(this);">
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="80" bgcolor="<%= adminColor("gray") %>">Manage No.</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">idx</td>
	<td width="120" bgcolor="<%= adminColor("gray") %>">type</td>
	<td width="300" bgcolor="<%= adminColor("gray") %>">상품id,옵션코드<br>or 쿠폰idx</td>
	<td width="80" bgcolor="<%= adminColor("gray") %>">사용여부</td>
</tr>
<tr height="30">
	<td bgcolor="FFFFFF"><input type="text" name="mng_idx" value="<%=vMngIdx%>" size="9" readonly></td>
	<td bgcolor="FFFFFF"><input type="text" name="idx" value="<%=vIdx%>" size="5" readonly></td>
	<td bgcolor="FFFFFF">
		<select name="type" onChange="buttonchange(this.value)">
			<option value="">type 선택</option>
			<option value="i" <% If vType = "i" Then Response.Write "selected" End If %>>상품</option>
			<option value="c" <% If vType = "c" Then Response.Write "selected" End If %>>쿠폰</option>
			<option value="s" <% If vType = "s" Then Response.Write "selected" End If %>>Secret선물</option>
		</select>
	</td>
	<td bgcolor="FFFFFF">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td valign="top"><input type="text" name="item" value="<%=vItem%>" size="10"><div id="c" style="display:none"><input type="button" value="쿠폰" onClick="popWindow('c')"></div><div id="i" style="display:none"><input type="button" value="상품" onClick="popWindow('i')"></div></td>
			<td style="padding-left:5"><div id="i1" style="display:none">옵션 코드 : <input type="text" name="option" value="<%=vOption%>" size="10"><br>옵션코드가 있는지 확인하고 있다면<br>반드시 옵션코드를 입력해야 합니다.</div></td>
		</tr>
		</table>
		<% If vType <> "" Then %><script>buttonchange("<%=vType%>");</script><% End If %>
	</td>
	<td bgcolor="FFFFFF">
		<input type="radio" name="useyn" value="y" <% If vUseYN = "y" Then Response.Write "checked" End If %>>Y&nbsp;
		<input type="radio" name="useyn" value="n" <% If vUseYN = "n" Then Response.Write "checked" End If %>>N
	</td>
</tr>
<tr height="30">
	<td colspan="4" bgcolor="FFFFFF">
		<input type="text" name="item_desc" value="<% If vItemDesc = "" Then Response.Write "해당 아이템 설명" Else Response.Write vItemDesc End If %>" size="58" onClick="clean()">
	</td>
	<td bgcolor="FFFFFF" align="center"><input type="submit" value="저  장"></td>
</tr>
</table>
</form>

<!-- 리스트 시작 -->
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if cCoinMng.FResultCount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
			<tr>
				<td>검색결과 : <b><%= cCoinMng.FTotalCount %></b></td>
				<td align="right"></td>
			</tr>
			</table>
		</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to cCoinMng.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF">
		<td>
			<table cellpadding="3" cellspacing="1" border="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
			    <td width="90" align="center" bgcolor="FFFFFF"><%= cCoinMng.FItemList(i).fmng_idx %></td>
			    <td width="55" align="center" bgcolor="FFFFFF"><%= cCoinMng.FItemList(i).fidx %></td>
			    <td width="130" align="center" bgcolor="FFFFFF"><%= MomoItemType(cCoinMng.FItemList(i).ftype) %></td>
			    <td width="330" align="center" bgcolor="FFFFFF"><%= cCoinMng.FItemList(i).fitem %>
			    	<% If cCoinMng.FItemList(i).foption <> "" Then Response.Write " (옵션:" & cCoinMng.FItemList(i).foption & ")" End If %>
			    </td>
			    <td width="80" align="center" bgcolor="FFFFFF"><%= cCoinMng.FItemList(i).fuseyn %></td>
			</tr>
			<tr>
				<td width="605" colspan="4" bgcolor="FFFFFF"><%= cCoinMng.FItemList(i).fitem_desc %></td>
				<td width="80" bgcolor="FFFFFF" align="center"><input type="button" value="수정" onClick="javascript:location.href='?mng_idx=<%= cCoinMng.FItemList(i).fmng_idx %>&idx=<%= cCoinMng.FItemList(i).fidx %>';"></td>
			</tr>
			</table>
	</tr>
	<% next %>
    </tr>   
    
<% else %>

	<tr bgcolor="#FFFFFF">
		<td width="500" colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if cCoinMng.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= cCoinMng.StartScrollPage-1 %>&mng_idx=<%=vMngIdx%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cCoinMng.StartScrollPage to cCoinMng.StartScrollPage + cCoinMng.FScrollCount - 1 %>
				<% if (i > cCoinMng.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cCoinMng.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&mng_idx=<%=vMngIdx%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cCoinMng.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&mng_idx=<%=vMngIdx%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<% 	set cCoinMng = nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
