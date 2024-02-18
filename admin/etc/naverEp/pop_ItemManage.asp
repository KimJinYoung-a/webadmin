<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<%
Dim page, nItem, i
page = request("page")
If page = "" Then page = 1

SET nItem = new epShop
	nItem.FCurrPage		= page
	nItem.FPageSize		= 10
	nItem.getNaver3depthNameCandi
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function itemManageProc(v, iitemid){
	if (v == "del") {
		if(confirm("정말 삭제 하시겠습니까?")){
			document.frmAct.target = "xLink";
			document.frmAct.mode.value = v;
			document.frmAct.itemid.value = iitemid;
			document.frmAct.submit();
		}
	}else if (v == "add") {
			document.frmAct.target = "xLink";
			document.frmAct.mode.value = v;
			document.frmAct.itemid.value = iitemid;

			var postfix =$("#"+iitemid+"_postfix").val();
			var applyyn = $("input:radio[name="+iitemid+"_apply]:checked").val();
			document.frmAct.postfix.value = postfix;
			document.frmAct.applyyn.value = applyyn;
			document.frmAct.submit();
	}else if (v == "all") {
		if(confirm("일괄 등록 하시겠습니까?")){
			var chkSel = 0;
			var itemidArr = "";
			var postfixArr = "";
			var applyynArr = "";

			var cksel = document.getElementsByName('cksel');
			if(cksel.length > 1) {
				for(var i=0;i<frmSvArr.cksel.length;i++) {
					chkSel++;
					itemidArr = itemidArr + frmSvArr.cksel[i].value + "*(^!";
					postfixArr = postfixArr + frmSvArr.postfix[i].value + "*(^!";
					applyynArr = applyynArr + $("input:radio[name="+frmSvArr.cksel[i].value+"_apply]:checked").val() + "*(^!";
				}
			}
			document.frmAct.target = "xLink";
			document.frmAct.mode.value = v;
			document.frmAct.itemidArr.value = itemidArr;
			document.frmAct.postfixArr.value = postfixArr;
			document.frmAct.applyynArr.value = applyynArr;
			document.frmAct.submit();
		}
	}
}
function goPage(page){
    var frm = document.frmSearch;
    frm.page.value=page;
	frm.submit();
}
</script>
<form name="frmSearch" >
<input type="hidden" name="page" value="">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" id="frmSvArr">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">상품코드</td>
	<td width="60">이미지</td>
	<td width="100">브랜드</td>
	<td width="100">유입키워드1</td>
	<td width="100">유입키워드2</td>
	<td width="100">유입키워드3</td>
	<td>노출상품명</td>
	<td width="130">적용 여부</td>
	<td width="100">기능</td>
	<td width="50">링크</td>
</tr>
<% For i = 0 To nItem.FResultCount - 1 %>
<input type="hidden" id="cksel" name="cksel" value="<%= nItem.FItemList(i).FItemid %>" />
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><%= nItem.FItemList(i).FItemid %></td>
	<td><img src="<%= nItem.FItemList(i).FImageurl %>" width="50"></td>
	<td><%= nItem.FItemList(i).FSocname %></td>
	<td><%= nItem.FItemList(i).FKeyword1 %></td>
	<td><%= nItem.FItemList(i).FKeyword2 %></td>
	<td><%= nItem.FItemList(i).FKeyword3 %></td>
	<td><%= nItem.FItemList(i).FItemname %>_<input type="text" name="postfix" id="<%= nItem.FItemList(i).FItemid %>_postfix" size="50" value="<%= nItem.FItemList(i).FPostfix %>" /></td>
	<td>
		Y : <input type="radio" name="<%= nItem.FItemList(i).FItemid %>_apply" id="<%= nItem.FItemList(i).FItemid %>_applyY" value="Y" checked />
		N : <input type="radio"	name="<%= nItem.FItemList(i).FItemid %>_apply" id="<%= nItem.FItemList(i).FItemid %>_applyN" value="N" />
	</td>
	<td>
		<input type="button" name="" class="button" value="등록" onclick="itemManageProc('add', '<%= nItem.FItemList(i).FItemid %>');" />&nbsp;&nbsp;
		<input type="button" name="" class="button" value="삭제" onclick="itemManageProc('del', '<%= nItem.FItemList(i).FItemid %>');" />
	</td>
	<td><a href="http://www.10x10.co.kr/<%= nItem.FItemList(i).FItemid %>" target="_blank"><img src="/images/icon_search.jpg" /></a></td>
</tr>
<% Next %>
<% If (nItem.FTotalCount = 0) Then %>
<tr align="center" bgcolor="#FFFFFF">
	<td height="30" colspan="15">
		검색결과가 없습니다.
	</td>
</tr>
<% Else %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% If nItem.HasPreScroll then %>
		<a href="javascript:goPage('<%= nItem.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End if %>
		<% For i=0 + nItem.StartScrollPage to nItem.FScrollCount + nItem.StartScrollPage - 1 %>
			<% if i>nItem.FTotalPage Then Exit For %>
			<% if CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% Next %>

		<% If nItem.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% Else %>
			[next]
		<% End If %>
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<input type="button" value="일괄등록" class="button" onclick="itemManageProc('all', '');" />
		<input type="button" value="창닫기" class="button" onclick="window.close();" />
	</td>
</tr>
<% End If %>
</form>
</table>
<% SET nItem = nothing %>
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
<form name="frmAct" method="post" action="pop_itemManage_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="postfix" value="">
<input type="hidden" name="applyyn" value="">
<input type="hidden" name="itemidArr" value="">
<input type="hidden" name="postfixArr" value="">
<input type="hidden" name="applyynArr" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->