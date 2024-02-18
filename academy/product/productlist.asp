<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_productcls.asp"-->

<%

dim oacademyprd
dim page
dim makerid, sellyn, isusing, selBest

page = RequestCheckvar(request("page"),10)
if page="" then page=1
selBest = RequestCheckvar(request("selBest"),1)

set oacademyprd = new CAcademyProduct
oacademyprd.FCurrPage = page
oacademyprd.FPageSize = 20
oacademyprd.FRectMakerid = makerid
oacademyprd.FRectSellYn = sellyn
oacademyprd.FRectIsUsing = isusing
oacademyprd.FRectBest	= selBest

oacademyprd.GetProductList


dim i
%>
<script language='javascript'>
function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

function popitemsearch(frm){
	var popwin;
	popwin = window.open("/admin/pop/viewitemlist.asp?designerid=" + "&target=" + frm, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function AddIttems(){
	var ret=confirm('선택 상품을 추가 하시겠습니까?');

	if(ret){
		frmbuf.submit();
	}
}

function QuickAdd(frm){
	if (frm.itemidarr.value.length<1){
		alert('값을 입력하세요.');
		frm.itemidarr.focus();
		return;
	}

	var ret=confirm('상품을 추가 하시겠습니까?');

	if(ret){
		frm.submit();
	}
}

function DellItems(frm, stype){
	var ret=confirm('선택 상품을 삭제 하시겠습니까?');

	if(ret){
		frm.mode.value = stype;
		frm.submit();
	}
}

//베스트 등록
function BestItems(frm, stype){
	var ret=confirm('선택 상품을 베스트로 등록 하시겠습니까?');

	if(ret){
		frm.mode.value = stype;
		frm.submit();
	}
}

//베스트 취소
function BestCancel(frm,stype,sId){
var ret=confirm('베스트를 취소 하시겠습니까?');

	if(ret){
		frm.mode.value = stype;
		frm.bestId.value = sId;
		frm.submit();
	}
}

//판매수정
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit_aca','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

// 이미지수정
function editItemImage(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage_aca','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" >
<form name="frmbuf" method="post" action="doacademyproduct.asp">
<input type="hidden" name="mode" value="addArr">
<tr>
	<td bgcolor="#FFFFFF" colspan="3">
	<input type="text" name="itemidarr" size="90" maxlength="90">
	<input type="button" value="상품코드로추가" onclick="QuickAdd(frmbuf)">
	</td>
</tr>
<tr>
	<td width="50">
		<input type="button" value="선택상품삭제" onclick="DellItems(frmlist,'dellarr');">
	</td>
	<td width="75%">
		<input type="button" value="선택상품 베스트등록 " onclick="BestItems(frmlist,'bestarr');">
	</td>
	<td width="50" bgcolor="#FFFFFF" align="right">
		<input type="button" value="목록에서선택추가" onclick="popitemsearch('frmbuf.itemidarr');">
	</td>
</tr>
</form>
</table>

<br>

<!-- 상단 검색폼 시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm" >
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="30">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" align="right">
       	베스트 구분 : 
       	<select name="selBest" onchange="javascript:document.frm.submit();">
       	<option value="">전체</option>
       	<option value="1" <%IF selBest = "1" THEN%>selected<%END IF%>>베스트</option>
       	<option value="2" <%IF selBest = "2" THEN%>selected<%END IF%>>베스트 제외</option>
       	</select>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- 상단 검색폼 끝 -->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frmlist" method=post action="doacademyproduct.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="bestId" value="">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="13" align="right">검색건수 : <%= oacademyprd.FTotalCount %> 건 Page : <%= page %>/<%= oacademyprd.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="20"></td>
		<td align="center" width="50">상품번호</td>
		<td align="center" width="50">이미지</td>
		<td align="center" width="70">브랜드</td>
		<td align="center">상품명</td>
		<td align="center" width="40">매입<br>구분</td>
		<td align="center" width="50">판매가</td>
		<td align="center" width="50">매입가</td>
		<td align="center" width="50">판매</td>
		<td align="center" width="50">한정</td>
		<td align="center" width="50">비고</td>
		<td align="center" width="50">베스트</td>
	</tr>
	<% for i=0 to oacademyprd.FResultCount -1 %>
	<tr <% if oacademyprd.FITemList(i).FisBest = "Y" then%>bgcolor="#F3F3FF"<%else%>bgcolor="#FFFFFF"<%end if%> align="center">
		<td><input type="checkbox" name="itemidarr" value="<%= oacademyprd.FITemList(i).FItemID %>" onClick="AnCheckClick(this);"></td>
		<td><a href="javascript:PopItemSellEdit('<%= oacademyprd.FITemList(i).FItemID %>')"><%= oacademyprd.FITemList(i).FItemID %></a></td>
		<td><a href="javascript:editItemImage('<%= oacademyprd.FITemList(i).FItemID %>')"><img src="<%= oacademyprd.FITemList(i).FSmallImage %>" width="50" border="0"></a></td>
		<td align="left"><%= oacademyprd.FITemList(i).FMakerid %></td>
		<td align="left"><a href="/admin/itemmaster/itemmodify.asp?itemid=<%= oacademyprd.FITemList(i).FItemID %>&menupos=594" target="_blank"><%= oacademyprd.FITemList(i).FItemName %></a></td>
		<td><%= oacademyprd.FITemList(i).GetMWdivStr %></td>
		<td align="right"><%= FormatNumber(oacademyprd.FITemList(i).FSellcash,0) %></td>
		<td align="right"><%= FormatNumber(oacademyprd.FITemList(i).FBuycash,0) %></td>
		<td><%= oacademyprd.FITemList(i).FSellyn %></td>
		<td><%= oacademyprd.FITemList(i).GetLimitStr %></td>
		<td>
			<% if oacademyprd.FITemList(i).IsSoldOut then %>
			<font color="red">품절</font>
			<% end if %>
		</td>
		<td>
			<%if oacademyprd.FITemList(i).FisBest = "Y" then%>
				<font color="red">베스트</font><br>
				<a href="javascript:BestCancel(frmlist,'unbest',<%= oacademyprd.FITemList(i).FItemID %>);">[x취소]</a>
			<%END IF%>
		</td>
	</tr>
	<% next %>
	<tr>
		<td align="center" colspan="13" bgcolor="#F0F0FD">
			<!-- 페이지 시작 -->
				<%
				if oacademyprd.HasPreScroll then
					Response.Write "<a href='javascript:NextPage(" & oacademyprd.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for i=0 + oacademyprd.StarScrollPage to oacademyprd.FScrollCount + oacademyprd.StarScrollPage - 1

					if i>oacademyprd.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:NextPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if oacademyprd.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:NextPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- 페이지 끝 -->
		</td>
	</tr>
</form>
</table>
<%
set oacademyprd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
