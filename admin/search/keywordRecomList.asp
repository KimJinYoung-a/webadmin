<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/search/recomKeywordItemCls.asp" -->
<%

dim i, page
dim research : research         = request("research")
dim searchKeyword : searchKeyword = requestCheckvar(Trim(request("searchKeyword")),50)

''catecode  = Trim(requestCheckvar(request("catecode"),30))

page = request("page")
if (page="") then page=1
    

'// ============================================================================
dim oRecomKeyword

set oRecomKeyword = new CRecomKeywordItem
oRecomKeyword.FPageSize=50
oRecomKeyword.FCurrPage = page
oRecomKeyword.FRectSearchKeyword = searchKeyword

oRecomKeyword.getRecomKeywordMasterList

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function NextPage(i){
    document.frm.page.value=i;
    document.frm.submit();
}


function AddRecomKeywords(){
    var frm = document.frmaddkey;
    if (frm.keyword.value.length<1){
        alert('키워드를 입력해주세요.');
        frm.keyword.focus();
        return;
    }
    
    
    if (confirm('추가하시겠습니까?')){
        frm.submit();
    }
}


function delMaster(group_no,keyword){
    var frm = document.frmedtkey;
    frm.mode.value=="delmaster";
    frm.group_no.value=group_no;
    frm.keyword.value=keyword;
    
    if (confirm('삭제 하시겠습니까?')){
        if (confirm('정말 삭제 하시겠습니까?. 등록된 상품 목록도 삭제됩니다.')){
            frm.submit();
        }
    }   
}

function viewItemList(group_no,keyword){
    var popwin = window.open('popRecomKeywordItemlist.asp?group_no='+group_no+'&keyword='+keyword,'popRecomKeywordItemlist','width=1200,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left" height="30" >
			
			키워드 : <input type="text" class="text" name="searchKeyword" value="<%=searchKeyword%>" size="20">
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value=" 검 색 " onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
<p>
<!-- 액션 시작 -->
<form name="frmaddkey" method="post" action="keywordRecom_Process.asp">
<input type="hidden" name="mode" value="addmaster">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			
		</td>
		<td align="right">
		    키워드:<input type="text" name="keyword" value="" size="10" maxlength="20">
		    <input type="button" class="button" value="키워드 추가" onClick="AddRecomKeywords()">
			&nbsp;
		</td>
	</tr>
</table>
</form>
<!-- 액션 끝 -->
<p>

<!-- 리스트 시작 -->
<form name="frmSubmit" method="post" action="keywordRecom_Process.asp">
<table width="100%" align="center" cellpadding="4" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
	    <td colspan="20">
    		검색결과 : <b><%= oRecomKeyword.FTotalcount %></b>
    		&nbsp;
    		페이지 : <b><%= page %> / <%= oRecomKeyword.FTotalPage %></b>
    	</td>
    </tr>
    
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="80" height="22" >No.</td>
    	<td width="100">검색어</td>
    	
		<td width="80">등록상품수</td>
		<td width="200">상품코드</td>
		<td >상품명</td>
       
		<td width="80">삭제</td>
        <td width="80">목록</td>
	</tr>
	<%
	for i = 0 To oRecomKeyword.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
	    <td height="22" >
	        <%= oRecomKeyword.FItemList(i).Fgroup_no %>
	    </td>
		<td>
			<%= oRecomKeyword.FItemList(i).Fkeyword %>
		</td>
		<td><%= formatNumber(oRecomKeyword.FItemList(i).Fitemcnt,0) %></td>
			
		<td align="left">
			<%= oRecomKeyword.FItemList(i).Fitemid_list %>
		</td>
		<td align="left">
		    <%= oRecomKeyword.FItemList(i).Fitemname_list %>
		</td>
		<td>
		    <input type="button" value="키워드 삭제" class="button" onClick="delMaster('<%= oRecomKeyword.FItemList(i).Fgroup_no %>','<%=oRecomKeyword.FItemList(i).Fkeyword%>')">    
		</td>
        <td>
		    <input type="button" value="목록보기" class="button" onClick="viewItemList('<%=oRecomKeyword.FItemList(i).Fgroup_no%>','<%=oRecomKeyword.FItemList(i).Fkeyword%>')">    
		</td>
	</tr>
	<%
	next
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="7">
	<% if (oRecomKeyword.FTotalCount <1) then %>
			검색결과가 없습니다.
    <% else %>
        <% if oRecomKeyword.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oRecomKeyword.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oRecomKeyword.StartScrollPage to oRecomKeyword.FScrollCount + oRecomKeyword.StartScrollPage - 1 %>
			<% if i>oRecomKeyword.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oRecomKeyword.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	<% end if %>
	    </td>
	</tr>
</table>
</form>

<form name="frmedtkey" method="post" action="keywordRecom_Process.asp">
<input type="hidden" name="mode" value="delmaster">
<input type="hidden" name="group_no" value="">
<input type="hidden" name="keyword" value="">

</form>

<%
set oRecomKeyword = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
