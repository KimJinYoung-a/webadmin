<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 판매 등록 관리
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/search/recomKeywordItemCls.asp" -->
<%
Dim i
Dim group_no : group_no = requestCheckvar(Trim(request("group_no")),10)
Dim keyword : keyword = requestCheckvar(Trim(request("keyword")),50)
Dim page : page = requestCheckvar(Trim(request("page")),10)

if (page="") then page=1

Dim oRecomKeywordItem

set oRecomKeywordItem = new CRecomKeywordItem
oRecomKeywordItem.FPageSize = 50
oRecomKeywordItem.FCurrPage = page
oRecomKeywordItem.FRectGroup_no = group_no

if (group_no<>"") then
oRecomKeywordItem.getRecomKeywordItemList
end if
%>
<script language="javascript">
function NextPage(i){
    document.frm.page.value=i;
    document.frm.submit();
}

function AddRecomKeywordItem() {
	var frm = document.frmadditem;

	if (frm.itemid.value.length<4) {
		alert('상품코드를 입력하세요.');
        frm.itemid.focus();
		return;
	}

	if (frm.group_no.value.length<1) {
		alert('해당 그룹번호가 지정되지 않았습니다.');
		return;
	}

	if (confirm('상품을 추가 하시겠습니까?') == true) {
		frm.submit();
	}
}

function delItem(group_no,itemid){
    var frm = document.frmDel;

    if (confirm("해당 상품을 삭제하시겠습니까?")){
        frm.group_no.value=group_no;
        frm.itemid.value=itemid;
        frm.submit();
    }
    
}
</script>
<!-- 액션 시작 -->
<p>
<form name="frm" method="get">
<input type="hidden" name="group_no" value="<%=group_no%>">
<input type="hidden" name="keyword" value="<%=keyword%>">
<input type="hidden" name="page" value="<%=page%>">
</form>
<p>
<form name="frmadditem" method="post" action="keywordRecom_Process.asp">
<input type="hidden" name="mode" value="additem">
<input type="hidden" name="group_no" value="<%=group_no%>">
<input type="hidden" name="keyword" value="<%=keyword%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			키워드 : <strong><%=keyword %></strong>
		</td>
		<td align="right">
		    상품코드:<input type="text" name="itemid" value="" size="10" maxlength="10">
		    <input type="button" class="button" value="상품 추가" onClick="AddRecomKeywordItem()">
			&nbsp;
		</td>
	</tr>
</table>
</form>
<!-- 액션 끝 -->
<p>

<table width="100%" align="center" cellpadding="4" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
	    <td colspan="20">
    		검색결과 : <b><%= oRecomKeywordItem.FTotalcount %></b>
    		&nbsp;
    		페이지 : <b><%= page %> / <%= oRecomKeywordItem.FTotalPage %></b>
    	</td>
    </tr>
    
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="80" height="22" >상품코드</td>
    	<td width="50">이미지</td>
    	<td width="100">브랜드ID</td>
		<td width="200">상품명</td>
		
        <td width="100">판매가</td>
        <td width="100">매입가</td>
        <td width="100">매입구분</td>
        <td width="70">판매여부</td>
        <td width="70">사용여부</td>
        <td width="90">한정여부</td>
		<td width="80">삭제</td>
	</tr>
	<%
	for i = 0 To oRecomKeywordItem.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
	    <td height="22" ><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oRecomKeywordItem.FItemList(i).Fitemid %>" target="_blank" title="미리보기"><%= oRecomKeywordItem.FItemList(i).Fitemid %></a></td>
		<td><img src="<%= oRecomKeywordItem.FItemList(i).Fsmallimage %>"></td>
        <td align="left"><%= oRecomKeywordItem.FItemList(i).Fmakerid %></td>
        <td align="left"><%= oRecomKeywordItem.FItemList(i).Fitemname %></td>
        
		<td align="right">
        <%
            Response.Write FormatNumber(oRecomKeywordItem.FItemList(i).Forgprice,0) 
			'할인가
			if oRecomKeywordItem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((oRecomKeywordItem.FItemList(i).Forgprice-oRecomKeywordItem.FItemList(i).Fsailprice)/oRecomKeywordItem.FItemList(i).Forgprice*100) & "%할)" & FormatNumber(oRecomKeywordItem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'쿠폰가
			if oRecomKeywordItem.FItemList(i).FitemCouponYn="Y" then
				Select Case oRecomKeywordItem.FItemList(i).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oRecomKeywordItem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oRecomKeywordItem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
				end Select
			end if
        %>
        </td>
        <td align="right">
        <%
            '할인가
			if oRecomKeywordItem.FItemList(i).Fsailyn="Y" then
			    if (oRecomKeywordItem.FItemList(i).Fsailsuplycash>oRecomKeywordItem.FItemList(i).Forgsuplycash) then
			        Response.Write "<strong>"&FormatNumber(oRecomKeywordItem.FItemList(i).Forgsuplycash,0)&"</strong>"
			        Response.Write "<br><strong><font color=#F08050>" & FormatNumber(oRecomKeywordItem.FItemList(i).Fsailsuplycash,0) & "</font></strong>"
			    else
			        Response.Write FormatNumber(oRecomKeywordItem.FItemList(i).Forgsuplycash,0)
    				Response.Write "<br><font color=#F08050>" & FormatNumber(oRecomKeywordItem.FItemList(i).Fsailsuplycash,0) & "</font>"
    			end if
    		else
    		    Response.Write FormatNumber(oRecomKeywordItem.FItemList(i).Forgsuplycash,0)
			end if
			'쿠폰가
			if oRecomKeywordItem.FItemList(i).FitemCouponYn="Y" then
				if oRecomKeywordItem.FItemList(i).FitemCouponType="1" or oRecomKeywordItem.FItemList(i).FitemCouponType="2" then
					if oRecomKeywordItem.FItemList(i).Fcouponbuyprice=0 or isNull(oRecomKeywordItem.FItemList(i).Fcouponbuyprice) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(oRecomKeywordItem.FItemList(i).Forgsuplycash,0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(oRecomKeywordItem.FItemList(i).Fcouponbuyprice,0) & "</font>"
					end if
				end if
			end if
        %>
        </td>
			
		<td><%= fnColor(oRecomKeywordItem.FItemList(i).Fmwdiv,"mw") %></td>
        <td ><%= fnColor(oRecomKeywordItem.FItemList(i).Fsellyn,"yn") %></td>
		<td><%= fnColor(oRecomKeywordItem.FItemList(i).Fisusing,"yn") %></td>

        <td>
            <% if oRecomKeywordItem.FItemList(i).Flimityn="Y" then %>
                한정(<%=oRecomKeywordItem.FItemList(i).GetLimitEa%>)
            <% end if %>
        </td>
		<td>
		    <input type="button" value="삭제" class="button" onClick="delItem('<%= group_no %>','<%=oRecomKeywordItem.FItemList(i).Fitemid%>')">    
		</td>
        
	</tr>
	<%
	next
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="11">
	<% if (oRecomKeywordItem.FTotalCount <1) then %>
			검색결과가 없습니다.
    <% else %>
        <% if oRecomKeywordItem.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oRecomKeywordItem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oRecomKeywordItem.StartScrollPage to oRecomKeywordItem.FScrollCount + oRecomKeywordItem.StartScrollPage - 1 %>
			<% if i>oRecomKeywordItem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oRecomKeywordItem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	<% end if %>
	    </td>
	</tr>
</table>

<form name="frmDel" method="post" action="keywordRecom_Process.asp">
<input type="hidden" name="mode" value="delitem">
<input type="hidden" name="group_no" value="">
<input type="hidden" name="itemid" value="">
</form>
<% 
SET oRecomKeywordItem = NOTHING
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
