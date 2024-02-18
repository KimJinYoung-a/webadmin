<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/upcheitemqna/LecDiyqnaCls.asp"-->

<%
'####################################################
' Description :  강좌&상품 Q&A 관리 리스트
' History : 2016.08.05 유태욱 생성
'####################################################
%>
<%
Dim page, i, research, isanswer, searchDiv, searchGubun, searchKey, searchString
Dim oMyqna
page			= RequestCheckvar(request("page"),10)
research		= RequestCheckvar(request("research"),2)
isanswer		= RequestCheckvar(request("isanswer"),2)
searchDiv		= RequestCheckvar(request("searchDiv"),2)
searchGubun		= RequestCheckvar(request("searchGubun"),1)
searchKey		= RequestCheckvar(request("searchKey"),16)
searchString	= request("searchString")
if searchString <> "" then
	if checkNotValidHTML(searchString) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
If page = "" Then page = 1
If (research = "") Then
	isanswer	= "N"
	searchKey	= "title"
End If

Set oMyqna = new CQna
	oMyqna.FCurrPage			= page
	oMyqna.FPageSize			= 18
	oMyqna.FRectisanswer		= isanswer
	oMyqna.FRectMakerUserid		= session("ssBctId")
	oMyqna.FRectsearchDiv		= searchDiv
	oMyqna.FRectsearchGubun		= searchGubun
	oMyqna.FRectsearchKey		= searchKey
	oMyqna.FRectsearchString	= searchString
	oMyqna.getMyqnaList
	
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function goView(vidx, vgridx,qnagubun){
	location.href='/cscenterv2/upcheitemqna/Qna/myqnaView.asp?menupos=<%=menupos%>&idx='+vidx+'&gridx='+vgridx+'&qnagubun='+qnagubun;
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		상태 :
		<select name="isanswer" class="select" onChange="">
		    <option value="" <%= chkiif(isanswer = "", "selected", "") %> >전체</option>
			<option value="Y" <%= chkiif(isanswer = "Y", "selected", "") %> >완료</option>
			<option value="N" <%= chkiif(isanswer = "N", "selected", "") %> >대기</option>
		</select>
		&nbsp;
		문의분야 : 
		<select name="searchGubun" class="select" onChange="">
			<option value="">선택</option>
			<option value="L" <%= chkiif(searchGubun="L", "selected", "") %> >강좌</option>
			<option value="D" <%= chkiif(searchGubun="D", "selected", "") %> >상품</option>
		</select>
	</td>
<!--		&nbsp;
		문의분야 : 
		<select name="searchDiv" class="select" onChange="">
			<option value="">선택</option>
			<option value="1" <%= chkiif(searchDiv="1", "selected", "") %> >강좌문의</option>
			<option value="2" <%= chkiif(searchDiv="2", "selected", "") %> >재료문의</option>
			<option value="3" <%= chkiif(searchDiv="3", "selected", "") %> >신설 강좌요청</option>
			<option value="4" <%= chkiif(searchDiv="4", "selected", "") %> >수강료,재료비 문의</option>
			<option value="5" <%= chkiif(searchDiv="5", "selected", "") %> >입금확인</option>
			<option value="6" <%= chkiif(searchDiv="6", "selected", "") %> >강좌취소</option>
			<option value="7" <%= chkiif(searchDiv="7", "selected", "") %> >강좌대기문의</option>
			<option value="8" <%= chkiif(searchDiv="8", "selected", "") %> >DIY 주문문의</option>
			<option value="9" <%= chkiif(searchDiv="9", "selected", "") %> >DIY 주문취소문의</option>
			<option value="10" <%= chkiif(searchDiv="10", "selected", "") %> >DIY 상품문의</option>
			<option value="11" <%= chkiif(searchDiv="11", "selected", "") %> >기타 문의</option>
		</select>
	</td>
-->
	<td align="right">
		검색 : 
		<select name="searchKey" class="select">
			<option value="idx" <%= chkiif(searchKey = "idx", "selected", "") %> >번호</option>
			<option value="title" <%= chkiif(searchKey = "title", "selected", "") %> >제목</option>
			<option value="comment" <%= chkiif(searchKey = "comment", "selected", "") %> >내용</option>
			<option value="titlecomment" <%= chkiif(searchKey = "titlecomment", "selected", "") %> >제목+내용</option>
			<option value="searchmakerid" <%= chkiif(searchKey = "searchmakerid", "selected", "") %> >작가/강사 ID</option>
			<option value="regmakername" <%= chkiif(searchKey = "regmakername", "selected", "") %> >작가/강사 이름</option>
			<option value="regid" <%= chkiif(searchKey = "regid", "selected", "") %> >등록자ID</option>
			<option value="regname" <%= chkiif(searchKey = "regname", "selected", "") %> >등록자이름</option>
		</select>
		<input type="text" class="text" name="searchString" size="20" value="<%=searchString%>">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td align="left" colspan="7">
		건수 : <b><%= FormatNumber(oMyqna.FTotalCount,0) %>건</b>
	</td>
	<td align="right">
		Page : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oMyqna.FTotalPage,0) %> </b>
	</td>
</tr>
<tr height="35" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">번호</td>
	<td width="80">문의분야</td>
	<td width="80">작가/강사 ID</td>
	<td width="150">작가/강사 이름</td>
	<td width="80">상태</td>
	<td>제목</td>
	<td width="140">등록자</td>
	<td width="140">등록일</td>
</tr>
<% For i=0 to oMyqna.FResultCount - 1 %>
<tr height="30" style="cursor:pointer;" align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff'; onclick="goView('<%= oMyqna.FItemList(i).FIdx %>','<%= oMyqna.FItemList(i).FReply_group_idx %>','<%= oMyqna.FItemList(i).FPagegubun %>')">
	<td align="center"><%= oMyqna.FItemList(i).FIdx %></td>
	<td align="center"><%=chkIIF(oMyqna.FItemList(i).FPagegubun="L","강좌","상품")%></td>
	<td align="center"><%= oMyqna.FItemList(i).Fmakerid %></td>
	<td align="center"><%=chkIIF(oMyqna.FItemList(i).FPagegubun="L",oMyqna.FItemList(i).Flecturer_name,oMyqna.FItemList(i).Fbrandname)%></td>
	<td align="center"><%= oMyqna.FItemList(i).getAnswerName %></td>
	<td align="left"><%= oMyqna.FItemList(i).FTitle %></td>
	<td align="center"><%= oMyqna.FItemList(i).FUserid %></td>
	<td align="center"><%= FormatDate(oMyqna.FItemList(i).FRegdate,"0000.00.00") %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oMyqna.HasPreScroll then %>
		<a href="javascript:goPage('<%= oMyqna.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oMyqna.StartScrollPage to oMyqna.FScrollCount + oMyqna.StartScrollPage - 1 %>
    		<% if i>oMyqna.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oMyqna.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<% Set oMyqna = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->