<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.29 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/helloCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim lhello, page, makerid, isusing, i, brandgubun
dim catecode, standardCateCode, mduserid
	catecode	= request("catecode")
	standardCateCode	= request("standardCateCode")
	mduserid	= request("mduserid")
	brandgubun	= request("brandgubun")	
	page	= request("page")
	makerid	= request("makerid")
	isusing	= request("isusing")
	
If page = ""	Then page = 1

SET lhello = new chello
	lhello.FCurrPage		= page
	lhello.FPageSize		= 20
	lhello.FRectMakerid		= makerid
	lhello.FRectIsusing		= isusing
	lhello.Frectcatecode = catecode
	lhello.FrectstandardCateCode = standardCateCode
	lhello.Frectbrandgubun		= brandgubun	
	lhello.Frectmduserid = mduserid	
	lhello.sbhelloList
%>
<script language="javascript">
function goHelloView(makerid){
	location.replace('/admin/brand/Hello/helloModify.asp?makerid='+makerid);
}
function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}
</script>
<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>Hello</b>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : 
		<%' Hello_ID_with_Name "makerid" ,makerid, " onchange='gosubmit("""");'"%>
		<% drawSelectBoxDesignerwithName "makerid",makerid %>
		&nbsp;&nbsp;
		* 브랜드구분 : <% drawSelectBoxbrandgubun "brandgubun",brandgubun , " onchange=""gosubmit('');""" %>		
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 대표카테고리 : 
		기능<% SelectBoxBrandCategory "catecode", catecode %>
		전시<%= fnStandardDispCateSelectBox(1,"", "standardCateCode", standardCateCode, "")%>
		&nbsp;&nbsp;
		* 담당MD : <% drawSelectBoxCoWorker_OnOff "mduserid", mduserid, "on" %>
		&nbsp;&nbsp;
		* 사용유무 : 
		<select name="isusing" class="select">
			<option value="">전체</option>
			<option value="Y" <%= Chkiif(isusing="Y", "selected", "") %>>Y</option>
			<option value="N" <%= Chkiif(isusing="N", "selected", "") %>>N</option>
		</select>		
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">	
	</td>
	<td align="right">	
		<input type="button" value="신규등록" onclick="location.replace('/admin/brand/Hello/helloModify.asp?mode=I');" class="button">
	</td>
</tr>	
</table>
<!-- 액션 끝 -->

<table width="100%", cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%=lhello.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= lhello.FTotalPage %></b>		
	</td>
</tr>
<input type= "hidden" name="makerid" value="<%=session("ssBctID")%>">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >브랜드ID</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >브랜드명(영문)</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >브랜드명(한글)</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >사용유무</td>
</tr>
<% If lhello.FResultcount > 0 Then %>
<% For i = 0 to lhello.FResultcount -1 %>

<% If lhello.FItemList(i).FIsusing="Y" Then %>
<tr height="25" bgcolor="FFFFFF"  align="center"  onclick="goHelloView('<%= lhello.FItemList(i).FUserid %>');" style="cursor:pointer;">
<% Else %>
<tr height="25" bgcolor="f1f1f1"  align="center"  onclick="goHelloView('<%= lhello.FItemList(i).FUserid %>');" style="cursor:pointer;">
<% End If %>	
	<td align="center"><%= lhello.FItemList(i).FUserid %></td>
	<td align="center"><%= lhello.FItemList(i).FSocname %></td>
	<td align="center"><%= lhello.FItemList(i).FSocname_kor %></td>
	<td align="center"><%= lhello.FItemList(i).FIsusing %></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If lhello.HasPreScroll Then %>
			<span class="lhello_link"><a href="javascript:gosubmit('<%= lhello.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + lhello.StartScrollPage to lhello.StartScrollPage + lhello.FScrollCount - 1 %>
			<% If (i > lhello.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(lhello.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="lhello_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If lhello.HasNextScroll Then %>
			<span class="lhello_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF">
	<td colspan="4" align="center" class="page_link">[검색결과가 없습니다.]</td>
</tr>
<% End If %>
</table>
<% Set lhello = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->