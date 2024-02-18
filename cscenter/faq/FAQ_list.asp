<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [CS]각종설정>>[FAQ]관리 
' Hieditor : 2009.03.02 이영진 생성
'			 2021.07.30 한용민 수정(사용여부 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/cscenter/faq_cls.asp"-->
<%
	'// 변수 선언 //
	dim faqid
	dim page, searchDiv, searchKey, searchString, param

	dim ofaq, i, lp, bgcolor, strUsing,isusing


	'// 파라메터 접수 //
	faqid = request("faqid")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")
	isusing = requestcheckvar(request("isusing"),1)

	if page="" then page=1
	if searchKey="" then searchKey="title"

	param = "&searchKey=" & searchKey & "&searchString=" & searchString & "&searchDiv=" & searchDiv

	'// 클래스 선언
	set ofaq = new Cfaq
	ofaq.FCurrPage = page
	ofaq.FPageSize = 20
	ofaq.FRectsearchDiv = searchDiv
	ofaq.FRectsearchKey = searchKey
	ofaq.FRectsearchString = searchString
	ofaq.FRectisusing = isusing
	ofaq.GetFAQList
%>
<script language='javascript'>
<!--
	function chk_form(){
		var frm = document.frm_search;

//		if(!frm.searchKey.value){
//			alert("검색 조건을 선택해주십시오.");
//			frm.searchKey.focus();
//			return;
//		}
//		else if(!frm.searchString.value)
//		{
//			alert("검색어를 입력해주십시오.");
//			frm.searchString.focus();
//			return;
//		}

		frm.submit();
	}

	function goPage(pg)
	{
		var frm = document.frm_search;

		frm.page.value= pg;
		frm.submit();
	}
//-->
</script>
<!-- 검색 시작 -->
<form name="frm_search" method="POST" action="faq_list.asp" onSubmit="return false">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			구분
    		<select name="searchDiv" class="select" onChange="goPage(frm_search.page.value)">
    			<option value="">선택</option>
    			<%= db2html((ofaq.optCommCd("Z200", searchDiv))) %>
    		</select>
    		/ 검색
    		<select name="searchKey" class="select">
    			<option value="">선택</option>
    			<option value="title">제목+내용</option>
    		</select>
    		<script language="javascript">
    			document.frm_search.searchKey.value="<%=searchKey%>";
    		</script>
    		<input type="text" class="text" name="searchString" size="20" value="<%= searchString %>">
			/ 사용여부 : <% drawSelectBoxUsingYN "isusing", isusing %>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="chk_form()">
		</td>
	</tr>
</table>
</form>
<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onClick="location.href='faq_write.asp?menupos=<%=menupos%>'">			
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="11">
			검색결과 : <b><%= ofaq.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ofaq.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="40">IDX</td>
		<td width="40">순서</td>
		<td width="60">구분코드</td>
		<td width="120">구분</td>
		<td>제목</td>
		<td width="50">LinkURL</td>
		<td>Link명</td>
		<td width="70">등록자</td>
		<td width="50">조회수</td>
		<td width="80">등록일</td>
		<td width="50">사용유무</td>
	</tr>
	<%
		for lp=0 to ofaq.FResultCount - 1
	%>
	<tr align="center" <% if ofaq.FfaqList(lp).Fisusing = "N" then %> bgcolor="#EEEEEE" <% else %> bgcolor="#FFFFFF" <% end if %> >
	    <td><%= ofaq.FfaqList(lp).Ffaqid %></td>
		<td><%= ofaq.FfaqList(lp).Fdisporder %></td>
		<td><%= ofaq.FfaqList(lp).FcommCd %></td>
		<td align="left"><%= db2html(ofaq.FfaqList(lp).Fcomm_name) %></td>
		<td align="left">
			<a href="faq_view.asp?faqid=<%= ofaq.FfaqList(lp).Ffaqid %>&page=<%=page & param%>&menupos=<%=menupos%>">
			<%= ReplaceBracket(db2html(ofaq.FfaqList(lp).Ftitle)) %></a>
		</td>
	    <td>
	        <% if ofaq.FfaqList(lp).Flinkurl<>"" then %>
	        <acronym title="<%= ReplaceBracket(db2html(ofaq.FfaqList(lp).Flinkurl)) %>">YES</acronym>
	        <% end if %>
	    </td>
	    <td><a href="<%= ReplaceBracket(db2html(ofaq.FfaqList(lp).Flinkurl)) %>" target="_blank"><%= ReplaceBracket(ofaq.FfaqList(lp).Flinkname) %></a></td>
		<td><%= ofaq.FfaqList(lp).Fuserid %></td>
		<td><%= ofaq.FfaqList(lp).FhitCount %></td>
		<td><%= FormatDate(ofaq.FfaqList(lp).Fregdate,"0000.00.00") %></td>
	    <td><%= ofaq.FfaqList(lp).Fisusing %></td>
	</tr>
	<%
		next
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" height="30" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
					<% sbDisplayPaging "page="&page, ofaq.FTotalCount, ofaq.FPageSize, 10%>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<%
set ofaq = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
