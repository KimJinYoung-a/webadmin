<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 고객안내FAQ
' Hieditor : 2019.10.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/board/faq/customer_manual_faq_cls.asp"-->

<%
Dim ofaq,i,page, isusing, reloading, fidx, gubun, contents, solution, manualtype
	page = requestCheckVar(getNumeric(request("page")),10)
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
    isusing = requestCheckVar(request("isusing"),1)
    reloading = requestCheckVar(request("reloading"),2)
    fidx = requestCheckVar(getNumeric(request("fidx")),10)
    gubun = requestCheckVar(getNumeric(request("gubun")),10)
    contents = requestCheckVar(request("contents"),128)
    solution = requestCheckVar(request("solution"),128)

if page = "" then page = 1
if reloading="" and isusing="" then isusing="Y"
manualtype="customer_faq"

set ofaq = new cfaq_list
	ofaq.FPageSize = 100
	ofaq.FCurrPage = page
    ofaq.frectisusing = isusing
    ofaq.frectfidx = fidx
    ofaq.frectgubun = gubun
	ofaq.frectmanualtype = manualtype
    ofaq.frectcontents = contents
    ofaq.frectsolution = solution
	ofaq.Getcustomer_manual_faq()
%>

<script type="text/javascript">

function fnfaq_reg(fidx){
	var reg = window.open('/cscenter/board/faq/customer_manual_faq_edit.asp?menupos=<%=menupos%>&fidx='+fidx,'reg','width=1280,height=700,scrollbars=yes,resizable=yes');
	reg.focus();	
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<input type="hidden" name="reloading" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        * 사용여부 : <% drawSelectBoxisusingYN "isusing",isusing,"" %>
        &nbsp;
        * 구분 : <% Drawcustomerfaqgubun "gubun",gubun,"" %>
        <Br><Br>
        * 번호 : <input type="text" name="fidx" value="<%= fidx %>" size=11 maxlength=10>
        &nbsp;
        * 문의내용 : <input type="text" name="contents" value="<%= contents %>" size=50 maxlength=48>
        &nbsp;
        * 처리방법 : <input type="text" name="solution" value="<%= solution %>" size=50 maxlength=48>
	</td>	
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">	
		<input type="button" class="button" value="신규등록" onclick="fnfaq_reg('');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ofaq.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ofaq.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>구분</td>
	<td>문의내용</td>
	<td>처리방법</td>
    <td>사용여부</td>
	<td>최종수정</td>
	<td>비고</td>
</tr>
<% if ofaq.FresultCount>0 then %>
<% for i=0 to ofaq.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
	<td width=60>
		<%= ofaq.FItemList(i).ffidx %>
	</td>		
	<td width=200>
		<%= getcustomerfaqgubunname(ofaq.FItemList(i).fgubun) %>
	</td>
	<td align="left" width=400>
		<%= nl2br(ReplaceBracket(ofaq.FItemList(i).fcontents)) %>
	</td>	
	<td align="left">
		<%= nl2br(ReplaceBracket(ofaq.FItemList(i).fsolution)) %>
	</td>
	<td width=50>
		<%= ofaq.FItemList(i).fisusing %>
	</td>
	<td width=80>
		<%= left(ofaq.FItemList(i).flastupdate,10) %>
        <br><%= mid(ofaq.FItemList(i).flastupdate,11,22) %>
        <% if ofaq.FItemList(i).flastadminid<>"" then %>
            <br><%= ofaq.FItemList(i).flastadminid %>
        <% end if %>
	</td>
	<td width=40>
		<input type="button" value="수정" class="button" onclick="fnfaq_reg('<%= ofaq.FItemList(i).ffidx %>');">
	</td>	
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ofaq.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= ofaq.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ofaq.StartScrollPage to ofaq.StartScrollPage + ofaq.FScrollCount - 1 %>
			<% if (i > ofaq.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ofaq.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ofaq.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<%
set ofaq = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->