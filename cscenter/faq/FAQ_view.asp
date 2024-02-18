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
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/faq_cls.asp"-->
<%
	'// 변수 선언 //
	dim faqid
	dim page, searchDiv, searchKey, searchString, param

	dim ofaq, i, lp

	'// 파라메터 접수 //
	faqid = request("faqid")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")

	if page="" then page=1
	if searchKey="" then searchKey="titleLong"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수

	'// 내용 접수
	set ofaq = new Cfaq
	ofaq.FRectfaqid = faqid

	ofaq.GetFAQRead

%>
<script language="javascript">

	// 글삭제
	function GotofaqDel(){
		if (confirm('삭제 하시겠습니까?')){
            document.frm_trans.mode.value = "DEL";
			document.frm_trans.submit();
		}
	}
	
    // 사용전환
	function GotofaqUsing(){
		if (confirm('사용전환 하시겠습니까?')){
            document.frm_trans.mode.value = "USE";
			document.frm_trans.submit();
		}
	}

</script>
<!-- 보기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr height="25">
			<td align="left"><b>FAQ 상세 정보</b></td>
			<td align="right"><%=ofaq.FfaqList(0).Fregdate%>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">작성자</td>
	<td bgcolor="#FFFFFF"><%= ofaq.FfaqList(0).Fusername & "(" & ofaq.FfaqList(0).Fuserid & ")" %></td>
</tr>
<%	if Not(ofaq.FfaqList(0).FlastWorker="" or isNull(ofaq.FfaqList(0).FlastWorker)) then %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">최종수정</td>
	<td bgcolor="#FFFFFF"><%= ofaq.FfaqList(0).FlastWorkerName & "(" & ofaq.FfaqList(0).FlastWorker & ") / " & ofaq.FfaqList(0).FlastUpdate %></td>
</tr>
<%	end if %>
<tr>
	<td align="center" bgcolor="#DDDDFF">구분</td>
	<td bgcolor="#FFFFFF"><%= db2html(ofaq.FfaqList(0).Fcomm_name) %></td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">표시순서</td>
	<td bgcolor="#FFFFFF"><%= ofaq.FfaqList(0).Fdisporder %></td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">제목</td>
	<td bgcolor="#F8F8FF"><%= ReplaceBracket(db2html(ofaq.FfaqList(0).Ftitle)) %></td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">내용</td>
	<td bgcolor="#FFFFFF"><%= nl2br(ReplaceBracket(db2html(ofaq.FfaqList(0).Fcontents))) %></td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">Link명</td>
	<td bgcolor="#FFFFFF"><%= ReplaceBracket(db2html(ofaq.FfaqList(0).Flinkname)) %></td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">LinkURL</td>
	<td bgcolor="#FFFFFF">
        <% if ofaq.FfaqList(0).Flinkurl<>"" then %>
    	<%= ReplaceBracket(db2html(ofaq.FfaqList(0).Flinkurl)) %>
    	&nbsp;&nbsp;
    	<a href="<%= db2html(ofaq.FfaqList(0).Flinkurl) %>" target="_blank"><font color="blue">>><%= db2html(ofaq.FfaqList(0).Flinkname) %> 바로가기</font></a>
    	<% end if %>
    </td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">사용여부</td>
	<td bgcolor="#FFFFFF">
        <%= ofaq.FfaqList(0).fisusing %>
    </td>
</tr>
<tr>
	<td colspan="2" height="30" bgcolor="#FAFAFA" align="center">
		<input type="button" class="button" value="수정" onClick="self.location='faq_modi.asp?menupos=<%=menupos%>&faqid=<%=faqid & param%>'"> &nbsp;
		<% if ofaq.FfaqList(0).Fisusing = "Y" then %>
		<input type="button" class="button" name="mode" value="삭제" onClick="GotofaqDel()"> &nbsp;
		<% else %>
		<input type="button" class="button" name="mode" value="사용전환" onClick="GotofaqUsing()"> &nbsp;
	    <% end if %>
		<input type="button" class="button" value="리스트" onClick="self.location='faq_list.asp?menupos=<%=menupos & param %>'">
	</td>
</tr>
<form name="frm_trans" method="POST" action="faq_process.asp">
<input type="hidden" name="faqid" value="<%=faqid%>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
</form>


</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->