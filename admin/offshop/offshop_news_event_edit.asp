<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 오프라인 뉴스
' History : 최초생성자모름
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/board/offshop_newscls.asp" -->
<%
dim i, j, idx
	idx = requestCheckVar(request("idx"),10)

'==============================================================================
'나의 1:1질문답변
dim offnews
set offnews = New CNoticeDetail
offnews.GetOffshopNews idx

%>
<script>
function SubmitForm()
{
//alert('수정중입니다.');
//return;
        if (document.f.gubun.value == "") {
                alert("글 유형을 선택하세요.");
                return;
        }
        
        if (document.f.shopid.value == "") {
                alert("샵명을 선택하세요.");
                return;
        }
        
		if (document.f.title.value == "") {
                alert("제목을 입력하세요.");
                return;
        }
        if (document.f.contents.value == "") {
                alert("내용을 입력하세요.");
                return;
        }
        
        if (document.f.enddate.value == "") {
                alert("종료일을 입력하세요.");
                return;
        }
        
        if (confirm('수정 하시겠습니까?')){
            document.f.submit();
        }
}
</script>
<table  border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0" width="650" class="a">
<form method="post" name="f" action="<%= uploadImgUrl %>/linkweb/offshop/OffshopNewsEvent_process.asp" onsubmit="return false" enctype="multipart/form-data">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="idx" value="<% = idx %>">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<% ''사무실 직원만 수정가능 %>
<% if (session("ssBctDiv")<10) then %>
<input type="hidden" name="AssignFront" value="on">
<% end if %>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">샵명</td>
		<td bgcolor="white" style="padding:0">
			<% drawSelectBoxOffShopAll "shopid",offnews.Fshopid %>
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">글유형</td>
		<td bgcolor="white" style="padding:0">
			<select name="gubun">
				<option value="">선택</option>
			<%=fnOptCommonCode("noticegubun",offnews.Fgubun)%>
			</select>
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">제목</td>
		<td bgcolor="white" style="padding:0">
				<input name="title" style="width:450" maxlength="40" style="border:1 solid" value="<% = offnews.Ftitle %>">
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">내용</td>
		<td bgcolor="white" style="padding:0">
				<textarea name="contents" cols="50" rows="15" style="border:1 solid"><% = offnews.Fcontents %></textarea>
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">종료일</td>
		<td bgcolor="white" style="padding:0">
				<input type="text" name="enddate" size="12"  maxlength="10" style="border:1 solid" value="<% = offnews.Fenddate %>">
		        <a href="javascript:calendarOpen(f.enddate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
				(<%= Left(now(),10) %>)
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">첨부사진</td>
		<td bgcolor="white" style="padding:0">
				<input type="file" name="file1" size="50" class="input_b">
				<% if Not IsNULL(offnews.Ffile1) and (offnews.Ffile1<>"") then %>
				<img src="<%= offnews.Ffile1 %>">
				<% end if %>
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">사용여부</td>
		<td bgcolor="white" style="padding:0">
				<input type="radio" name="isusing" value="Y" <% if offnews.Fisusing = "Y" then response.write "checked" %>>Y <input type="radio" name="isusing" value="N" <% if offnews.Fisusing = "N" then response.write "checked" %>>N
		</td>
	</tr>
	<tr>
		<td style="padding:0" colspan="2" align="right" bgcolor="white">
			<input type="button" value="Save" onclick="SubmitForm()" style="background-color:#dddddd; height:25; border:1 solid buttonface">&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
</form>
</table>
<% set offnews = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
