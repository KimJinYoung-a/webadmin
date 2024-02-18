<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%

dim yyyy1,mm1,nowdate
dim yyyy2,mm2,dd2
dim lecturer
dim lec_idx, lec_title, lecturdate
dim lecturdateyn
dim page

lec_idx = RequestCheckvar(request("lec_idx"),10)
lecturer = RequestCheckvar(request("lecturer"),32)
lec_title = request("lec_title")
page = RequestCheckvar(request("page"),10)
if page="" then page=1
if lec_title <> "" then
	if checkNotValidHTML(lec_title) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
nowdate = now()

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1   = RequestCheckvar(request("mm1"),2)

if yyyy1="" then
	yyyy1 = Left(Cstr(nowdate),4)
	mm1	  = Mid(Cstr(nowdate),6,2)
end if

lecturdateyn = request("lecturdateyn")
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2   = RequestCheckvar(request("mm2"),2)
dd2   = RequestCheckvar(request("dd2"),2)

if yyyy2="" then
	yyyy2 = Left(Cstr(nowdate),4)
	mm2	  = Mid(Cstr(nowdate),6,2)
	dd2	  = Mid(Cstr(nowdate),9,2)
end if
lecturdate = yyyy2 + "-" + mm2 + "-" + dd2

dim olecture
set olecture = new CLecture
olecture.FCurrPage = page
olecture.FPageSize=20
olecture.FRectidx = lec_idx
olecture.FRectLecturer = lecturer
olecture.FRectSearchTitle = lec_title

if lecturdateyn="on" then
	olecture.FRectSearchLectureDay = lecturdate
end if

olecture.GetLectureList

dim i
%>
<script language='javascript'>
function SelectLec(lecidx,lecname,lecturer){
	opener.SelectLecture(lecidx,lecname,lecturer);
}

function NextPage(page){
	frm.page.value= page;
	frm.submit();
}

function GetOnload(){
	ckEnabled(frm.lecturdateyn);
}

function ckEnabled(comp){
	frm.yyyy2.disabled = (!comp.checked);
	frm.mm2.disabled = (!comp.checked);
	frm.dd2.disabled = (!comp.checked);
}

function popLecSimpleEdit(lec_idx){
	popwin = window.open('/academy/lecture/poplecsimpleedit.asp?lec_idx=' + lec_idx,'popLecSimpleEdit','width=600,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popLecReg(lec_idx){
	popwin = window.open('/academy/lecture/poplecreg.asp?lec_idx=' + lec_idx,'popLecReg','width=600,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

window.onload = GetOnload;

</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method=get>
	<input type="hidden" name="page" >
    <tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="30" >
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
			강사 : <% drawSelectBoxLecturer "lecturer",lecturer  %>
			<input type="checkbox" name="lecturdateyn" <% if lecturdateyn = "on" then response.write "checked" %> onclick="ckEnabled(this)">
			강좌일 : <% DrawOneDateBox2 yyyy2,mm2,dd2 %>
        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22"  border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#DDDDFF">
		<td colspan="14" align="right">검색건수 : <%= olecture.FTotalCount %> 건 Page : <%= page %>/<%= olecture.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="50">이미지</td>
		<td align="center" width="50">강좌코드</td>
		<td align="center">강좌명</td>
		<td align="center" width="70">강사명</td>
		<td align="center" width="80">강좌(시작)일</td>
		<td align="center" width="70">접수기간</td>
		<td align="center" width="50">수강료</td>
		<td align="center" width="40">정원</td>
		<td align="center" width="40">신청인원(웹상)</td>
		<td align="center" width="40">마감<br>여부</td>
		<td align="center" width="40">선택</td>
	</tr>
<% for i=0 to olecture.FResultCount - 1 %>
	<% if olecture.FItemList(i).FIsUsing="N" then %>
	<tr align="center" bgcolor="#EEEEEE">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><img src="<%= olecture.FItemList(i).Fsmallimg %>" width="50" border="0"></td>
		<td><%= olecture.FItemList(i).Fidx %></td>
		<td><%= olecture.FItemList(i).Flec_title %></td>
		<td><%= olecture.FItemList(i).Flecturer_id %><br>(<%= olecture.FItemList(i).Flecturer_name %>)</td>
		<td><%= olecture.FItemList(i).Flec_startday1 %></td>
		<td align="center"><%= olecture.FItemList(i).Freg_startday %><br>~<br><%= olecture.FItemList(i).Freg_endday %></td>
		<td align="right"><%= FormatNumber(olecture.FItemList(i).Flec_cost,0) %></td>
		<td><%= olecture.FItemList(i).Flimit_count %></td>
		<td><%= olecture.FItemList(i).Flimit_sold %></td>
		<td>
			<% if olecture.FItemList(i).IsSoldOut then %>
			<font color="#CC3333">마감</font>
			<% end if %>
		</td>
		<td><input type="button" value="선택" onclick="SelectLec('<%= olecture.FItemList(i).Fidx %>','<%= replace(olecture.FItemList(i).Flec_title,"'","") %>','<%= olecture.FItemList(i).Flecturer_id %>');"></td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="14" height="30" align="center">
		<% if olecture.HasPreScroll then %>
			<a href="javascript:NextPage('<%= olecture.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + olecture.StarScrollPage to olecture.FScrollCount + olecture.StarScrollPage - 1 %>
			<% if i>olecture.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if olecture.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<%
set olecture = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->