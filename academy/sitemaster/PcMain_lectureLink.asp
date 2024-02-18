<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  핑거스 아카데미 PC메인 작가&강사 링크
' History : 2016-10-24 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/academy/PcMainLectureLinkCls.asp" -->
<%
	Dim oPcMainLectureLink, i , page , idx , startdate , titletext , contentstext, isusing, lectureid
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)
	startdate = requestCheckVar(request("startdate"),10)	'시작일
	titletext = requestCheckVar(request("titletext"),255)	''제목
	contentstext = requestCheckVar(request("contentstext"),255)	''설명
	isusing = requestCheckVar(request("isusing"),1)	''사용여부
	lectureid = requestCheckVar(request("lectureid"),32)	''강사/작가 ID
  	if titletext <> "" then
		if checkNotValidHTML(titletext) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if contentstext <> "" then
		if checkNotValidHTML(contentstext) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"

set oPcMainLectureLink = new CPcMainLectureLinkContents
	oPcMainLectureLink.FPageSize = 20
	oPcMainLectureLink.FCurrPage = page
	oPcMainLectureLink.FRecttitletext = titletext
	oPcMainLectureLink.FRectcontentstext = contentstext
	oPcMainLectureLink.FRectlectureid = lectureid
	oPcMainLectureLink.FRectisusing = isusing
	oPcMainLectureLink.fnGetPcMainLectureLinkList()
%>
<script type="text/javascript">
	function NextPage(page){
		frm.page.value = page;
		frm.submit();
	}

	function AddNewContents(idx){
		var popwin = window.open('/academy/sitemaster/poppcmainlecturelinkEdit.asp?idx=' + idx,'pcmainlecturelinkEdit','width=700,height=800,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	function jsSerach(){
		var frm;
		frm = document.frm;
		frm.target = "_self";
		frm.action ="PcMain_lectureLink.asp";
		frm.submit();
	}

	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//이미지 확대 새창으로 보여주기
	function showimage(img){
		var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
	}

</script>

<form name="frm" method="post" style="margin:0px;">	
<input type="hidden" name="page" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	
	<td align="left">
    사용구분
	<select name="isusing">
	<option value="">전체
	<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함
	<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함
	</select>
	&nbsp;&nbsp;&nbsp;
	작가/강사ID 검색 : <input type="text" name="lectureid" size=20 value="<%=lectureid%>" />
<!--	제목검색 : <input type="text" name="titletext" size=20 value="<%'=titletext%>" />-->
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onclick="javascript:jsSerach();">
	</td>

</tr>
</table>
</form>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="AddNewContents('0');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<font color="red">※최근 등록순으로 시작일이 오늘이거나 오늘보다 작은걸로 1개 뿌려짐</font>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="6">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				검색결과 : <b><%= oPcMainLectureLink.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> / <%=  oPcMainLectureLink.FTotalpage %></b>
			</td>
			<td align="right"></td>			
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="3%">idx</td>
	<td width="15%">작가/강사 ID</td>
	<td width="15%">제목</td>
	<td width="15%">설명</td>
	<td width="5%">시작일</td>
	<td width="5%">수정</td>
</tr>
<% if oPcMainLectureLink.FresultCount > 0 then %>
<% for i=0 to oPcMainLectureLink.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center"><%= oPcMainLectureLink.FItemList(i).Fidx %></td>
	<td align="center"><%= oPcMainLectureLink.FItemList(i).Flectureid %></td>
	<td align="center"><%= db2html(oPcMainLectureLink.FItemList(i).Ftitletext) %></td>
	<td align="center"><%= db2html(oPcMainLectureLink.FItemList(i).Fcontentstext) %></td>
	<td align="center"><%= left(oPcMainLectureLink.FItemList(i).Fstartdate,10) %></td>
	<td align="center"><input type="button" class="button" value="수정" onclick="AddNewContents('<%= oPcMainLectureLink.FItemList(i).Fidx %>');"/></td>
</tr>
<% Next %>
<tr>
	<td colspan="6" align="center" bgcolor="#FFFFFF">
	 	<% if oPcMainLectureLink.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oPcMainLectureLink.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oPcMainLectureLink.StartScrollPage to oPcMainLectureLink.FScrollCount + oPcMainLectureLink.StartScrollPage - 1 %>
			<% if i>oPcMainLectureLink.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oPcMainLectureLink.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="6" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>
<% set oPcMainLectureLink = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->