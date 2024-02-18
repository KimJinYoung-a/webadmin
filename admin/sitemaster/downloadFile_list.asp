<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/downloadFileCls.asp"-->
<%
'###############################################
' PageName : downloadFile_List.asp
' Discription : 파일 다운로드 관리 목록
'           2012.04.04 정윤정 이벤트코드 추가
'           2014.05.09 허진원 직접링크 복사기능 추가
'###############################################

dim page, i, lp

page = requestCheckvar(request("page"),10)
if page = "" then page=1

dim oFile
set oFile = New cDownFile
oFile.FCurrPage = page
oFile.FPageSize=20
oFile.FRectUsing = "Y"
oFile.GetfileList

%>
 
<script type="text/javascript">
// 페이지 이동
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="downloadFile_list.asp";
	document.refreshFrm.submit();
}

 //파일 다운로드 스크립트 생성
function copyScrt(vSn) {
	var doc = "javascript:fileDownload(" + vSn + ");";
	copyStringToClipboard(doc);
	alert('선택하신 파일의 다운로드 스크립트가 복사되었습니다. 사용하실 곳에 Ctrl+V 하시면됩니다.\n\n※자바스크립트이므로 링크는 텐바이텐 사이트 내에서만 쓸 수 있습니다.');
} 

 //파일 다운로드 직접링크 생성
function copyLink(vSn) {
	var doc = "http://upload.10x10.co.kr/linkweb/download/fileDownload.asp?fn=" + vSn;
	copyStringToClipboard(doc);
	alert('선택하신 파일의 다운로드 링크가 복사되었습니다. 사용하실 곳에 Ctrl+V 하시면됩니다.\n\n※대 고객용으로는 스크립트를 이용하시고 이 링크는 사용하지 마세요.');
} 
</script> 
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right"><input type="button" value="파일 추가" onclick="self.location='downloadFile_Write.asp?mode=add&menupos=<%= menupos %>'" class="button"></td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%=oFile.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oFile.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>이벤트코드</td>
	<td>제목</td>
	<td>파일명</td>
	<td>크기</td>
	<td>다운로드</td>
	<td>등록일</td>
	<td>&nbsp;</td>
</tr>
<%	if oFile.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 파일이 없습니다.</td>
</tr>
<%
	else
		for i=0 to oFile.FResultCount-1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><a href="downloadFile_Write.asp?mode=edit&menupos=<%= menupos %>&fileSn=<%= oFile.FItemList(i).FfileSn %>"><%= oFile.FItemList(i).FfileSn %></a></td> 
	<td align="center"><a href="downloadFile_Write.asp?mode=edit&menupos=<%= menupos %>&fileSn=<%= oFile.FItemList(i).FfileSn %>"><%= oFile.FItemList(i).Fevt_code %></a></td> 
	<td align="center"><a href="downloadFile_Write.asp?mode=edit&menupos=<%= menupos %>&fileSn=<%= oFile.FItemList(i).FfileSn %>"><%= oFile.FItemList(i).FfileTitle %></a></td>
	<td align="center"><%= oFile.FItemList(i).FfileDownNm & "<br>(" & oFile.FItemList(i).FfileName & ")"%></td>
	<td align="center">
	<%
		if oFile.FItemList(i).FfileSize >= 1048576 then
			Response.Write FormatNumber(oFile.FItemList(i).FfileSize/1024/1024,1) & "MBytes"
		elseif oFile.FItemList(i).FfileSize >= 1024 then
			Response.Write FormatNumber(oFile.FItemList(i).FfileSize/1024,0) & "KBytes"
		else
			Response.Write FormatNumber(oFile.FItemList(i).FfileSize,0) & "Bytes"
		end if
	%>
	</td>
	<td align="center"><%= oFile.FItemList(i).FdownCount & "회<br>" & left(oFile.FItemList(i).FlastDownDate,10) %></td>
	<td align="center"><%= left(oFile.FItemList(i).Fregdate,10) %></td>
	<td align="center"> 
		<input type="button"  id="btnLink" class="button" value="스크립트 복사" title="텐바이텐 사이트용 다운로드 스크립트 복사" onClick="copyScrt('<%=oFile.FItemList(i).FfileSn %>')"><br>
		<input type="button"  id="btnLink" class="button" value="직접링크 복사" title="내부 공유용 직접 다운로드 링크 복사" onClick="copyLink('<%=oFile.FItemList(i).FfileSn %>')">
	</td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<!-- 페이지 시작 -->
	<%
		if oFile.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oFile.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oFile.StartScrollPage to oFile.FScrollCount + oFile.StartScrollPage - 1

			if lp>oFile.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oFile.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
</tr>
</table>
<form name="refreshFrm" method="get" action="downloadFile_list.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
</form>
<%
set oFile = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->