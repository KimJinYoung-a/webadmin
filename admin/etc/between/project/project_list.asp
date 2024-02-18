<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<%
Dim opjt, i, page
Dim pjt_kind, selPjt, sPtxt, pjt_state, pjt_gender, isusing
page    = request("page")
If page = "" Then page = 1

pjt_kind	= request("pjt_kind")
selPjt		= request("selPjt")
sPtxt		= request("sPtxt")
pjt_state	= request("pjt_state")
pjt_gender	= request("pjt_gender")
isusing		= request("isusing")

SET opjt = new cProject
	opjt.FPageSize 					= 20
	opjt.FCurrPage					= page
	opjt.FRectPjt_kind				= pjt_kind
	opjt.FRectSelPjt				= selPjt
	opjt.FRectSPtxt					= sPtxt
	opjt.FRectPjt_state				= pjt_state
	opjt.FRectPjt_gender			= pjt_gender
	opjt.FRectIsusing				= isusing	
	opjt.getProjectList()
%>
<script language="javascript">
function jsGoUrl(sUrl){
	self.location.href = sUrl;
}
function goPage(page){
    var frm = document.frmpjt;
    frm.page.value=page;
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmpjt" method="get"  action="<%= CurrURL %>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>
					기획전 구분: <% sbGetOptProjectCodeValue "pjt_kind",pjt_kind,"" %> &nbsp;&nbsp;
					코드/명:
					<select name="selPjt" class="select">
						<option value="">- Choice -</option>
				    	<option value="pjt_code" <%= Chkiif(selPjt="pjt_code", "selected", "")%> >기획전코드</option>
				    	<option value="pjt_name" <%= Chkiif(selPjt="pjt_name", "selected", "")%> >기획전명</option>
			    	</select>
					<input type="text" name="sPtxt" value="<%=sPtxt%>" maxlength="60">
			        &nbsp;&nbsp;
					진행상태:
		   			<select class="select" name="pjt_state">
		   				<option value="">- Choice -</option>
		   				<option value="0" <%= Chkiif(pjt_state="0", "selected", "")%> >등록대기</option>
		   				<option value="7" <%= Chkiif(pjt_state="7", "selected", "")%> >오픈</option>
		   			</select>
		   			&nbsp;&nbsp;
					성별:
		   			<select name="pjt_gender" class="select">
		   				<option value="">- Choice -</option>
		   				<option value="A" <%= Chkiif(pjt_gender="A", "selected", "")%> >전체</option>
		   				<option value="M" <%= Chkiif(pjt_gender="M", "selected", "")%> >남자</option>
		   				<option value="F" <%= Chkiif(pjt_gender="F", "selected", "")%> >여자</option>
		   			</select>
		   			&nbsp;&nbsp;
					사용유무
		   			<select name="isusing" class="select">
		   				<option value="">- Choice -</option>
		   				<option value="Y" <%= Chkiif(isusing="Y", "selected", "")%> >Y</option>
		   				<option value="N" <%= Chkiif(isusing="N", "selected", "")%> >N</option>
		   			</select>
				</td>
			</tr>
			</table>
        </td>
    		<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<a href="javascript:document.frmpjt.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
</form>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
    <tr height="40" valign="bottom">
        <td align="left">
        	<input type="button" value="새로등록" onclick="jsGoUrl('project_regist.asp?menupos=<%=menupos%>');" class="button">
	    </td>
	</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="9">검색결과 : <b><%= FormatNumber(opjt.FTotalCount,0) %></b>&nbsp;&nbsp;페이지 : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(opjt.FTotalPage,0) %></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap>기획전코드</td>
		<td nowrap>기획전구분</td>
		<td nowrap>기획전명</td>
		<td nowrap>성별</td>
		<td nowrap>진행상태</td>
		<td nowrap>배너이미지</td>
		<td nowrap>관리</td>
    </tr>
<%
	If opjt.FResultCount > 0 Then
		For i = 0 to opjt.FResultCount - 1
%>
    <tr align="center" <%= Chkiif(opjt.FItemList(i).FPjt_using="Y", "bgcolor=#FFFFFF", "bgcolor=#BFBFBF") %>  height="30">
    	<td><a href="/admin/etc/between/project/project_modify.asp?mode=U&pjt_code=<%=opjt.FItemList(i).FPjt_code%>&menupos=<%=menupos%>"><%= opjt.FItemList(i).FPjt_code %></a></td>
		<td><a href="/admin/etc/between/project/project_modify.asp?mode=U&pjt_code=<%=opjt.FItemList(i).FPjt_code%>&menupos=<%=menupos%>"><%= getDBcodeByName(opjt.FItemList(i).FPjt_kind) %></a></td>
		<td><a href="/admin/etc/between/project/project_modify.asp?mode=U&pjt_code=<%=opjt.FItemList(i).FPjt_code%>&menupos=<%=menupos%>"><%= opjt.FItemList(i).FPjt_name %></a></td>
		<td>
		<%
			Select Case opjt.FItemList(i).FPjt_gender
				Case "A"	response.write "전체"
				Case "M"	response.write "남자"
				Case "F"	response.write "여자"
			End Select
		%>
		</td>
		<td><%= getDBcodeByName(opjt.FItemList(i).FPjt_state) %></td>
		<td><a href="/admin/etc/between/project/project_modify.asp?mode=U&pjt_code=<%=opjt.FItemList(i).FPjt_code%>&menupos=<%=menupos%>"><img src="<%= opjt.FItemList(i).FPjt_topImgUrl %>" width="100" border="0"></a></td>
		<td align="center" nowrap>
			<input type="button" value="상품" class="button" onClick="javascript:jsGoUrl('projectitem_regist.asp?pjt_code=<%=opjt.FItemList(i).FPjt_code%>&menupos=<%=menupos%>')">
		</td>
    </tr>
<%
		Next
%>
	<tr height="20">
	    <td colspan="17" align="center" bgcolor="#FFFFFF">
	        <% if opjt.HasPreScroll then %>
			<a href="javascript:goPage('<%= opjt.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + opjt.StartScrollPage to opjt.FScrollCount + opjt.StartScrollPage - 1 %>
	    		<% if i>opjt.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if opjt.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<%
	Else
%>
   	<tr height="50" align="center" bgcolor="#FFFFFF">
   		<td colspan="11">등록된 내용이 없습니다.</td>
   	</tr>
<% End If %>
</table>
<% SET opjt = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->