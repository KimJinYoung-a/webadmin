<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateMainCls.asp"-->
<%
	Dim cDisp, cMain, vCateCode, vCurrPage, i, vIdx, vStartDate, vEndDate, vImgURL, vLinkURL, vTitle, vSubCopy
	vIdx = Request("idx")
	vCateCode = Request("catecode")
	vCurrPage = Request("cpg")
	If vCurrPage = "" Then vCurrPage = 1 End If
	
	If vIdx <> "" Then
	SET cMain = New cDispCateMain
	cMain.FRectCateCode = vCateCode
	cMain.FPageSize = 5
	cMain.FCurrPage = vCurrPage
	cMain.FRectIdx = vIdx
	cMain.GetCateMainIssueList()
	If cMain.FTotalCount > 0 Then
		vStartDate	= cMain.FItemList(0).Fstartdate
		vEndDate	= cMain.FItemList(0).Fenddate
		vImgURL		= cMain.FItemList(0).Fimgurl
		vLinkURL	= cMain.FItemList(0).Flinkurl
		vTitle		= cMain.FItemList(0).Ftitle
		vSubCopy	= cMain.FItemList(0).Fsubcopy
	End IF
	SET cMain = Nothing
	End If
	
	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 1
	cDisp.GetDispCateList()
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
document.domain ="10x10.co.kr";
function goSearch(c,p,i){
	location.href = "<%=CurrURL()%>?catecode="+c+"&cpg="+p+"&idx="+i+"";
}
function goSaveIssue(){
	frm1.submit();
}
function goDeleteIssue(){
	if(confirm("선택한 항목이 완전 삭제 됩니다.\n정말 삭제하시겠습니까?") == true) {
		frm1.action.value = "delete";
		frm1.submit();
	}
}
function jsSetImg(){
	var popImg;
	popImg = window.open('pop_uploadimg.asp?catecode=<%=vCateCode%>','popImg','width=370,height=150');
	popImg.focus();
}
function calendarOpenAA(objTarget){
    if (typeof calPopup == "function"){
        var compname = 'document.' + objTarget.form.name + '.' + objTarget.name;
        calPopup(objTarget,'calendarPopup',20+80,0, compname,'');
    }else{
        var fName = objTarget.form.name;
        var sName = objTarget.name;
    	var winCal = window.open('/lib/common_cal.asp?in_domain=o&FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
    	winCal.focus();
    }
}
</script>

<form name="frm1" action="issue_update_proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="idx" value="<%=vIdx%>">
<input type="hidden" name="action" value="">
<table class=a cellpadding=3 cellspacing=1 bgcolor=#999999>
<% If vIdx <> "" Then %>
<tr>
	<td align=center bgcolor=#E6E6E6 height=25 width=100>idx</td>
	<td bgcolor="FFFFFF" width="400"><%=vIdx%></td>
</tr>
<% End If %>
<tr>
	<td align=center bgcolor=#E6E6E6 height=25 width=100>카테고리</td>
	<td bgcolor="FFFFFF" width="400">
	<%
	If cDisp.FResultCount > 0 Then
		Response.Write "<select name=""catecode"" class=""select"" onChange=""goSearch(this.value,'1','');"">" & vbCrLf
		Response.Write "<option value="""">선택</option>" & vbCrLf
		For i=0 To cDisp.FResultCount-1
			Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
		Next
		Response.Write "</select>&nbsp;&nbsp;&nbsp;"
	End If
	Set cDisp = Nothing
	%>
	</td>
</tr>
<tr>
	<td align=center bgcolor=#E6E6E6 height=25 width=100>게시기간</td>
	<td bgcolor="FFFFFF" width="400">
		<input type="text" name="startdate" size="10" maxlength="10" style="border:1px solid black;" readonly value="<%=vStartDate%>">
		<a href="javascript:calendarOpenAA(frm1.startdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		~
		<input type="text" name="enddate" size="10" maxlength="10" style="border:1px solid black;" readonly value="<%=vEndDate%>">
		<a href="javascript:calendarOpenAA(frm1.enddate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>
</tr>
<tr>
	<td align=center bgcolor=#E6E6E6 height=25 width=100>이미지(200x200)</td>
	<td bgcolor="FFFFFF" width="400">
		<input type="hidden" name="imgurl" value="<%=vImgURL%>">
		<input type="button" value=" 이미지선택 " onClick="jsSetImg();">
		<span id="imgspan"><% If vImgURL <> "" Then %><img src="<%=vImgURL%>" width="30" height="30"><% End If %></span>
	</td>
</tr>
<tr>
	<td align=center bgcolor=#E6E6E6 height=25 width=100>링크</td>
	<td bgcolor="FFFFFF" width="400">
		<input type="text" name="linkurl" value="<%=vLinkURL%>" size="50">
	</td>
</tr>
<tr>
	<td align=center bgcolor=#E6E6E6 height=25 width=100>타이틀</td>
	<td bgcolor="FFFFFF" width="400">
		<input type="text" name="title" value="<%=vTitle%>" size="50">
	</td>
</tr>
<tr>
	<td align=center bgcolor=#E6E6E6 height=25 width=100>카피</td>
	<td bgcolor="FFFFFF" width="400">
		<textarea name="subcopy" cols="50" rows="6"><%=vSubCopy%></textarea>
	</td>
</tr>
<% If vCateCode <> "" Then %>
<tr>
	<td bgcolor="FFFFFF" colspan="2" align="right" height=25>
		<input type="button" value=" 닫 기 " onclick="window.close();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value=" 삭 제 " onclick="goDeleteIssue();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value=" 새글쓰기 " onclick="goSearch('<%=vCateCode%>','1','');">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value=" 저 장 " onclick="goSaveIssue();">
	</td>
</tr>
<% End If %>
</table>
</form>
<%
If vCateCode <> "" Then
	Set cDisp = New cDispCateMain
	cDisp.FRectCateCode = vCateCode
	cDisp.FPageSize = 5
	cDisp.FCurrPage = vCurrPage
	cDisp.GetCateMainIssueList()
%>
<br>
<table class=a cellpadding=3 cellspacing=1 bgcolor=#999999>
<tr align=center bgcolor=#E6E6E6 height=20>
	<td align=center width=50>idx</td>
	<td align=center width=150>시작 ~ 종료</td>
	<td align=center width=340>타이틀</td>
</tr>
<%	If cDisp.FResultCount < 1 Then %>
		<tr bgcolor="FFFFFF" height="20" onClick="" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
			<td align="center" colspan="3">데이터가 없습니다.</td>
		</tr>
<%	Else
		For i = 0 To cDisp.FResultCount-1
%>
		<tr bgcolor="FFFFFF" height="20" onClick="goSearch('<%=vCateCode%>','<%=vCurrPage%>','<%=cDisp.FItemList(i).FIdx%>');" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
			<td align="center"><%=cDisp.FItemList(i).FIdx%></td>
			<td align="center"><%=cDisp.FItemList(i).Fstartdate%> ~ <%=cDisp.FItemList(i).Fenddate%></td>
			<td><%=cDisp.FItemList(i).Ftitle%></td>
		</tr>
<%	Next %>
		<tr height="20" bgcolor="FFFFFF">
			<td colspan="20" align="center">
				<% if cDisp.HasPreScroll then %>
				<a href="javascript:goSearch('<%=vCateCode%>','<%= cDisp.StartScrollPage-1 %>','')">[pre]</a>
	    		<% else %>
	    			[pre]
	    		<% end if %>
	
	    		<% for i=0 + cDisp.StartScrollPage to cDisp.FScrollCount + cDisp.StartScrollPage - 1 %>
	    			<% if i>cDisp.FTotalpage then Exit for %>
	    			<% if CStr(vCurrpage)=CStr(i) then %>
	    			<font color="red">[<%= i %>]</font>
	    			<% else %>
	    			<a href="javascript:goSearch('<%=vCateCode%>','<%= i %>','')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>
	
	    		<% if cDisp.HasNextScroll then %>
	    			<a href="javascript:goSearch('<%=vCateCode%>','<%= i %>','')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
			</td>
		</tr>
<%
	End If
	Set cDisp = Nothing
	rw "</table>"
End If
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->