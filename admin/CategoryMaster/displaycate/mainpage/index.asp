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
	Dim cDisp, cMain, i, vCateCode, vPage, vStartdate, vCurrPage
	vCateCode = Request("catecode")
	vPage = Request("page")
	vStartdate = Request("startdate")
	vCurrPage = Request("cpg")
	If vCurrPage = "" Then vCurrPage = 1 End If
	
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
function goCateCode(c){
	frm1.submit();
}
function goCatePage(p){
	frm1.submit();
}
function searchFrm(p){
	frm1.cpg.value = p;
	frm1.submit();
}
function goStartDate(d,p){
	menuiframe.location.href = "/admin/CategoryMaster/displaycate/mainpage/templete.asp?catecode=<%=vCateCode%>&page="+p+"&startdate="+d+"";
}
function goIssuePop(){
	var issuepop = window.open("issue_update.asp?catecode=<%=vCateCode%>","issuepop","width=600,height=500, scrollbars=yes, resizable=yes");
	issuepop.focus();
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
<style type="text/css">
.box1 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#FFF8F8; padding:7px 10px;}
</style>
<div class="box1">
* 새로운 페이지를 만들때 <b>페이지 선택한 뒤 반영일 입력 후</b> 저장하세요. 그후 해당 항목을 업데이트 하시면 됩니다.<br>
* 모든 반영일에는 <b>반드시 1 페이지부터 생성</b>하셔야 합니다.<br>
* ISSUE UPDATE 관리 : <input type="button" value=" ISSUE UPDATE " onClick="goIssuePop()"><br>
</div>
<br>
<table class="a">
<tr>
	<td>
<form name="frm1" action="" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="cpg" value="1">
카테고리선택 : 
<%
If cDisp.FResultCount > 0 Then
	Response.Write "<select name=""catecode"" class=""select"" onChange=""goCateCode(this.value);"">" & vbCrLf
	Response.Write "<option value="""">선택</option>" & vbCrLf
	For i=0 To cDisp.FResultCount-1
		Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
	Next
	Response.Write "</select>&nbsp;&nbsp;&nbsp;"
	
	Response.Write "페이지 : <select name=""page"" class=""select"" onChange=""goCatePage(this.value);"">" & vbCrLf
	Response.Write "<option value="""">선택</option>" & vbCrLf
	For i=1 To 5
		Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(vPage)=CStr(i),"selected","") & ">"&i&" 페이지</option>" & vbCrLf
	Next
	Response.Write "</select>&nbsp;&nbsp;&nbsp;"
%>
		찾을반영일 : <input type="text" name="startdate" size="10" maxlength="10" style="border:1px solid black;" readonly value="<%=vStartDate%>">
		<a href="javascript:calendarOpenAA(frm1.startdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		[<a href="javascript:" onClick="frm1.startdate.value='';">비움</a>]
<%
End If
Set cDisp = Nothing
%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="search">
</form>
	</td>
</tr>
<tr>
	<td>
<%
If vCateCode <> "" Then
Set cMain = New cDispCateMain
cMain.FPageSize = 5
cMain.FCurrPage = vCurrPage
cMain.FRectCateCode = vCateCode
cMain.FRectPage = vPage
cMain.FRectStartDate = vStartDate
cMain.GetDispCateMainList	
If cMain.FResultCount > 0 Then
	Response.Write "<table class=a cellpadding=3 cellspacing=1 bgcolor=#999999><tr align=center bgcolor=#E6E6E6 height=20>"
	Response.Write "<td align=center width=50>idx</td><td align=center width=90>반영일</td><td align=center width=50>페이지</td><td align=center width=120>등록자</td><td align=center width=150>등록일</td></tr>"
	For i=0 To cMain.FResultCount-1
%>
	<tr bgcolor="FFFFFF" height="20" onClick="goStartDate('<%=cMain.FItemList(i).Fstartdate%>','<%=cMain.FItemList(i).Fpage%>');" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
		 <td align="center"><%=cMain.FItemList(i).FIdx%></td>
		 <td align="center"><%=cMain.FItemList(i).Fstartdate%></td>
		 <td align="center"><%=cMain.FItemList(i).Fpage%></td>
		 <td align="center"><%=cMain.FItemList(i).Fregusername%>(<%=cMain.FItemList(i).Freguserid%>)</td>
		 <td align="center"><%=cMain.FItemList(i).Fregdate%></td>
	</tr>
<%
	Next
%>
	<tr height="20" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if cMain.HasPreScroll then %>
			<a href="javascript:searchFrm('<%= cMain.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + cMain.StartScrollPage to cMain.FScrollCount + cMain.StartScrollPage - 1 %>
    			<% if i>cMain.FTotalpage then Exit for %>
    			<% if CStr(vCurrpage)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if cMain.HasNextScroll then %>
    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
	</table>
<%
End If
Set cMain = Nothing
End If
%>
	</td>
</tr>
<tr>
	<td>
<%
If vCateCode <> "" Then
%>
	<iframe name="menuiframe" id="menuiframe" src="/admin/CategoryMaster/displaycate/mainpage/templete.asp?catecode=<%=vCateCode%>&page=<%=vPage%>" width="950px" height="2150px" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<%
Else
	'Response.Write "<br><br>"
End If
%>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->