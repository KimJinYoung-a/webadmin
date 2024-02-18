<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateMenuCls.asp"-->

<%
	Dim cMenu, vArr, i, vDisp1, vType, vPage, vUseYN, vOrderBy
	vPage = Request("page")
	vDisp1 = Request("disp1")
	vUseYN = Request("useyn")
	vOrderBy = Request("orderby")
	
	If vPage = "" Then vPage = "1" End If
	If vUseYN = "" Then vUseYN = "y" End If
	If vOrderBy = "" Then
		if (session("ssAdminPsn")="11" or session("ssAdminPsn")="21") then	'MD라면 정렬번호순
			vOrderBy = "sortno asc, idx desc"
		else	'그외 등록순
			vOrderBy = "idx desc, sortno asc"
		end if
	End If
	
	
	Set cMenu = New cDispCateMenu
	vArr = cMenu.GetDispCate1Depth()
	Set cMenu = Nothing
	
	Set cMenu = New cDispCateMenu
	cMenu.FCurrPage = vPage
	cMenu.FDisp1 = vDisp1
	cMenu.FType = "topbanner"
	cMenu.FUseYN = vUseYN
	cMenu.FOrderBy = vOrderBy
	cMenu.GetCateMainIssueList
%>

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script>
function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function popreg(idx){
    var poprreg = window.open('reg.asp?disp1=<%=vDisp1%>&idx='+idx+'','poprreg','width=770,height=600');
    poprreg.focus();
}
function Check_All()
{
	var chk = f.idx;
	alert(chk.length);
	var cnt = 0;
	var ischecked = ""
	if(document.getElementById("chkall").checked){
		ischecked = "checked"
	}else{
		ischecked = ""
	}
	
	if(cnt == 0 && chk.length != 0){
		for(i = 0; i < chk.length; i++){ chk.item(i).checked = ischecked; }
		cnt++;
	}
}
function jsUseYNNO(){
	if(confirm("선택한 것들을 삭제하시겠습니까?") == true) {
		f.submit();
	}
}
function jsRealServerReg(){
	if(confirm("[<%=vDisp1%>] 카테고리 메뉴를 생성하시겠습니까?") == true){
	    var popCreateTemp = window.open("http://<%=CHKIIF(application("Svr_Info")="Dev","2015www","www1")%>.10x10.co.kr/chtml/dispcate/menu_make_xml_New.asp?catecode=<%=vDisp1%>","popCreateTemp","width=1200 height=930 scrollbars=yes resizable=yes");
		popCreateTemp.focus();
	}
}
</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="type" value="topbanner">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="40">
	<td align="left">
		<font color="blue" size="2"><strong>※ 2015 PC 리뉴얼부터 사용 될 메뉴입니다.</strong></font><br />
		&nbsp;
		<select name="disp1" class="select" onChange="frm.submit();">
		<option value="">-카테고리선택-</option>
		<%
			For i=0 To UBound(vArr,2)
				Response.Write "<option value='" & vArr(0,i) & "' " & CHKIIF(CStr(vDisp1)=CStr(vArr(0,i)),"selected","") & ">" & vArr(1,i) & "</option>" & vbCrLf
			Next
		%>
		</select>
		&nbsp;&nbsp;&nbsp;
		<select name="useyn" class="select" onChange="frm.submit();">
			<option value="">-사용여부선택-</option>
			<option value="y" <%=CHKIIF(vUseYN="y","selected","")%>>사용</option>
			<option value="n" <%=CHKIIF(vUseYN="n","selected","")%>>사용안함</option>
		</select>
		&nbsp;&nbsp;&nbsp;
		<select name="orderby" class="select" onChange="frm.submit();">
			<option value="sortno asc, idx desc" <%=CHKIIF(vOrderBy="sortno asc, idx desc","selected","")%>>정렬번호순</option>
			<option value="idx desc, sortno asc" <%=CHKIIF(vOrderBy="idx desc, sortno asc","selected","")%>>최근등록순</option>
		</select>
	</td>
</tr>
</table>
</form>
<br>
<% If vDisp1 <> "" Then %>
	<font color="red" size="3"><strong>※ 정렬번호가 가장 작은것, 가장 최근글(idx숫자가 높은것) 1개가 적용됩니다.</strong></font><br />
	<input type="button" value="[<%=vDisp1%>]카테고리 메뉴 적용하기" onClick="jsRealServerReg();">
<br>
<% End If %>
<br>
<form name="f" method="post" action="proc.asp" target="ifram">
<input type="hidden" name="action" value="del">
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="40" bgcolor="FFFFFF">
	<td colspan="12">
		<table width="100%" class="a">
		<tr>
			<td>
				검색결과 : <b><%=cMenu.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= vPage %> / <%=cMenu.FTotalPage%></b>
				<br>
				<input type="button" value="체크한것 사용안함 처리" onClick="jsUseYNNO();">
			</td>
			<td align="right">
				<input type="button" value="신규등록" onClick="popreg('');">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td><input type="checkbox" name="chkall" id="chkall" value="" onClick="Check_All()"></td>
    <td>idx</td>
    <td>전시카테고리</td>
    <td>이미지</td>
    <td>사용여부</td>
    <td>정렬번호</td>
    <td>등록자</td>
    <td>등록일</td>
    <td></td>
</tr>
<%
	for i=0 to cMenu.FResultCount - 1
%>
<% if cMenu.FItemList(i).Fuseyn = "n" then %>
<tr height="30" bgcolor="#DDDDDD">
<% else %>
<tr height="30" bgcolor="#FFFFFF">
<% end if %>
	<td align="center"><input type="checkbox" name="idx" value="<%= cMenu.FItemList(i).FIdx %>"></td>
	<td align="center"><%= cMenu.FItemList(i).FIdx %></td>
	<td align="center"><a href="?disp1=<%=cMenu.FItemList(i).Fdisp1%>&useyn=<%=vUseYN%>&orderby=<%=vOrderBy%>"><%= cMenu.FItemList(i).Fcatename %></a></td>
	<td align="center"><img src="<%= cMenu.FItemList(i).FimgurlReal %>" height="70"></td>
	<td align="center"><%= cMenu.FItemList(i).Fuseyn %></td>
	<td align="center"><%= cMenu.FItemList(i).Fsortno %></td>
	<td align="center"><%= cMenu.FItemList(i).Fregusername %></td>
	<td align="center"><%= cMenu.FItemList(i).Fregdate %></td>
	<td align="center"><input type="button" value="수정" onClick="popreg('<%= cMenu.FItemList(i).FIdx %>');"></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" height="30">
    <td colspan="12" align="center">
    <% if cMenu.HasPreScroll then %>
		<a href="javascript:NextPage('<%= cMenu.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + cMenu.StartScrollPage to cMenu.FScrollCount + cMenu.StartScrollPage - 1 %>
		<% if i>cMenu.FTotalpage then Exit for %>
		<% if CStr(vPage)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if cMenu.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>

<iframe src="" name="ifram" width="0" height="0"></iframe>
<%
	Set cMenu = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->