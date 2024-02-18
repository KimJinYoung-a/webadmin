<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.30 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/managerCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim olist, page, i, makerid, brandgubun, brandgubunexists, research
dim catecode, standardCateCode, mduserid
	catecode	= request("catecode")
	standardCateCode	= request("standardCateCode")
	mduserid	= request("mduserid")
	page	= request("page")
	makerid	= request("makerid")
	brandgubun	= request("brandgubun")
	menupos	= request("menupos")
	research	= request("research")
	brandgubunexists	= request("brandgubunexists")
	
If page = ""	Then page = 1
	
SET olist = new cmanager
	olist.FCurrPage		= page
	olist.FPageSize		= 50
	olist.FrectMakerid		= makerid
	olist.Frectbrandgubun		= brandgubun
	olist.Frectbrandgubunexists = brandgubunexists
	olist.Frectcatecode = catecode
	olist.FrectstandardCateCode = standardCateCode
	olist.Frectmduserid = mduserid	
	olist.sbmanagerlist
%>

<script language="javascript">

var ichk = 1;
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.type=="checkbox")) {
			e.checked = blnChk ;
		}
	}
}

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

</script>

<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>브랜드구분지정</b>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : 
		<%' drawmanager_ID_with_Name "makerid",makerid %>
		<% drawSelectBoxDesignerwithName "makerid",makerid %>
		&nbsp;&nbsp;
		* 브랜드구분 : <% drawSelectBoxbrandgubun "brandgubun",brandgubun , " onchange=""gosubmit('');""" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 대표카테고리 : 
		기능<% SelectBoxBrandCategory "catecode", catecode %>
		전시<%= fnStandardDispCateSelectBox(1,"", "standardCateCode", standardCateCode, "")%>
		&nbsp;&nbsp;
		* 담당MD : <% drawSelectBoxCoWorker_OnOff "mduserid", mduserid, "on" %>
		&nbsp;&nbsp;
		* 브랜드구분지정여부 : <% drawSelectBoxUsingYN "brandgubunexists", brandgubunexists %>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<!--<input type="button" value="신규등록" onclick="javascript:location.href='manager_write.asp?menupos=<%=menupos%>';" class="button">-->
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=olist.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= olist.FTotalPage %></b>		
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<!--<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>-->
	<td>브랜드ID</td>
	<td>브랜드구분</td>
	<td>이미지(PREMIUM 전용)</td>
	<td>최종수정</td>
	<td>비고</td>
</tr>
<% if olist.fresultcount >0 then %>
<% For i = 0 to olist.fresultcount -1 %>

<% if olist.FItemlist(i).fidx <> "" then %>
	<tr height="25" bgcolor="#FFFFFF" align="center">
<% else %>
	<tr height="25" bgcolor="#f1f1f1" align="center">
<% end if %>

	<!--<td align="center"><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%'= olist.FItemlist(i).FIdx %>"></td>-->
	<td><%= olist.FItemlist(i).FMakerid %></td>
	<td>
		<%= getbrandgubunname(olist.FItemlist(i).fbrandgubun) %>
	
		<% if olist.FItemlist(i).fidx = "" or isnull(olist.FItemlist(i).fidx) then %>
			(미지정)
		<% end if %>
	</td>
	<td>
		<% if olist.FItemlist(i).fbrandgubun = "4" then %>
			<% if olist.FItemlist(i).fsubtopimage<>"" then %>
				<img src="<%=uploadUrl%>/brandstreet/manager/<%= olist.FItemlist(i).fsubtopimage %>" width=100 height=50>
			<% else %>
			<%= olist.FItemlist(i).fdesignis %>
			<% end if %>
		<% end if %>
	</td>
	<td>
		<% if olist.FItemlist(i).flastupdate <> "" then %>
			<%= olist.FItemlist(i).flastupdate %>
			<br>(<%= olist.FItemlist(i).flastadminid %>)
		<% end if %>
	</td>
	<td>
		<% if olist.FItemlist(i).fidx <> "" then %>
			<input type="button" onclick="javascript:location.href='manager_write.asp?idx=<%=olist.FItemlist(i).FIdx%>&makerid=<%= olist.FItemlist(i).FMakerid %>&menupos=<%=menupos%>';" value="수정" class="button">
		<% else %>
			<input type="button" onclick="javascript:location.href='manager_write.asp?idx=<%=olist.FItemlist(i).FIdx%>&makerid=<%= olist.FItemlist(i).FMakerid %>&menupos=<%=menupos%>';" value="등록" class="button">
		<% end if %>
	</td>
</tr>
<% Next %>

<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If olist.HasPreScroll Then %>
			<span class="olist_link"><a href="javascript:gosubmit('<%= olist.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + olist.StartScrollPage to olist.StartScrollPage + olist.FScrollCount - 1 %>
			<% If (i > olist.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(olist.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="olist_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If olist.HasNextScroll Then %>
			<span class="olist_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</form>
</table>

<%
SET olist = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->