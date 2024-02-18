<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/sbsvshopCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : 모바일 메인 메인 기획전 링크
' History : 2016.04.07 이종화
'###############################################
	
	dim dramaList , i , page , isusing
	Dim sDt , modiTime , idx

	page	= RequestCheckVar(request("page"),10)
	isusing = RequestCheckVar(request("isusing"),1)
	sDt		= RequestCheckVar(request("prevDate"),10)
	idx		= RequestCheckVar(request("idx"),10)

	if page="" then page=1
	If isusing = "" Then isusing = 1

	set dramaList = new sbsvshop
	dramaList.FPageSize		= 20
	dramaList.FCurrPage		= page
	dramaList.FRectisusing	= isusing
	dramaList.FRectidx		= idx
	dramaList.Fsdt			= sDt
	dramaList.fnDramaContentsListGet()

%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script type='text/javascript'>
<!--
//수정
function jsmodify(v){
	location.href = "drama_insert.asp?menupos=<%=menupos%>&listidx="+v+"&prevDate=<%=sDt%>";
}

function popDramaManage(){
    var popwin = window.open('pop_dramalist.asp','mainposcodeedit','width=1200,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

$(function() {
  	// 라디오버튼
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
	
});

function jssearch(){
	document.frm.submit();
}
-->
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			* 사용여부 :&nbsp;&nbsp;
			<select name="isusing">
				<option value="1" <%=chkiif(isusing = 1," selected","")%>>사용함</option>
				<option value="0" <%=chkiif(isusing = 0," selected","")%>>사용안함</option>
			</select>
			&nbsp;&nbsp;&nbsp;
			지정일자 <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			* 드라마명 : <% Call getdramaname("idx",idx,"on") %>
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onclick="jssearch();">
		</td>
	</tr>
</form>	
</table>
<!-- 검색 끝 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td align="right">
		<input type="button" class="button" value="드라마관리" onClick="popDramaManage();">&nbsp;
    	<a href="drama_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		총 등록수 : <b><%=dramaList.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=dramaList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
    <td width="10%">드라마명</td>
	<td width="22%">제목/내용</td>
    <td width="18%">시작일/종료일</td>
    <td width="10%">등록일</td>
    <td width="10%">등록자</td>
    <td width="10%">최종수정자</td>
    <td width="10%">사용여부</td>
</tr>
<% 
	for i=0 to dramaList.FResultCount-1 
%>
<tr height="30" align="center" bgcolor="<%=chkiif(isusing=1,"#FFFFFF","#F9BF3C")%>">
    <td onclick="jsmodify('<%=dramaList.FItemList(i).Flistidx%>');" style="cursor:pointer;"><%=dramaList.FItemList(i).Flistidx%></td>
    <td onclick="jsmodify('<%=dramaList.FItemList(i).Flistidx%>');" style="cursor:pointer;"><%=dramaList.FItemList(i).Fdramatitle%></td>
    <td onclick="jsmodify('<%=dramaList.FItemList(i).Flistidx%>');" style="cursor:pointer;"><%=dramaList.FItemList(i).Ftitle%><br/><%=dramaList.FItemList(i).Fcontents%></td>
	<td onclick="jsmodify('<%=dramaList.FItemList(i).Flistidx%>');" style="cursor:pointer;" align="left">
		<% 
			If dramaList.FItemList(i).Fstartdate <> "" And dramaList.FItemList(i).Fenddate Then 
				Response.Write "시작: "
				Response.Write replace(left(dramaList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(dramaList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(dramaList.FItemList(i).Fstartdate),2,"0","R")
				Response.Write "<br />종료: "
				Response.Write replace(left(dramaList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(dramaList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(dramaList.FItemList(i).Fenddate),2,"0","R")
				If clng(datediff("d", now() , dramaList.FItemList(i).Fenddate)) < 0 Or clng(datediff("h", now() , dramaList.FItemList(i).Fenddate )) < 0  Then 
					Response.write " <span style=""color:red"">(종료)</span>"
				ElseIf clng(datediff("d", dramaList.FItemList(i).Fenddate , now())) = 0  Then '종료일이 오늘날짜
					If clng(datediff("h", now() , dramaList.FItemList(i).Fenddate )) >= 0 And clng(datediff("h", now() , dramaList.FItemList(i).Fenddate )) < 24 Then ' 오늘
						Response.write " <span style=""color:red"">(약 "& clng(datediff("h", now() , dramaList.FItemList(i).Fenddate )) &" 시간후 종료)</span>"
					Else  ' 시작전
						Response.write " <span style=""color:red"">(시작전)</span>"					
					End If 
				'// 시작일이 오늘날짜이고 종료일이 오늘이 아니면
				ElseIf clng(datediff("d", dramaList.FItemList(i).Fstartdate , now()))>=0 And clng(datediff("d", dramaList.FItemList(i).Fenddate , now())) < 0 Then
					Response.write " <span style=""color:red"">(약 "&clng(datediff("d", now() , dramaList.FItemList(i).Fenddate ))&"일 " &clng(datediff("h", now() , dramaList.FItemList(i).Fenddate ))-clng(datediff("d", now() , dramaList.FItemList(i).Fenddate ))*24 &"시간후 종료)</span>"
				ElseIf clng(datediff("d", dramaList.FItemList(i).Fenddate , now())) < 0 Then
					If clng(datediff("h", now() , dramaList.FItemList(i).Fenddate )) >= 0 And clng(datediff("h", now() , dramaList.FItemList(i).Fenddate )) < 24 Then ' 오늘
						Response.write " <span style=""color:red"">(약 "& clng(datediff("h", now() , dramaList.FItemList(i).Fenddate )) &" 시간후 종료)</span>"
					Else  ' 시작전
						Response.write " <span style=""color:red"">(시작전)</span>"
					End If 
				End If
			End If 
		%>
	</td>
	<td><%=left(dramaList.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(dramaList.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = dramaList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(dramaList.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=chkiif(dramaList.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if dramaList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= dramaList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + dramaList.StartScrollPage to dramaList.StartScrollPage + dramaList.FScrollCount - 1 %>
				<% if (i > dramaList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(dramaList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if dramaList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set dramaList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->