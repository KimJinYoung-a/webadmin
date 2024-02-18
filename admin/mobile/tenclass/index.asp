<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/tenclass_Cls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : 모바일 ten_class
' History : 2018-02-27 이종화
'###############################################

	Dim isusing , dispcate
	dim page
	Dim i
	dim tenClassList
	Dim sDt , modiTime

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")

	if page="" then page=1
	If isusing = "" Then isusing = 1

	set tenClassList = new tenClass
	tenClassList.FPageSize	= 20
	tenClassList.FCurrPage	= page
	tenClassList.Fisusing	= isusing
	tenClassList.Fsdt		= sDt
	tenClassList.GetContentsList()

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
	location.href = "tenclass_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
}
$(function() {
  	// 라디오버튼
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");

});

function RefreshCaFavKeyWordRec(term){
	if(confirm("모바일- mdpick에 적용하시겠습니까?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_new_mdpick_xml.asp?term=" + term;
			refreshFrm.submit();
	}
}

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
			* 사용여부 :
			<select name="isusing">
				<option value="1" <%=chkiif(isusing=1,"selected","")%>>사용</option>
				<option value="0" <%=chkiif(isusing=0,"selected","")%>>사용안함</option>
			</select>
			지정일자 <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
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
		<!-- 신규등록 -->
    	<a href="tenclass_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		총 등록수 : <b><%=tenClassList.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=tenClassList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="22%">메인카피/서브카피</td>
    <td width="18%">시작일/종료일</td>
    <td width="10%">등록일</td>
    <td width="10%">등록자</td>
    <td width="10%">최종수정자</td>
    <td width="10%">사용여부</td>
</tr>
<%
	for i=0 to tenClassList.FResultCount-1
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(tenClassList.FItemList(i).Fisusing,"#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=tenClassList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=tenClassList.FItemList(i).Fidx%></td>
    <td onclick="jsmodify('<%=tenClassList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=tenClassList.FItemList(i).Fmaincopy%><br/><%=tenClassList.FItemList(i).Fsubcopy%></td>
	<td onclick="jsmodify('<%=tenClassList.FItemList(i).Fidx%>');" style="cursor:pointer;" align="left">
		<%
			If tenClassList.FItemList(i).Fstartdate <> "" And tenClassList.FItemList(i).Fenddate Then
				Response.Write "시작: "
				Response.Write replace(left(tenClassList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(tenClassList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(tenClassList.FItemList(i).Fstartdate),2,"0","R")
				Response.Write "<br />종료: "
				Response.Write replace(left(tenClassList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(tenClassList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(tenClassList.FItemList(i).Fenddate),2,"0","R")

				If cInt(datediff("d", now() , tenClassList.FItemList(i).Fenddate)) < 0 Or cInt(datediff("h", now() , tenClassList.FItemList(i).Fenddate )) < 0  Then
					Response.write " <span style=""color:red"">(종료)</span>"
				ElseIf cInt(datediff("d", tenClassList.FItemList(i).Fenddate , now())) < 1  Then '오늘날짜

					If cInt(datediff("d", tenClassList.FItemList(i).Fstartdate , now())) < 0 Then ' 시작전
						Response.write " <span style=""color:red"">(시작전)</span>"
					ElseIf cInt(datediff("h", now() , tenClassList.FItemList(i).Fenddate )) >= 0 And cInt(datediff("h", now() , tenClassList.FItemList(i).Fenddate )) < 24 Then ' 오늘
						Response.write " <span style=""color:red"">(약 "& cInt(datediff("h", now() , tenClassList.FItemList(i).Fenddate )) &" 시간후 종료)</span>"
					Elseif cInt(datediff("h", now() , tenClassList.FItemList(i).Fenddate )) > 24  Then 
						Response.write " <span style=""color:red"">("& cInt(datediff("d", now() , tenClassList.FItemList(i).Fenddate )) &"일 "& cInt(datediff("h", now() , tenClassList.FItemList(i).Fenddate )) - (cInt(datediff("d", now() , tenClassList.FItemList(i).Fenddate ))*24) &" 시간후 종료)</span>"
					End If 

				End If
			End If
		%>
	</td>
	<td><%=left(tenClassList.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(tenClassList.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = tenClassList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(tenClassList.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=chkiif(tenClassList.FItemList(i).Fisusing,"사용함","사용안함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if tenClassList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= tenClassList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + tenClassList.StartScrollPage to tenClassList.StartScrollPage + tenClassList.FScrollCount - 1 %>
				<% if (i > tenClassList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(tenClassList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if tenClassList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set tenClassList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->