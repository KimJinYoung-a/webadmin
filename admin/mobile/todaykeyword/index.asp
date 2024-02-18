<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/today_keywordCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : 모바일 메인 enjoybanner
' History : 2014.06.23 이종화 생성
'			2018.02.07 한용민 수정(에러 수정)
'###############################################

	Dim isusing , dispcate , validdate , research
	dim page
	Dim i
	dim oBrandinfo
	Dim sDt , modiTime , sedatechk
	Dim addtype

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")
	validdate= request("validdate")
	research= request("research")
	sedatechk = request("sedatechk")
	addtype = request("addtype")

	if ((research="") and (isusing="")) then
	    isusing = "Y"
	    validdate = "on"
	end if

	if page="" then page=1

	set oBrandinfo = new CMainbanner
	oBrandinfo.FPageSize			= 20
	oBrandinfo.FCurrPage			= page
	oBrandinfo.Fisusing				= isusing
	oBrandinfo.Fsdt					= sDt
	oBrandinfo.FRectvaliddate		= validdate
	oBrandinfo.FRectsedatechk		= sedatechk '//시작일 기준 체크
	oBrandinfo.GetContentsList()

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
	location.href = "todaykeyword_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
}
$(function() {
  	$("input[type=submit]").button();

  	// 라디오버튼
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");

});

function RefreshCaFavKeyWordRec(term){
	if(confirm("모바일- enjoyevent에 적용하시겠습니까?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_todayenjoy_xml.asp";
			refreshFrm.submit();
	}
}

function NextPage(page){
    frm.page.value = page;
    frm.submit();
}
function addContents(){
	var dateOptionParam
	var frm = document.frm
	dateOptionParam = frm.prevDate.value

	document.location.href="todaykeyword_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&dateoption="+dateOptionParam
}
-->
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전&nbsp;
			* 사용여부 :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			시작일기준 <input type="checkbox" name="sedatechk" <% if sedatechk="on" then response.write "checked" %> />
			지정일자 <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			&nbsp;
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:submit();">
		</td>
	</tr>
</form>
</table>
<!-- 검색 끝 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
<!-- 	<td><a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>XML Real 적용(예약)</a></td> -->
    <td align="right">
		<!-- 신규등록 -->
    	<a href="javascript:addContents();"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		총 등록수 : <b><%=oBrandinfo.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oBrandinfo.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="20%">키워드이미지</td>
	<td width="10%">키워드</td>
    <td width="15%">시작일/종료일</td>
    <td width="10%">등록일</td>
    <td width="10%">등록자</td>
    <td width="10%">최종수정자</td>
    <td width="10%">사용여부</td>
</tr>
<%
	Dim ii , itemname1 ,  itemimg1 , itemname2 ,  itemimg2 , itemname3 ,  itemimg3 , itemname4 ,  itemimg4
	Dim itemid1 ,  itemid2 , itemid3 , itemid4
	for i=0 to oBrandinfo.FResultCount-1
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oBrandinfo.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=oBrandinfo.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oBrandinfo.FItemList(i).Fidx%></td>
    <td align="left">
		<%
			itemimg1 =  oBrandinfo.FItemList(i).Fitemimg1
			itemimg2 =  oBrandinfo.FItemList(i).Fitemimg2
			itemimg3 =  oBrandinfo.FItemList(i).Fitemimg3
			itemimg4 =  oBrandinfo.FItemList(i).Fitemimg4
			If oBrandinfo.FItemList(i).Fiteminfo <> "" and not isnull(oBrandinfo.FItemList(i).Fiteminfo) Then
				If ubound(Split(oBrandinfo.FItemList(i).Fiteminfo,"^^")) > 0 Then ' 이미지 3개 정보
					For ii = 0 To ubound(Split(oBrandinfo.FItemList(i).Fiteminfo,","))
						If CStr(oBrandinfo.FItemList(i).Fitemid1) = CStr(Split(Split(oBrandinfo.FItemList(i).Fiteminfo,",")(ii),"|")(0)) And oBrandinfo.FItemList(i).Fitemimg1 = "" Then
							itemimg1 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(oBrandinfo.FItemList(i).Fitemid1) & "/" & Split(Split(oBrandinfo.FItemList(i).Fiteminfo,",")(ii),"|")(2)
						End If

						If CStr(oBrandinfo.FItemList(i).Fitemid2) = CStr(Split(Split(oBrandinfo.FItemList(i).Fiteminfo,",")(ii),"|")(0)) And oBrandinfo.FItemList(i).Fitemimg2 = "" Then
							itemimg2 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(oBrandinfo.FItemList(i).Fitemid2) & "/" & Split(Split(oBrandinfo.FItemList(i).Fiteminfo,",")(ii),"|")(2)
						End If

						If CStr(oBrandinfo.FItemList(i).Fitemid3) = CStr(Split(Split(oBrandinfo.FItemList(i).Fiteminfo,",")(ii),"|")(0)) And oBrandinfo.FItemList(i).Fitemimg3 = "" Then
							itemimg3 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(oBrandinfo.FItemList(i).Fitemid3) & "/" & Split(Split(oBrandinfo.FItemList(i).Fiteminfo,",")(ii),"|")(2)
						End If

						If CStr(oBrandinfo.FItemList(i).Fitemid4) = CStr(Split(Split(oBrandinfo.FItemList(i).Fiteminfo,",")(ii),"|")(0)) And oBrandinfo.FItemList(i).Fitemimg4 = "" Then
							itemimg4 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(oBrandinfo.FItemList(i).Fitemid4) & "/" & Split(Split(oBrandinfo.FItemList(i).Fiteminfo,",")(ii),"|")(2)
						End If
					Next
				End If
			End If
		%>
		<img src="<%=itemimg1%>" width="70" alt=""/>
		<img src="<%=itemimg2%>" width="70" alt=""/>
		<img src="<%=itemimg3%>" width="70" alt=""/>
		<img src="<%=itemimg4%>" width="70" alt=""/>
		<br/><font color="<%=chkiif(oBrandinfo.FItemList(i).Fpicknum=1,"red","")%>">#1:<%=oBrandinfo.FItemList(i).Fitemid1%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="<%=chkiif(oBrandinfo.FItemList(i).Fpicknum=2,"red","")%>">#2:<%=oBrandinfo.FItemList(i).Fitemid2%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="<%=chkiif(oBrandinfo.FItemList(i).Fpicknum=3,"red","")%>">#3:<%=oBrandinfo.FItemList(i).Fitemid3%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="<%=chkiif(oBrandinfo.FItemList(i).Fpicknum=4,"red","")%>">#4:<%=oBrandinfo.FItemList(i).Fitemid4%></font>
	</td>
	<td><%=oBrandinfo.FItemList(i).Fkeyword%></td>
	<td>
		<%
			Response.Write "시작: "
			Response.Write replace(left(oBrandinfo.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oBrandinfo.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oBrandinfo.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />종료: "
			Response.Write replace(left(oBrandinfo.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(oBrandinfo.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(oBrandinfo.FItemList(i).Fenddate),2,"0","R")
		%>
	</td>
	<td><%=left(oBrandinfo.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(oBrandinfo.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = oBrandinfo.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(oBrandinfo.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=chkiif(oBrandinfo.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td colspan="11" align="center">
		<% if oBrandinfo.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oBrandinfo.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i = 0 + oBrandinfo.StartScrollPage to oBrandinfo.StartScrollPage + oBrandinfo.FScrollCount - 1 %>
			<% if (i > oBrandinfo.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oBrandinfo.FCurrPage) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oBrandinfo.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<%
set oBrandinfo = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->