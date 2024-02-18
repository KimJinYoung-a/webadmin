<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/mktevtbannerCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : 모바일 이벤트 마케팅 배너
' History : 2015-01-07 이종화
'###############################################
	
	Dim isusing , dispcate , validdate , research
	dim page 
	Dim i
	dim oTodaydealList
	Dim sDt , modiTime , gubun

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")
	validdate= request("validdate")
	research= request("research")
	gubun= request("gubun")

	if ((research="") and (isusing="")) then
	    isusing = "Y"
	    validdate = "on"
	end if
	
	if page="" then page=1

	set oTodaydealList = new CEvtMktbanner
	oTodaydealList.FPageSize		= 20
	oTodaydealList.FCurrPage		= page
	oTodaydealList.Fisusing			= isusing
	oTodaydealList.Fsdt				= sDt
	oTodaydealList.FRectvaliddate	= validdate
	oTodaydealList.FRectgubun		= gubun
	oTodaydealList.GetContentsList()

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
	location.href = "deal_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
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

function NextPage(page){
    frm.page.value = page;
    frm.submit();
}


function Addnewcontents(val){
    var popwin = window.open('mktban_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&idx='+val,'mainposcodeedit','width=800,height=550,scrollbars=yes,resizable=yes');
    popwin.focus();
}
-->
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="<%=page%>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전&nbsp;
			* 사용여부 :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			지정일자 <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			&nbsp;&nbsp;&nbsp;
			구분 : <select name="gubun" onchange="onchgbox(this.value);" width="100">
						<option value="">구분선택</option>
						<option value="1" <%=chkiif(gubun="1","selected","")%>>Mobile & Apps</option>
						<option value="2" <%=chkiif(gubun="2","selected","")%>>Mobile</option>
						<option value="3" <%=chkiif(gubun="3","selected","")%>>Apps</option>
					</select>&nbsp;&nbsp;
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
    <td align="right">
		<!-- 신규등록 -->
    	<a href="" onclick="Addnewcontents(''); return false;" target="_blank"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		총 등록수 : <b><%=oTodaydealList.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oTodaydealList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="10%">구분</td>
	<td width="10%">이벤트 구분</td>
	<td width="7%">배너이미지</td>	 
    <td width="20%">시작일/종료일</td>
    <td width="10%">등록자<br/>최종수정자</td>
    <td width="10%">고정유무</td>
    <td width="5%">우선순위</td>
    <td width="5%">사용여부</td>
</tr>
<% 
	for i=0 to oTodaydealList.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oTodaydealList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="Addnewcontents('<%=oTodaydealList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oTodaydealList.FItemList(i).Fidx%></td>
	<td><%=getGubun(oTodaydealList.FItemList(i).Fgubun)%></td>
	<td><%=chkiif(oTodaydealList.FItemList(i).Fevtgubun = "1","기획전","마케팅")%></td>
	<td>
		<img src="<%=oTodaydealList.FItemList(i).Fmktimg%>" width="300" alt="<%=oTodaydealList.FItemList(i).Faltname%>"/>
	</td>
	<td align="left">
		<% 
			Response.Write "등록일 : "
			Response.Write left(oTodaydealList.FItemList(i).Fregdate,10) &"</br>"
			Response.Write "시작: "
			Response.Write replace(left(oTodaydealList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oTodaydealList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oTodaydealList.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />종료: "
			Response.Write replace(left(oTodaydealList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(oTodaydealList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(oTodaydealList.FItemList(i).Fenddate),2,"0","R")
			
			If cInt(datediff("d", now() , oTodaydealList.FItemList(i).Fenddate)) < 0 Or cInt(datediff("h", now() , oTodaydealList.FItemList(i).Fenddate )) < 0  Then 
					Response.write " <span style=""color:red"">(종료)</span>"
			ElseIf cInt(datediff("d", oTodaydealList.FItemList(i).Fenddate , now())) < 1  Then '오늘날짜

				If cInt(datediff("h", now() , oTodaydealList.FItemList(i).Fenddate )) >= 0 Then ' 오늘
				Response.write " <span style=""color:red"">(약 "& cInt(datediff("h", now() , oTodaydealList.FItemList(i).Fenddate )) &" 시간후 종료)</span>"
				Else  ' 시작전
				Response.write " <span style=""color:red"">(시작전)</span>"					
				End If 

			End If 
		%>
	</td>
	<td>
		<%=getStaffUserName(oTodaydealList.FItemList(i).Fadminid)%><br/>
		<%
			modiTime = oTodaydealList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write "(최종 : " & getStaffUserName(oTodaydealList.FItemList(i).Flastadminid) & " " & left(modiTime,10) & ")"
			end if
		%>
	</td>
	<td><%=chkiif(oTodaydealList.FItemList(i).Ftopfixed = "Y","고정","비고정")%></td>
    <td><%=oTodaydealList.FItemList(i).Fsortnum%></td>
    <td><%=chkiif(oTodaydealList.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td colspan="11" align="center">
		<% if oTodaydealList.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oTodaydealList.StarScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i = 0 + oTodaydealList.StartScrollPage to oTodaydealList.StartScrollPage + oTodaydealList.FScrollCount - 1 %>
			<% if (i > oTodaydealList.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oTodaydealList.FCurrPage) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oTodaydealList.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<%
set oTodaydealList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->