<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : index.asp
' Discription : 모바일 사이트 알림배너
' History : 2013.04.01 이종화 생성
'			2016.07.21 한용민 수정
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/main_noticebanner.asp" -->
<%	
Dim userlevel , isusing, sdate, edate, page, i, vParam
	page = request("page")
	sdate = RequestCheckVar(request("startday"),13)
	edate = RequestCheckVar(request("endday"),13)
	userlevel = RequestCheckVar(request("userlevel"),13)
	isusing = RequestCheckVar(request("isusing"),13)

If sdate = "" Then sdate = date
If edate = "" Then edate = DateAdd("d" , 7 , date) '기본 일주일후
if page="" then page=1

vParam = "&sdate="&sdate&"&edate="&edate&"&userlevel="&userlevel&"&isusing="&isusing

dim oNoticebanner
set oNoticebanner = new CMainbanner
	oNoticebanner.FPageSize		= 20
	oNoticebanner.FCurrPage		= page
	oNoticebanner.FSearchSdate = sdate
	oNoticebanner.FSearchEdate = edate
	oNoticebanner.Fisusing			= isusing
	oNoticebanner.Fuserlevel		= userlevel
	oNoticebanner.GetContentsList()

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

// 오늘 내일 지난 7일
function chkdate(v){
	var frm = document.frm;
	var nowdate = new Date();
	var year  = nowdate.getFullYear();
	var month = nowdate.getMonth() + 1; // 1월=0,12월=11이므로 1 더함
	var day   = nowdate.getDate();

	if (("" + month).length == 1) { month = "0" + month; }
	if (("" + day).length   == 1) { day   = "0" + day;   }

	today  = year + "-" + month + "-" + day; //오늘임

	if (v == "N"){
		frm.startday.value = today;
		frm.endday.value =  today;
	}else if (v =="T"){
		frm.startday.value = today;
		frm.endday.value = mathdate(today,"1");
	}else if (v =="W"){
		frm.startday.value = mathdate(today,"-7");
		frm.endday.value = today;
	}
}

// 날짜 계산
function mathdate(date,v){
		var input1 = date;
		var input2 = v;
 		var dateinfo = input1.split("-");
		var src = new Date(dateinfo[0], dateinfo[1]-1, dateinfo[2]);

		src.setDate(src.getDate() + parseInt(input2));
		var year = src.getFullYear();
	    var month = src.getMonth() + 1;
		var date = src.getDate();

		if(month<10) month = "0" + month;
 
		if(date<10) date = "0" + date;
 
		var result = year + "-" + month + "-" + date;

		return result;
}

//수정
function jsmodify(v){
	location.href = "nb_insert.asp?menupos=<%=menupos%>&idx="+v;
}

function jschgusing(v,idx){
	location.href = "nb_proc.asp?iidx="+idx+"&isusing="+v+"&mode=chg";
}

$(function(){
	//달력대화창 설정
	var arrDayMin = ["일","월","화","수","목","금","토"];
	var arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
    	showCurrentAtPos: 1,
      	showOn: "button",
      	maxDate: "<%=edate%>",
    	onClose: function( selectedDate ) {
    		$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
    	}
    });
    $("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
      	showOn: "button",
      	minDate: "<%=sdate%>",
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
});

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			* 사용여부 :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			* 적용구분 :&nbsp;&nbsp;<% DrawselectboxUserLevel "userlevel", userlevel, "" %>
			</div>
			<div>
	       	* 조회기간 :&nbsp;&nbsp; 
			<input type="button"  class="button_s" value="오늘" onclick="chkdate('N');">&nbsp;
			<input type="button"  class="button_s" value="내일" onclick="chkdate('T');">&nbsp;
			<input type="button"  class="button_s" value="지난 7일" onclick="chkdate('W');">&nbsp;

			<input type="text" id="sDt" name="startday" size="10" value="<%=sdate%>" /> ~
			<input type="text" id="eDt" name="endday" size="10" value="<%=edate%>" />
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
    	<a href="nb_insert.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		총 등록수 : <b><%=oNoticebanner.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oNoticebanner.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="7%">번호(idx)</td>
    <td width="7%">등록자</td>
    <td width="16%">제목</td>
    <td width="17%">노출기간</td>
    <td width="7%">우선순위</td>
    <td width="10%">사용여부</td>
    <td >등급구분</td>
    <td width="10%">등록일</td>
</tr>
<% 
	Dim tempSdate , tempStime ,  tempEdate , tempEtime

	for i=0 to oNoticebanner.FResultCount-1 

		''날짜 시간 분리
		tempSdate = ""
		tempEdate = ""
		tempStime = ""
		tempEtime = ""
		If Len(oNoticebanner.FItemList(i).Fstartday) <= 10 Or Len(oNoticebanner.FItemList(i).Fendday) <= 10  Then
			tempSdate = oNoticebanner.FItemList(i).Fstartday
			tempEdate = oNoticebanner.FItemList(i).Fendday
		Else
			tempSdate = Left(oNoticebanner.FItemList(i).Fstartday,10)
			tempStime = Trim(right(oNoticebanner.FItemList(i).Fstartday,11))
			tempEdate = Left(oNoticebanner.FItemList(i).Fendday,10)
			tempEtime = Trim(right(oNoticebanner.FItemList(i).Fendday,11))
		End If 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oNoticebanner.FItemList(i).Fisusing="1","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=oNoticebanner.FItemList(i).Fiidx%>');" style="cursor:pointer;"><%=oNoticebanner.FItemList(i).Fiidx%></td>
    <td><%=oNoticebanner.FItemList(i).Fwriter%></td>
    <td><%=oNoticebanner.FItemList(i).Ftitle%></td>
    <td><%=tempSdate%> ~ <%=tempEdate%><% If tempStime <> "" Or tempEtime <>"" Then %><br/>(<%=tempStime%> ~ <%=tempEtime%>)<% End If %></td>
    <td><%=oNoticebanner.FItemList(i).Fsorting%></td>
    <td><%=chkiif(oNoticebanner.FItemList(i).Fisusing="0","사용안함","사용함")%>&nbsp;<input type="button" value="변경"  class="button_s" onclick="jschgusing('<%=oNoticebanner.FItemList(i).Fisusing%>','<%=oNoticebanner.FItemList(i).Fiidx%>');"/></td>
    <td><%=oNoticebanner.FItemList(i).FutnArr%></td>
    <td><%=Left(oNoticebanner.FItemList(i).Fwritedate,10)%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if oNoticebanner.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oNoticebanner.StartScrollPage-1 %><%=vParam%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oNoticebanner.StartScrollPage to oNoticebanner.StartScrollPage + oNoticebanner.FScrollCount - 1 %>
				<% if (i > oNoticebanner.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oNoticebanner.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %><%=vParam%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oNoticebanner.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %><%=vParam%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set oNoticebanner = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->