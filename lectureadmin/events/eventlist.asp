<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/event/eventCls.asp"-->
<%
Dim research, evt_startdate, evt_enddate
Dim searchKey, searchString, gubun, isusing, evting
Dim page, oEvent, i
page			= requestCheckvar(request("page"),10)
research		= requestCheckvar(request("research"),10)
evt_startdate	= requestCheckvar(request("evt_startdate"),10)
evt_enddate		= requestCheckvar(request("evt_enddate"),10)
isusing			= requestCheckvar(request("isusing"),2)
searchKey		= requestCheckvar(request("searchKey"),10)
searchString	= requestCheckvar(request("searchString"),128)
evting			= requestCheckvar(request("evting"),10)

If page = "" Then page = 1
If (research = "") Then
	searchKey	= "evt_name"
End If

Set oEvent = new CEvent
	gubun						= oEvent.getWhatMyJob()
	If gubun = "X" Then
		response.write "<script>alert('강좌 또는 DIY가 선택되지 않았습니다.\n핑거스 관리자에게 문의주세요.');history.back(-1);</script>"
		response.end
	End If

	oEvent.FCurrPage			= page
	oEvent.FPageSize			= 12
	oEvent.FRectStartdate		= evt_startdate
	oEvent.FRectEnddate			= evt_enddate
	oEvent.FRectGubun			= gubun
	oEvent.FRectIsusing			= isusing
	oEvent.FRectEvting			= evting
	oEvent.FRectSearchKey		= searchKey
	oEvent.FRectSearchString	= searchString
	oEvent.getEventItemList
%>
<script>
function goRegEvent(v, g){
	location.href='/lectureadmin/events/event_regist.asp?idx='+v+'&gubun='+g+'&menupos=<%=menupos%>'
}
function goPage(pg){
    frmEvt.page.value = pg;
    frmEvt.submit();
}
function jsSetDate(n, m){
	document.getElementById("evt_startdate").value = "";
	document.getElementById("evt_enddate").value = "";
	var date = new Date();
	if(n == 7 || n == 15){
		var start = new Date(Date.parse(date) - n * 1000 * 60 * 60 * 24);
		var today = new Date(Date.parse(date) - m * 1000 * 60 * 60 * 24);
	
		var yyyy = start.getFullYear();
		var mm = start.getMonth()+1;
		var dd = start.getDate();

		var t_yyyy = today.getFullYear();
		var t_mm = today.getMonth()+1;
		var t_dd = today.getDate();
	}else{
        var t_mm = date.getMonth() + 1;
        var t_dd = date.getDate();
        var t_yyyy = date.getFullYear();
 		if(n == 30){
        	var preDate = new Date(date.setMonth(t_mm - 1)); 
        }else{
        	var preDate = new Date(date.setMonth(t_mm - 3)); 
        }
        var mm = preDate.getMonth() ; 
        var dd = preDate.getDate();
        var yyyy = preDate.getFullYear();
	}
	if (t_mm <10){
		t_mm = "0"+t_mm;
	}
	if (mm <10){
		mm = "0"+mm;
	}
	if (dd <10){
		dd = "0"+dd;
	}
	if (t_dd <10){
		t_dd = "0"+t_dd;
	}
	document.getElementById("evt_startdate").value = yyyy + "-" + mm + "-" + dd; 
	document.getElementById("evt_enddate").value = t_yyyy + "-" + t_mm + "-" + t_dd;
}
</script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmEvt" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		기간
		<input id="evt_startdate" readonly name="evt_startdate" value="<%=evt_startdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="evt_startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> 00:00:00 ~
		<input id="evt_enddate" readonly name="evt_enddate" value="<%=evt_enddate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="evt_enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> 00:00:00
		<script language="javascript">
		var CAL_Start = new Calendar({
			inputField : "evt_startdate", trigger    : "evt_startdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var CAL_End = new Calendar({
			inputField : "evt_enddate", trigger    : "evt_enddate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		</script>
		<input type="button" value="최근7일" class="button" onClick="jsSetDate(7,0)">
		<input type="button" value="최근15일" class="button" onClick="jsSetDate(15,0)">
		<input type="button" value="최근1개월" class="button" onClick="jsSetDate(30,0)">
		<input type="button" value="최근3개월" class="button" onClick="jsSetDate(90,0)">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		검색
		<select name="searchKey" class="select">
			<option value="eCode" <%= chkiif(searchKey="eCode", "selected", "") %>>번호</option>
			<option value="contentsCode" <%= chkiif(searchKey="contentsCode", "selected", "") %>>작품/강좌코드</option>
			<option value="evt_name" <%= chkiif(searchKey="evt_name", "selected", "") %>>이벤트명</option>
		</select>
		<input type="text" class="text" name="searchString" size="50" value="<%= searchString %>">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frmEvt.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		진행/마감
		<select name="evting" class="select">
			<option value="">전체
			<option value="ing" <%= chkiif(evting="ing", "selected", "") %>>진행
			<option value="end" <%= chkiif(evting="end", "selected", "") %>>마감
			<option value="will" <%= chkiif(evting="will", "selected", "") %>>예정
		</select>
		&nbsp;&nbsp;&nbsp;
		사용여부
		<select name="isusing" class="select">
			<option value="">전체
			<option value="Y" <%= chkiif(isusing = "Y", "selected", "") %> >Y
			<option value="N" <%= chkiif(isusing = "N", "selected", "") %> >N
		</select>
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td align="left" colspan="6">
		건수 : <b><%= FormatNumber(oEvent.FTotalCount,0) %>건</b>&nbsp;&nbsp;
		Page : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oEvent.FTotalPage,0) %> </b>
	</td>
	<td align="center">
		<input type="button" class="button" value="신규등록" onclick="goRegEvent('', '<%=gubun%>');">
	</td>
</tr>
<tr height="35" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">번호</td>
	<td width="350">기간</td>
	<td width="120">등록위치</td>
	<td width="100">작품/강좌코드</td>
	<td>이벤트명</td>
	<td width="80">사용여부</td>
	<td width="150">등록일</td>
</tr>

<% For i=0 to oEvent.FResultCount - 1 %>
<tr height="35" align="center" bgcolor="#FFFFFF" style="cursor:pointer;" onclick="goRegEvent('<%= oEvent.FItemList(i).FIdx %>', '<%=gubun%>')" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff';>
	<td><%= oEvent.FItemList(i).FIdx %></td>
	<td>
	<%
		response.write FormatDate(oEvent.FItemList(i).FEvt_startdate, "0000.00.00") & " 00:00:00 ~ " & FormatDate(oEvent.FItemList(i).FEvt_enddate, "0000.00.00") & " 00:00:00"
		If oEvent.FItemList(i).FEvt_startdate <= now() AND oEvent.FItemList(i).FEvt_enddate >= now() Then
			response.write "<font color='BLUE'> (진행)</font>"
		ElseIf oEvent.FItemList(i).FEvt_enddate <= now() Then
			response.write "<font color='RED'> (마감)</font>"
		Else
			response.write "<font color='GRAY'> (예정)</font>"
		End If
	%>
	</td>
	<td><%= Chkiif(oEvent.FItemList(i).FGubun="D", "작가 프로필(작품)", "강사 프로필(강좌)") %></td>
	<td><%= oEvent.FItemList(i).FContentsCode %></td>
	<td><%= oEvent.FItemList(i).FEvt_name %></td>
	<td><%= Chkiif(oEvent.FItemList(i).FIsusing="Y", "<font color='green'>사용함</font>", "사용안함") %></td>
	<td><%= FormatDate(oEvent.FItemList(i).FRegdate, "0000.00.00") %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oEvent.HasPreScroll then %>
		<a href="javascript:goPage('<%= oEvent.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oEvent.StartScrollPage to oEvent.FScrollCount + oEvent.StartScrollPage - 1 %>
    		<% if i>oEvent.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oEvent.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->