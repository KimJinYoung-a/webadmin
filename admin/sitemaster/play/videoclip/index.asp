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
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<%
	Dim oVideoClip, i , page , state ,idx , title , reservationdate , viewtitle , tagyn , partwdid , viewno
	Dim playcate : playcate = 6 'videoclip
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	title = request("title")
	reservationdate = request("reservationdate")
	viewtitle = request("viewtitle")
	tagyn = request("tagyn")
	partwdid = request("partwdid")
	viewno = request("viewno")
	
	if page = "" then page = 1

'//이벤트 리스트
set oVideoClip = new CPlayContents
	oVideoClip.FPageSize = 50
	oVideoClip.FCurrPage = page
	oVideoClip.FRectstate = state
	oVideoClip.FRecttitle = viewtitle
	oVideoClip.FRPlaycate = playcate
	oVideoClip.FRectTag = tagyn
	oVideoClip.FRectNo = viewno
	oVideoClip.FRectpartWDid = partwdid
	oVideoClip.fnGetVideoClipList()
%>
<script type="text/javascript">
	function NextPage(page){
		frm.page.value = page;
		frm.submit();
	}

	function AddNewContents(idx){
		var popwin = window.open('/admin/sitemaster/play/videoclip/popvideoclipEdit.asp?idx=' + idx,'cateHotPosCodeEdit','width=800,height=500,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	function jsSerach(){
		var frm;
		frm = document.frm;
		frm.target = "_self";
		frm.action ="index.asp";
		frm.submit();
	}

	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function jsTagview(idx){	
		var poptag;
		poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+idx+'&playcate='+<%=playcate%>,'poptag','width=500,height=400,scrollbars=yes,resizable=yes');
		poptag.focus();
	}

	function jsSetItem(idx , styleno){
		var popitem;
		popitem = window.open('/admin/sitemaster/play/lib/pop_itemReg.asp?idx='+idx+'&number='+styleno,'popitem','width=500,height=400,scrollbars=yes,resizable=yes');
		popitem.focus();
	}
</script>

<form name="frm" method="post" style="margin:0px;">	
<input type="hidden" name="page" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	상태 : <% Draweventstate2 "state" , state ,"" %>
	&nbsp;&nbsp;&nbsp;
	번호 : <input type="text" name="viewno" value="<%=viewno%>" size="5"/>
	<!-- &nbsp;&nbsp;&nbsp;
	시작일 : <input type="text" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>" onClick="jsPopCal('reservationdate');"  style="cursor:pointer;"/> -->
	&nbsp;&nbsp;&nbsp;
	제목검색 : <input type="text" name="viewtitle" size=20 value="<%=viewtitle%>" />
	&nbsp;&nbsp;&nbsp;
	태그 : <select name="tagyn">
				<option value="">전체</option>
				<option value="Y" <%=chkiif(tagyn="Y","selected","")%>>등록</option>
				<option value="N" <%=chkiif(tagyn="N","selected","")%>>미등록</option>
			 </select>
	&nbsp;&nbsp;&nbsp;
	담당자WD : <% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach();">
	</td>
</tr>
</table>
</form>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="left">
		<font color="red"> ※ 리스트 노출 : 상태가 오픈인 것과 시작일 =< 오늘 인것만 노출이 됩니다. 순서는 No. 번호(높은순서) 순서로 노출됩니다.</font>		
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="AddNewContents('0');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				검색결과 : <b><%= oVideoClip.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> / <%=  oVideoClip.FTotalpage %></b>
			</td>
			<td align="right"></td>			
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="5%">idx</td>
	<td width="5%">No.</td>
	<td width="10%">제목</td>
	<td width="5%">상태(코드)</td>
	<td width="10%">타이틀이미지</td>
	<td width="10%">태그(개수)</td>
	<td width="10%">시작일</td>
	<td width="10%">담당WD</td>
	<td width="15%">비고</td>
</tr>
<% if oVideoClip.FresultCount > 0 then %>
<% for i=0 to oVideoClip.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center"><%= oVideoClip.FItemList(i).Fidx %></td>
	<td align="center"><%= oVideoClip.FItemList(i).Fviewno %></td>
	<td align="center"><%= oVideoClip.FItemList(i).Fviewtitle %></td>
	<td align="center"><%= geteventstate(oVideoClip.FItemList(i).Fstate) %> (<%=oVideoClip.FItemList(i).Fstate %>)</td>
	<td align="center"><img src="<%= oVideoClip.FItemList(i).Flistimg %>" width=70 border=0></td>
	<td align="center"><a href="#" onclick="jsTagview('<%= oVideoClip.FItemList(i).Fidx %>');" style="cursor:pointer;"><%=chkiif(oVideoClip.FItemList(i).Ftagcnt>0,"등록","미등록")%>(<%=oVideoClip.FItemList(i).Ftagcnt%>)</a></td>
	<td align="center"><%= left(oVideoClip.FItemList(i).Freservationdate,10) %></td>
	<td align="center"><%= oVideoClip.FItemList(i).FpartWDname %></td>
	<td align="center">
		<input type="button" class="button" value="수정" onclick="AddNewContents('<%= oVideoClip.FItemList(i).Fidx %>');"/>
	</td>
</tr>
<% Next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oVideoClip.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oVideoClip.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oVideoClip.StartScrollPage to oVideoClip.FScrollCount + oVideoClip.StartScrollPage - 1 %>
			<% if i>oVideoClip.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oVideoClip.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>


<% 
	set oVideoClip = nothing 
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
