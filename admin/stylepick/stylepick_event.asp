<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 스타일픽 관리
' Hieditor : 2011.04.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->

<%
Dim cd1,i,page,isusing ,oevent ,state ,evtidx , title
	cd1 = request("cd1")
	isusing = request("isusing")
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	evtidx = request("evtidx")
	title = request("title")
	
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"
		
'//이벤트 리스트
set oevent = new cstylepick
	oevent.FPageSize = 50
	oevent.FCurrPage = page
	oevent.frectcd1 = cd1
	oevent.frectstate = state
	oevent.frectisusing = isusing
	oevent.frectevtidx = evtidx
	oevent.frecttitle = title
	oevent.fnGetEventList()
%>

<script language="javascript">

//전체 선택
function jsChkAll(){	
var frm;
frm = document.frm;
	if (frm.chkAll.checked){			      
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
		   	 	frm.chkitem.checked = true;	   	 
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					frm.chkitem[i].checked = true;
			 	}		
		   }	
	   }	
	} else {	  
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;	  
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}	
		}		
	  }	
	}
}

function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="stylepick_event.asp";
	frm.submit();
}

// 페이지 이동
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="stylepick_event.asp";
	document.frm.submit();
}

//이벤트 등록 & 수정
function eventedit(evtidx){
	var eventedit = window.open('/admin/stylepick/stylepick_event_edit.asp?evtidx='+evtidx+'&menupos=<%=menupos%>','eventedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	eventedit.focus();
}

//새상품추가
function addnewItem(evtidx){
	location.href="/admin/stylepick/stylepick_event_item.asp?evtidx="+evtidx+"&menupos=<%=menupos%>";
}
	
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="page" >
<input type="hidden" name="evtidxarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		스타일 : <% Drawcategory "cd1",cd1," onchange='jsSerach();'","CD1" %>
		사용 : <% drawSelectBoxUsingYN "isusing", isusing %>
		상태 : <% Draweventstate2 "state" , state ," onchange='jsSerach();'" %>		
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		기획전번호 : <input type="text" name="evtidx" value="<%= evtidx %>" size=10>
		제목 : <input type="text" name="title" value="<%= title %>" size=30>
	</td>
</tr>    
</table>
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red"> ※상태가 "오픈"이고 , 현재날짜가 시작일보다 크면 프론트에 최근 등록순으로 노출됩니다</font>		
	</td>
	<td align="right">
		<input type="button" class="button" value="기획전신규등록" onclick="eventedit('');">
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
				검색결과 : <b><%= oevent.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> /<%=  oevent.FTotalpage %></b>				
			</td>
			<td align="right">
			</td>			
		</tr>
		</table>
	</td>	
</tr>		
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td>기획전번호</td>
	<td>스타일</td>
	<td>상태(코드)</td>
	<td>배너이미지</td>
	<td>제목</td>
	<td>기간</td>
	<td>오픈날짜<br>종료날짜</td>
	<td>기획MD</td>
	<td>기획WD</td>
	<td>비고</td>
</tr>
<% if oevent.FresultCount > 0 then %>
<% for i=0 to oevent.FresultCount-1 %>
<% if oevent.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<input type="checkbox" name="chkitem" value="<%= oevent.FItemList(i).Fevtidx %>">
	</td>
	<td align="center">		
		<a href="/admin/stylepick/stylepick_collect_testview.asp?evtidx=<%= oevent.FItemList(i).Fevtidx %>" onfocus="this.blur()" target="_blink">
		<%= oevent.FItemList(i).Fevtidx %> [미리보기]</a>
	</td>	
	<td align="center"><%= oevent.FItemList(i).fcatename %></td>
	<td align="center"><%= geteventstate(oevent.FItemList(i).fstatename) %> (<%=oevent.FItemList(i).fstate %>)</td>
	<td align="center"><img src="<%= oevent.FItemList(i).fbanner_img %>" width=50 height=50 border=0></td>
	<td align="center"><%= oevent.FItemList(i).ftitle %></td>
	<td align="center"><%= left(oevent.FItemList(i).fstartdate,10) %><Br>~ <%= left(oevent.FItemList(i).fenddate,10) %></td>
	<td align="center">
		<% 
		if oevent.FItemList(i).fopendate <> "1900-01-01" then response.write oevent.FItemList(i).fopendate
		if oevent.FItemList(i).fclosedate <> "1900-01-01" then response.write oevent.FItemList(i).fclosedate
		%>
	</td>	
	<td align="center"><%= oevent.FItemList(i).fpartMDname %></td>
	<td align="center"><%= oevent.FItemList(i).fpartwDname %></td>
	<td align="center">
		<input type="button" class="button" value="수정" onclick="eventedit('<%= oevent.FItemList(i).Fevtidx %>');">
		<input type="button" value="상품추가[<%= oevent.FItemList(i).fevtitemcnt %>]" onclick="addnewItem('<%= oevent.FItemList(i).Fevtidx %>');" class="button">
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oevent.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oevent.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oevent.StartScrollPage to oevent.FScrollCount + oevent.StartScrollPage - 1 %>
			<% if i>oevent.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oevent.HasNextScroll then %>
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
</form>
</table>

<% set oevent = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->