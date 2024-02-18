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
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
Dim cd1,i,page,isusing ,oTheme ,state ,idx , title
	cd1 = request("cd1")
	isusing = request("isusing")
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	idx = request("idx")
	title = request("title")
	
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"
		
'//이벤트 리스트
set oTheme = new ClsStyleLife
	oTheme.FPageSize = 50
	oTheme.FCurrPage = page
	oTheme.frectcd1 = cd1
	oTheme.frectstate = state
	oTheme.frectisusing = isusing
	oTheme.frectidx = idx
	oTheme.frecttitle = title
	oTheme.fnGetThemeList()
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
	frm.action ="stylelife_theme.asp";
	frm.submit();
}

// 페이지 이동
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="stylelife_theme.asp";
	document.frm.submit();
}

//이벤트 등록 & 수정
function eventedit(idx){
	var eventedit = window.open('/admin/stylepick/stylelife_theme_edit.asp?idx='+idx+'&menupos=<%=menupos%>','eventedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	eventedit.focus();
}

//새상품추가
function addnewItem(idx){
	location.href="/admin/stylepick/stylelife_theme_item.asp?idx="+idx+"&menupos=<%=menupos%>";
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">	
<input type="hidden" name="page" >
<input type="hidden" name="idxarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<select name="cd1" onchange='jsSerach();'>
			<option value="">-스타일-</option>
			<option value="010" <%=CHKIIF(cd1="010","selected","")%>>클래식</option>
			<option value="020" <%=CHKIIF(cd1="020","selected","")%>>큐트</option>
			<option value="040" <%=CHKIIF(cd1="040","selected","")%>>모던</option>
			<option value="050" <%=CHKIIF(cd1="050","selected","")%>>네추럴</option>
			<option value="060" <%=CHKIIF(cd1="060","selected","")%>>오리엔탈</option>
			<option value="070" <%=CHKIIF(cd1="070","selected","")%>>팝</option>
			<option value="080" <%=CHKIIF(cd1="080","selected","")%>>로맨틱</option>
			<option value="090" <%=CHKIIF(cd1="090","selected","")%>>빈티지</option>
			<option value="0P0" <%=CHKIIF(cd1="0P0","selected","")%>>스타일픽</option>
		</select>
		상태 : <% Draweventstate2 "state" , state ," onchange='jsSerach();'" %>		
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		테마번호 : <input type="text" name="idx" value="<%= idx %>" size=10>
		제목 : <input type="text" name="title" value="<%= title %>" size=30>
	</td>
</tr>    
</table>
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="left">
		<font color="red"> ※ 리스트 노출 순서(실제 오픈일때) : 1. 순서(높은번호순), 2. 테마번호(높은번호순), 3. 시작일(최근순) 순서로 노출됩니다</font>
		<br>오픈되었을 경우 <b>순서 숫자가 절대 같으면 안됨!!</b> 에러남.
	</td>
	<td align="right">
		<input type="button" class="button" value="테마신규등록" onclick="eventedit('');">
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
				검색결과 : <b><%= oTheme.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> /<%=  oTheme.FTotalpage %></b>				
			</td>
			<td align="right">
			</td>			
		</tr>
		</table>
	</td>
	
</tr>
		
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><!--<input type="checkbox" name="chkAll" onClick="jsChkAll();">//--></td>
	<td>테마번호</td>
	<td>스타일</td>
	<td>상태(코드)</td>
	<td>배너이미지</td>
	<td>제목</td>
	<td>시작일</td>
	<td>오픈날짜</td>
	<td>기획MD</td>
	<td>기획WD</td>
	<td>순서</td>
	<td>비고</td>
</tr>
<% if oTheme.FresultCount > 0 then %>
<% for i=0 to oTheme.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center">
		<!--<input type="checkbox" name="chkitem" value="<%= oTheme.FItemList(i).Fidx %>">//-->
	</td>
	<td align="center">		
		<%= oTheme.FItemList(i).Fidx %><br><a href="<%=wwwUrl%>/stylelife/theme/view.asp?idx=<%= oTheme.FItemList(i).Fidx %>&isadmin=admin" onfocus="this.blur()" target="_blink">[미리보기]</a>
	</td>
	
	<td align="center"><%= CHKIIF(oTheme.FItemList(i).fcatename="","STYLE PICK",oTheme.FItemList(i).fcatename) %></td>
	<td align="center"><%= geteventstate(oTheme.FItemList(i).fstatename) %> (<%=oTheme.FItemList(i).fstate %>)</td>
	<td align="center"><img src="<%= oTheme.FItemList(i).fbanner_img %>" width=50 height=50 border=0></td>
	<td align="center"><%= oTheme.FItemList(i).ftitle %></td>
	<td align="center"><%= left(oTheme.FItemList(i).fstartdate,10) %></td>
	<td align="center">
		<% 
		if oTheme.FItemList(i).fopendate <> "1900-01-01" then response.write oTheme.FItemList(i).fopendate
		'if oTheme.FItemList(i).fclosedate <> "1900-01-01" then response.write oTheme.FItemList(i).fclosedate
		%>
	</td>	
	<td align="center"><%= oTheme.FItemList(i).fpartMDname %></td>
	<td align="center"><%= oTheme.FItemList(i).fpartwDname %></td>
	<td align="center"><%= oTheme.FItemList(i).fsortno %></td>
	<td align="center">
		<input type="button" class="button" value="수정" onclick="eventedit('<%= oTheme.FItemList(i).Fidx %>');">
		<input type="button" value="상품추가[<%= oTheme.FItemList(i).fitemcnt %>]" onclick="addnewItem('<%= oTheme.FItemList(i).Fidx %>');" class="button">
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oTheme.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oTheme.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oTheme.StartScrollPage to oTheme.FScrollCount + oTheme.StartScrollPage - 1 %>
			<% if i>oTheme.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oTheme.HasNextScroll then %>
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

<% set oTheme = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->