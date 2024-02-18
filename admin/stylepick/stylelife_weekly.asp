<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
	Dim oWeekly, cd1,i,page,isusing ,oTheme ,state ,idx , title
	isusing = request("isusing")
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	idx = request("idx")
	title = request("title")
	
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"
		
'//이벤트 리스트
set oWeekly = new ClsStyleLife
	oWeekly.FPageSize = 50
	oWeekly.FCurrPage = page
	oWeekly.frectstate = state
	oWeekly.frectidx = idx
	oWeekly.frecttitle = title
	oWeekly.fnGetWeeklyList()
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

// 페이지 이동
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="stylelife_weekly.asp";
	document.frm.submit();
}

//등록 & 수정
function reg(idx){
	var weeklyreg = window.open('/admin/stylepick/stylelife_weekly_edit.asp?idx='+idx+'&menupos=<%=menupos%>','weeklyreg','width=1024,height=768,scrollbars=yes,resizable=yes');
	weeklyreg.focus();
}

//새상품추가
function addnewItem(idx){
	var weeklyitem = window.open('/admin/stylepick/stylelife_weekly_item.asp?idx='+idx+'','weeklyitem','width=500,height=900,scrollbars=yes,resizable=yes');
	weeklyitem.focus();
}

function goReal()
{
	if(confirm("실서버에 적용하시겠습니까?\n\n※ 시작일에 맞게 최신 3개 위클리가 보여집니다.") == true) {
		var stylelifemain = window.open('<%=wwwUrl%>/chtml/stylelife/make_stylelife_main.asp','stylelifemain','width=400,height=300');
		stylelifemain.focus();
	}
}
</script>

<!-- 액션 시작 -->
<form name="frm" method="get" style="margin:0px;">	
<input type="hidden" name="page" >
<input type="hidden" name="idxarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="left">
		<input type="button" class="button" value="StyleLife 메인 상단 적용하기" onClick="goReal()">
		<font color="red"> ※ 리스트 노출 : 상태가 오픈인 것과 시작일 =< 오늘 인것만 노출이 됩니다. 순서는 No. 번호(높은순서) 순서로 노출됩니다.</font>		
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="reg('');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<br><center><b><font size="5">위클리작업을 마지막에 한 사람은 위 버튼(Stylelife메인상단적용하기) 클릭하세요.적용을 하지 않으면 안나옵니다.</font></b><br><br></center>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				검색결과 : <b><%= oWeekly.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> /<%=  oWeekly.FTotalpage %></b>
			</td>
			<td align="right">
			</td>			
		</tr>
		</table>
	</td>
	
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><!--<input type="checkbox" name="chkAll" onClick="jsChkAll();">//--></td>
	<td>No.</td>
	<td>제목</td>
	<td>상태(코드)</td>
	<td>타이틀이미지</td>
	<td>시작일</td>
	<td>담당자</td>
	<td>기획WD</td>
	<td>비고</td>
</tr>
<% if oWeekly.FresultCount > 0 then %>
<% for i=0 to oWeekly.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center">
		<!--<input type="checkbox" name="chkitem" value="<%= oWeekly.FItemList(i).Fidx %>">//-->
	</td>
	<td align="center">		
		<%= oWeekly.FItemList(i).Fidx %>
	</td>
	<td align="center"><%= oWeekly.FItemList(i).ftitle %></td>
	<td align="center"><%= geteventstate(oWeekly.FItemList(i).fstatename) %> (<%=oWeekly.FItemList(i).fstate %>)</td>
	<td align="center"><img src="<%= oWeekly.FItemList(i).ftitle_img %>" width=200 border=0></td>
	<td align="center"><%= left(oWeekly.FItemList(i).fstartdate,10) %></td>
	<td align="center"><%= oWeekly.FItemList(i).fpartMDname %></td>
	<td align="center"><%= oWeekly.FItemList(i).fpartwDname %></td>
	<td align="center">
		<input type="button" class="button" value="수정" onclick="reg('<%= oWeekly.FItemList(i).Fidx %>');">
		<input type="button" value="상품추가[<%= oWeekly.FItemList(i).fitemcnt %>]" onclick="addnewItem('<%= oWeekly.FItemList(i).Fidx %>');" class="button">
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oWeekly.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oWeekly.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oWeekly.StartScrollPage to oWeekly.FScrollCount + oWeekly.StartScrollPage - 1 %>
			<% if i>oWeekly.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oWeekly.HasNextScroll then %>
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
</form>

<% set oWeekly = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->