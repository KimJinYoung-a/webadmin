<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 컬러트랜드 관리
' Hieditor : 2012.03.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/color/colortrend_cls.asp"-->

<%
Dim ctcode,i,page,isusing ,ocolor ,state ,iColorCd , partwdid , partmdid , viewno
	iColorCd = request("iCD")
	ctcode = request("ctcode")
	isusing = request("isusing")
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	partwdid = request("partwdid")
	partmdid = request("partmdid")
	viewno = request("viewno")
	
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"

'//리스트
set ocolor = new ccolortrend_list
	ocolor.FPageSize = 50
	ocolor.FCurrPage = page
	ocolor.frectctcode = ctcode
	ocolor.frectcolorcode = iColorCd
	ocolor.frectstate = state
	ocolor.frectisusing = isusing
	ocolor.frectviewno = viewno
	ocolor.frectpartwdid = partwdid
	ocolor.frectpartmdid = partmdid
	ocolor.getcolortrend()
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

function jsSerach(ipage){
	var frm;
	frm = document.frm;
	
	if(frm.ctcode.value!=''){
		if (!IsDouble(frm.ctcode.value)){
			alert('컬러트랜드 코드는 숫자만 가능합니다.');
			frm.ctcode.focus();
			return;
		}
	}

	frm.page.value= ipage;
	frm.submit();
}

//등록 & 수정
function popedit(ctcode){
	var popedit = window.open('/admin/itemmaster/colortrend_edit.asp?ctcode='+ctcode+'&menupos=<%=menupos%>','popedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	popedit.focus();
}

//색상코드 선택
function selColorChip(cd) {
	document.frm.iCD.value= cd;
	document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="page" value=1>
<input type="hidden" name="iCD" value="<%=iColorCd%>">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		컬러트렌드코드 : <input type="text" name="ctcode" value="<%=ctcode%>" size=10/>
		&nbsp;No. : <input type="text" name="viewno" size="5" value="<%=viewno%>"/>
		&nbsp;사용 : <% drawSelectBoxUsingYN "isusing", isusing %>
		&nbsp;상태 : <% Drawcolortrendstate "state" , state ," onchange='jsSerach("""");'" %>
		&nbsp;담당자 : <% sbGetpartid "partmdid",partmdid,"","23" %>
		&nbsp;담당자WD : <% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach('');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		<%=FnSelectColorBar(iColorCd,32)%>
	</td>
</tr>    
</table>
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※상태가 "오픈" 컬러중 최근 내역이 프론트에 이주의 컬러로 노출 됩니다.
		<br>현재 진행중인 컬러가 없을경우 날짜가 지난 컬러중 최근내역이 이주의 컬러로 노출 됩니다.
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="popedit('');">
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
				검색결과 : <b><%= ocolor.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> /<%=  ocolor.FTotalpage %></b>				
			</td>
			<td align="right">
			</td>			
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td>컬러트렌드<br>코드</td>
	<td>No.</td>
	<td>컬러칩</td>
	<td>상태(코드)</td>
	<td>제목</td>
	<td>시작일</td>
	<td>최근수정</td>
	<td>담당자</td>
	<td>담당자WD</td>
	<td>비고</td>
</tr>
<% if ocolor.FresultCount > 0 then %>
<% for i=0 to ocolor.FresultCount-1 %>
<% if ocolor.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>    
<tr align="center" bgcolor="#FFFFaa">
<% end if %>
	<td align="center">
		<input type="checkbox" name="chkitem" value="<%= ocolor.FItemList(i).fctcode %>">
	</td>
	<td align="center">
		<%= ocolor.FItemList(i).fctcode %>
		<% if ocolor.FItemList(i).fthisweek = ocolor.FItemList(i).fctcode then %>
			&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2012/colortrend/ico_week.png" width=40 height=40>
		<% end if %>
	</td>
	<td align="center">
		<%= ocolor.FItemList(i).Fviewno %>
	</td>
	<td align="center">
		<img src="<%=ocolor.FItemList(i).FcolorIcon%>" width="20" height="20" alt="<%=ocolor.FItemList(i).fcolorName%>">
	</td>
	<td align="center">
		<%= getcolortrendstate(ocolor.FItemList(i).fstatename) %>
	</td>
	<td align="center"><%=ocolor.FItemList(i).Fcolortitle%></td>
	<td align="center">
		<%= left(ocolor.FItemList(i).fstartdate,10) %>
	</td>	
	<td align="center">
		<%= ocolor.FItemList(i).flastadminid %>
		<Br><%= ocolor.FItemList(i).flastupdate %>	
	</td>
	<td align="center">
		<%= ocolor.FItemList(i).FpartmdName %>
	</td>
	<td align="center">
		<%= ocolor.FItemList(i).FpartwdName %>
	</td>
	<td align="center">
		<input type="button" class="button" value="수정" onclick="popedit('<%= ocolor.FItemList(i).fctcode %>');">		
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if ocolor.HasPreScroll then %>
			<a href="javascript:jsSerach('<%= ocolor.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + ocolor.StartScrollPage to ocolor.FScrollCount + ocolor.StartScrollPage - 1 %>
			<% if i>ocolor.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:jsSerach('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if ocolor.HasNextScroll then %>
			<a href="javascript:jsSerach('<%= i %>')">[next]</a>
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

<% set ocolor = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->