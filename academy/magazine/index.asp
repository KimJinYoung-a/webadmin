<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description :  핑거스 아카데미 매거진 어드민 리스트
' History : 2016-03-03 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/academy/magazineCls.asp" -->
<%
	Dim oMagaZine, i , page , state ,idx , startdate , viewtitle , viewno
	Dim catecode'' : catecode = 6 'videoclip
	menupos = RequestCheckvar(request("menupos"),10)
	page = RequestCheckvar(request("page"),10)
	state = RequestCheckvar(request("state"),10)
	startdate = RequestCheckvar(request("startdate"),10)
	viewtitle = request("viewtitle")
	viewno = RequestCheckvar(request("viewno"),10)
	catecode = RequestCheckvar(request("catecode"),10)
	
	if page = "" then page = 1
  	if viewtitle <> "" then
		if checkNotValidHTML(viewtitle) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
set oMagaZine = new CMagaZineContents
	oMagaZine.FPageSize = 20
	oMagaZine.FCurrPage = page
	oMagaZine.FRectstate = state
	oMagaZine.FRectviewtitle = viewtitle
	oMagaZine.FRectcatecode = catecode
	oMagaZine.FRectviewno = viewno
	oMagaZine.fnGetMagazineList()
%>
<script type="text/javascript">
	function NextPage(page){
		frm.page.value = page;
		frm.submit();
	}

	function AddNewContents(idx){
		var popwin = window.open('/academy/magazine/popmagazineEdit.asp?idx=' + idx,'magazineEdit','width=700,height=800,scrollbars=yes,resizable=yes');
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


	//이미지 확대 새창으로 보여주기
	function showimage(img){
		var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
	}

	//''카테고리 관리
	function jsCatecodeview(idx){
		var poptag;
		poptag = window.open('/academy/magazine/lib/pop_catecodeReg.asp','popcatecode','width=300,height=400,scrollbars=yes,resizable=yes');
		poptag.focus();
	}

	//미리보기
	function jsOpen(sPURL,sTG){ 
	    if (sTG =="M" ){ 
	        var winView = window.open(sPURL,"popView","width=400, height=600,scrollbars=yes,resizable=yes");
	    }
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
	구분 : <% DrawMagazineGubun "catecode" , catecode ,"" %>
	&nbsp;&nbsp;&nbsp;
	번호 : <input type="text" name="viewno" value="<%=viewno%>" size="5"/>
	<!-- &nbsp;&nbsp;&nbsp;
	시작일 : <input type="text" name="startdate" size=20 maxlength=10 value="<%=startdate%>" onClick="jsPopCal('startdate');"  style="cursor:pointer;"/> -->
	&nbsp;&nbsp;&nbsp;
	제목검색 : <input type="text" name="viewtitle" size=20 value="<%=viewtitle%>" />
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onclick="javascript:jsSerach();">
	</td>
</tr>
</table>
</form>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="left">
		<font color="red"> ※ 리스트 노출 : 상태가 오픈이고 시작일 =< 오늘 인것만 노출이 됩니다. 순서는 No. 번호(높은순서) 순서로 노출됩니다.</font>		
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="AddNewContents('0');">
		<input type="button" name="btnviewImg" value="카테고리 관리" onClick="jsCatecodeview();" class="button"/>
	</td>
	
</tr>
<tr>
	<td><br><br>
		<a href="javascript:jsOpen('<%= mobFingers %>/magazine/','M');" ><font color="red"><b>프론트 링크</b></font></a>
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
				검색결과 : <b><%= oMagaZine.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> / <%=  oMagaZine.FTotalpage %></b>
			</td>
			<td align="right"></td>			
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="3%">idx</td>
	<td width="3%">No.</td>
	<td width="5%">구분</td>
	<td width="15%">제목</td>
	<td width="5%">상태(코드)</td>
	<td width="15%">리스트이미지</td>
	<td width="15%">상세이미지</td>
	<td width="5%">시작일</td>
	<td width="5%">비고</td>
</tr>
<% if oMagaZine.FresultCount > 0 then %>
<% for i=0 to oMagaZine.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center"><%= oMagaZine.FItemList(i).Fidx %></td>
	<td align="center"><%= oMagaZine.FItemList(i).Fviewno %></td>
	<td align="center"><%= getMagazinecatecode(oMagaZine.FItemList(i).Fcatecode) %></td>
	<td align="center"><%= oMagaZine.FItemList(i).Fviewtitle %></td>
	<td align="center"><%= geteventstate(oMagaZine.FItemList(i).Fstate) %> (<%=oMagaZine.FItemList(i).Fstate %>)</td>
	<td align="center"><img src="<%= oMagaZine.FItemList(i).Flistimg %>" width=70 border=0 onclick="showimage('<%= oMagaZine.FItemList(i).Flistimg %>');" onerror="this.src='http://webimage.10x10.co.kr/academy/magazine/noimg.jpg'" style="cursor:pointer;"></td>
	<td align="center"><img src="<%= oMagaZine.FItemList(i).Fviewimg1 %>" width=70 border=0 onclick="showimage('<%= oMagaZine.FItemList(i).Fviewimg1 %>');" onerror="this.src='http://webimage.10x10.co.kr/academy/magazine/noimg.jpg'" style="cursor:pointer;"></td>
	<td align="center"><%= left(oMagaZine.FItemList(i).Fstartdate,10) %></td>
	<td align="center"><input type="button" class="button" value="수정" onclick="AddNewContents('<%= oMagaZine.FItemList(i).Fidx %>');"/></td>
</tr>
<% Next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oMagaZine.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oMagaZine.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oMagaZine.StartScrollPage to oMagaZine.FScrollCount + oMagaZine.StartScrollPage - 1 %>
			<% if i>oMagaZine.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oMagaZine.HasNextScroll then %>
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
<% set oMagaZine = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->