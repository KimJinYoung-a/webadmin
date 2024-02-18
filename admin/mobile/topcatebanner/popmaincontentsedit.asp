<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : popMainContentsEdit.asp
' Discription : 모바일 사이트 메인 아이템 작성/수정
' History : 2010.02.23 허진원 생성
'           2012.02.14 허진원 - 미니달력 교체
'           2012.12.14 이종화 - alt 값 추가
'           2015-09-17 이종화 - 모바일 카테고리용
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/mobile/TopcateManageCls.asp" -->
<!-- #include virtual="/lib/classes/mobile/topcatecodeCls.asp" -->
<%
dim idx, poscode, reload , gcode
	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	gcode = request("gcode")
	if idx="" then idx=0

	if reload="on" then
	    response.write "<script>opener.location.reload(); window.close();</script>"
	    dbget.close()	:	response.End    
	end if

	dim oMainContents
		set oMainContents = new CMainContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneMainContents

dim oposcode, defaultMapStr
	set oposcode = new CMainContentsCode
	oposcode.FRectPosCode = poscode
	if poscode<>"" then
	    oposcode.GetOneContentsCode
	    
	    defaultMapStr = "<map name='Map_" +oposcode.FOneItem.FPosvarName + "'>" + VbCrlf
	    defaultMapStr = defaultMapStr + VbCrlf
	    defaultMapStr = defaultMapStr + "</map>"
	end if

dim orderidx
	if oMainContents.FOneItem.forderidx = "" then
		orderidx = 99
	else
		orderidx = oMainContents.FOneItem.forderidx
	end if
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

	function SaveMainContents(frm){
	    if (frm.poscode.value.length<1){
	        alert('구분을 먼저 선택 하세요.');
	        frm.poscode.focus();
	        return;
	    }

		if (frm.gcode.value.length<1){
	        alert('GNB 노출영역 먼저 선택 하세요.');
	        frm.gcode.focus();
	        return;
	    }
	    
	    if (frm.linkurl.value.length<1){
	        alert('링크 값을 입력 하세요.');
	        frm.linkurl.focus();
	        return;
	    }
	    <% If poscode = "1003" or oMainContents.FOneItem.Fposcode = "1003" Then %>
	    if (frm.backColor.value.length<1){
	        alert('배경색을 등록하세요');
	        frm.backColor.focus();
	        return;
	    }
		<% End If %>
	    if (frm.startdate.value.length!=10){
	        alert('시작일을 입력  하세요.');
	        return;
	    }
	    
	    if (frm.enddate.value.length!=10){
	        alert('종료일을 입력  하세요.');
	        return;
	    }
	    
	    var vstartdate = new Date(frm.startdate.value.substr(0,4), (1*frm.startdate.value.substr(5,2))-1, frm.startdate.value.substr(8,2));
	    var venddate = new Date(frm.enddate.value.substr(0,4), (1*frm.enddate.value.substr(5,2))-1, frm.enddate.value.substr(8,2));
	    
	    if (vstartdate>venddate){
	        alert('종료일이 시작일보다 빠르면 안됩니다.');
	        return;
	    }

		if (frm.altname.value.length<1){
	        alert('알트명을 입력  하세요.');
			frm.altname.focus();
	        return;
	    }
	
	    if (confirm('저장 하시겠습니까?')){
	        frm.submit();
	    }
	}
	
	function ChangeLinktype(comp){
	    if (comp.value=="M"){
	       document.all.link_M.style.display = "";
	       document.all.link_L.style.display = "none";
	    }else{
	       document.all.link_M.style.display = "none";
	       document.all.link_L.style.display = "";
	    }
	}
	
	//function getOnLoad(){
	//    ChangeLinktype(frmcontents.linktype.value);
	//}
	
	//window.onload = getOnLoad;
	
	function ChangeGubun(comp){
	    location.href = "?poscode=" + comp.value;
	    // nothing;
	}

	function putLinkText(key) {
		var frm = document.frmcontents;
		switch(key) {
			case 'search':
				frm.linkurl.value='/search/search_result.asp?rect=검색어';
				break;
			case 'event':
				frm.linkurl.value='/event/eventmain.asp?eventid=이벤트번호';
				break;
			case 'itemid':
				frm.linkurl.value='/category/category_itemprd.asp?itemid=상품코드';
				break;
			case 'category':
				frm.linkurl.value='/category/category_list.asp?cdl=카테고리';
				break;
			case 'brand':
				frm.linkurl.value='/street/street_brand.asp?makerid=브랜드아이디';
				break;
		}
	}
	
</script>

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/doMobileCateContentsReg.asp" onsubmit="return false;" enctype="multipart/form-data">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">구분명</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fposname %> (<%= oMainContents.FOneItem.Fposcode %>)
        <input type="hidden" name="poscode" value="<%= oMainContents.FOneItem.Fposcode %>">
        <% else %>
        <% call DrawMainPosCodeCombo("poscode", poscode, "onChange='ChangeGubun(this);'") %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">GNB 노출</td>
    <td>
        <% if oMainContents.FOneItem.Fgnbcode<>"" then %>
        <%= oMainContents.FOneItem.Fgnbname %> (<%= oMainContents.FOneItem.Fgnbcode %>)
        <input type="hidden" name="gcode" value="<%= oMainContents.FOneItem.Fgnbcode %>">
        <% else %>
        <% Call drawSelectBoxGNB("gcode" , gcode) %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크구분</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.getlinktypeName %>
        <input type="hidden" name="linktype" value="<%= oMainContents.FOneItem.Flinktype %>">
        <% else %>
            <% if poscode<>"" then %>
            <%= oposcode.FOneItem.getlinktypeName %>
            <input type="hidden" name="linktype" value="<%= oposcode.FOneItem.Flinktype %>">
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">적용구분(반영주기)</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.getfixtypeName %>
        <input type="hidden" name="fixtype" value="<%= oMainContents.FOneItem.Ffixtype %>">
        <% else %>
            <% if poscode<>"" then %>
            <%= oposcode.FOneItem.getfixtypeName %>
            <input type="hidden" name="fixtype" value="<%= oposcode.FOneItem.Ffixtype %>">
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">우선순위</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        	<% if oMainContents.FOneItem.Flinktype="X" then %>
            <input type="text" name="orderidx" size=5 value="<%= orderidx %>">
        	<% end if %>
        <% else %>
            <% if poscode<>"" then %>
            	<input type="text" name="orderidx" size=5 value="<%= orderidx %>">
            <% else %>
            	<font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이미지</td>
  <td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <br>
  <img src="<%= oMainContents.FOneItem.GetImageUrl %>" width="300">
  <br> <%= oMainContents.FOneItem.GetImageUrl %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">알트명 (필수)</td>
  <td><input type="text" name="altname" value="<%=oMainContents.FOneItem.Faltname%>" size="20" maxlength="20"> </td>
</tr>
<% If False Then 'oMainContents.FOneItem.Fposcode = "1" Or poscode = "1" Then %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">서브타이틀<br/>(사용안함 예정)</td>
  <td>
	<input type="text" name="texttitle1" value="<%=oMainContents.FOneItem.Ftexttitle1%>" size="30" maxlength="30"></br>
	<input type="text" name="texttitle2" value="<%=oMainContents.FOneItem.Ftexttitle2%>" size="30" maxlength="30">
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이벤트할인<br/>(사용안함 예정)</td>
  <td>할인 : <input type="radio" name="saleflag" value="1" <%=chkiif(oMainContents.FOneItem.Fsaleflag="1","checked","")%>/> 쿠폰 : <input type="radio" name="saleflag" value="2" <%=chkiif(oMainContents.FOneItem.Fsaleflag="2","checked","")%>/> <input type="text" name="saletext" value="<%=oMainContents.FOneItem.Fsaletext%>" size="10" maxlength="10"></td>
</tr>
<% End If %>

<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이미지Width</td>
  <td>
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <input type="text" name="imagewidth" value="<%= oMainContents.FOneItem.Fimagewidth %>" size="8" maxlength="16"> 
  <% else %>
        <% if poscode<>"" then %>
        <%= oposcode.FOneItem.Fimagewidth %>
        <% else %>
        <font color="red">구분을 먼저 선택하세요</font>
        <% end if %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이미지Height</td>
  <td>
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <input type="text" name="imageheight" value="<%= oMainContents.FOneItem.Fimageheight %>" size="8" maxlength="16"> 
  <% else %>
        <% if poscode<>"" then %>
        <%= oposcode.FOneItem.Fimageheight %>
        <% else %>
        <font color="red">구분을 먼저 선택하세요</font>
        <% end if %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">작업자 지시사항</td>
	<td ><textarea name="ordertext" cols="80" rows="8"/><%=oMainContents.FOneItem.Fordertext%></textarea></td>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크값</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
            <% if oMainContents.FOneItem.FLinkType="M" then %>
            <textarea name="linkurl" cols="60" rows="6"><%= oMainContents.FOneItem.Flinkurl %></textarea>
            <% else %>
            <input type="text" name="linkurl" value="<%= oMainContents.FOneItem.Flinkurl %>" maxlength="128" style="width:100%">
            <% end if %>
			<font color="#707070">
			- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
			- <span style="cursor:pointer" onClick="putLinkText('event')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
			- <span style="cursor:pointer" onClick="putLinkText('itemid')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
			- <span style="cursor:pointer" onClick="putLinkText('category')">카테고리 링크 : /category/category_list.asp?cdl=<font color="darkred">카테고리</font></span><br>
			- <span style="cursor:pointer" onClick="putLinkText('brand')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
			</font>
        <% else %>
            <% if poscode<>"" then %>
                <% if oposcode.FOneItem.FLinkType="M" then %>
                    <textarea name="linkurl" cols="60" rows="6"><%= defaultMapStr %></textarea>
                    <br>(이미지맵 변수값 변경 금지)
                <% else %>
                    <input type="text" name="linkurl" value="" maxlength="128" style="width:100%">
                    <br>
					<font color="#707070">
					- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
					- <span style="cursor:pointer" onClick="putLinkText('event')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('itemid')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('category')">카테고리 링크 : /category/category_list.asp?cdl=<font color="darkred">카테고리</font></span><br>
					- <span style="cursor:pointer" onClick="putLinkText('brand')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
					</font>
                <% end if %>
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">반영시작일</td>
    <td>
		<input id="startdate" name="startdate" value="<%=Left(oMainContents.FOneItem.Fstartdate,10)%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<% if oMainContents.FOneItem.Fidx<>"" then %>
			<% if oMainContents.FOneItem.Ffixtype="R" then %> 
			<input type="text" name="startdatetime" size="2" maxlength="2" value="<%= Format00(2,Hour(oMainContents.FOneItem.Fstartdate)) %>" />(시 00~23)
			<input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
			<% else %>
			<input type="text" name="dummy0" value="00:00:00" size="8" readonly class="text_ro" />
			<% end if %>
        <% else %>
            <% if poscode<>"" then %>
				<% if oposcode.FOneItem.Ffixtype="R" then %> 
				<input type="text" name="startdatetime" size="2" maxlength="2" value="<%= Format00(2,Hour(oMainContents.FOneItem.Fstartdate)) %>" />(시 00~23)
				<input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
				<% else %>
				<input type="text" name="dummy0" value="00:00:00" size="8" readonly class="text_ro" />
				<% end if %>
            <% end if %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">반영종료일</td>
    <td>
		<input id="enddate" name="enddate" value="<%=Left(oMainContents.FOneItem.Fenddate,10)%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<% if oMainContents.FOneItem.Fidx<>"" then %>
			<% if oMainContents.FOneItem.Ffixtype="R" then %> 
			<input type="text" name="enddatetime" size="2" maxlength="2" value="<%= ChkIIF(oMainContents.FOneItem.Fenddate="","23",Format00(2,Hour(oMainContents.FOneItem.Fenddate))) %>">(시 00~23)
			<input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
			<% else %>
			<input type="text" name="dummy1" value="23:59:59" size="8" readonly class="text_ro" />
			<% end if %>
        <% else %>
            <% if poscode<>"" then %>
				<% if oposcode.FOneItem.Ffixtype="R" then %> 
				<input type="text" name="enddatetime" size="2" maxlength="2" value="<%= ChkIIF(oMainContents.FOneItem.Fenddate="","23",Format00(2,Hour(oMainContents.FOneItem.Fenddate))) %>">(시 00~23)
				<input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
				<% else %>
				<input type="text" name="dummy1" value="23:59:59" size="8" readonly class="text_ro" />
				<% end if %>
            <% end if %>
        <% end if %>
		<script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "startdate", trigger    : "startdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "enddate", trigger    : "enddate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">등록일</td>
    <td>
        <%= oMainContents.FOneItem.Fregdate %> (<%= oMainContents.FOneItem.Freguserid %>)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">사용여부</td>
    <td>
        <% if oMainContents.FOneItem.Fisusing="N" then %>
        <input type="radio" name="isusing" value="Y">사용함
        <input type="radio" name="isusing" value="N" checked >사용안함
        <% else %>
        <input type="radio" name="isusing" value="Y" checked >사용함
        <input type="radio" name="isusing" value="N">사용안함
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<%
set oposcode = Nothing
set oMainContents = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->