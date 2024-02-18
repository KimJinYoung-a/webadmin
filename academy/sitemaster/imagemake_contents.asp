<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 코너관리
' History : 2009.09.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/sitemaster/sitemaster_cls.asp"-->

<%
dim idx, poscode,reload , ix, gubun
Dim sDt, sTm, eDt, eTm
dim srcSDT , srcEDT, sdate, edate, stdt, eddt
	idx = RequestCheckvar(request("idx"),10)
	gubun = RequestCheckvar(request("gubun"),24)
	poscode = RequestCheckvar(request("poscode"),10)
	srcSDT			=	RequestCheckvar(request("sDt"),10)
	srcEDT			=	RequestCheckvar(request("eDt"),10)
	sdate			=	RequestCheckvar(request("sdate"),10)
	edate			=	RequestCheckvar(request("edate"),10)
	reload = RequestCheckvar(request("reload"),2)
	if idx="" then idx=0
	if gubun="" then gubun="index"

if reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbACADEMYget.close() : dbget.close()	:	response.End
end if

dim oMainContents
	set oMainContents = new cposcode_list
	oMainContents.FRectIdx = idx
	oMainContents.fcontents_oneitem

dim oposcode, defaultMapStr
	set oposcode = new cposcode_list
	oposcode.FRectPosCode = poscode
	if poscode<>"" then
	    oposcode.fposcode_oneitem
	end if

if poscode = "" then
	poscode= oMainContents.FOneItem.Fposcode
end if 

if oMainContents.FOneItem.Fidx<>"" then
	sdate = oMainContents.FOneItem.Fsdate
	edate = oMainContents.FOneItem.Fedate
end if

if Not(sdate="" or isNull(sdate)) then
	sDt = left(sdate,10)
	sTm = Num2Str(hour(sdate),2,"0","R") &":"& Num2Str(minute(sdate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = date
	end if
	sTm = "00:00"
end if

if Not(edate="" or isNull(edate)) then
	eDt = left(edate,10)
	eTm = Num2Str(hour(edate),2,"0","R") &":"& Num2Str(minute(edate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = date
	end if
	eTm = "00:00"
end If

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript'>

function SaveMainContents(frm){
    if (frm.poscode.value.length<1){
        alert('구분을 먼저 선택 하세요.');
        frm.poscode.focus();
        return;
    }

    if (frm.image_order.value.length<1){
        alert('이미지 우선순위를 입력 하세요.');
        frm.image_order.focus();
        return;
    }

	<% if poscode = "999" then %>
	    if (frm.image_order.value=="선택"){
	        alert('이미지 우선순위를 입력 하세요.');
	        frm.image_order.focus();
	        return;
	    }
	<% end if %>

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

function ChangeGubun(comp){
    location.href = "?poscode=" + comp.value + "&gubun=<%=gubun%>";
    // nothing;
}

function ChangeGroup(comp){
    location.href = "?gubun=" + comp.value;
}

function PreViewContents(frm){
	var bannerimg = "<%=imgFingers%>/main/<%= oMainContents.FOneItem.fimagepath %>";
	var textimg = "<%=imgFingers%>/main/<%= oMainContents.FOneItem.fimagepath_etc %>";
	var div = frm.poscode.value;
	if(div==999)
	{
		var bgcolor = frm.leftimagecolor.value;
	}else{
		var bgcolor = 0;
	}
	
	var linkurl = frm.linkpath.value;
	if(div==""){
		alert("구분명을 선택해 주세요.");
	}else if(bannerimg==""){
		alert("이미지를 선택해 주세요.");
	}else if(linkurl==""){
		alert("링크값을 입력해 주세요.");
	}else{
		var popPreView = window.open('http://www.thefingers.co.kr/chtml/admin_banner_check.asp?bannerimg='+bannerimg+'&div='+div+'&textimg='+textimg+'&bgcolor='+bgcolor+'&linkurl='+linkurl,'popPreView');
		popPreView.focus();
	}	
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
		<% if oMainContents.FOneItem.Fidx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
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
		<% if oMainContents.FOneItem.Fidx<>"" then %>minDate: "<%=sDt%>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
		}
	});
});
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="center">

		</td>
		<td align="right">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="<%=UploadImgFingers%>/linkweb/sitemaster/image_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="ckUserId" value="<%=session("ssBctId")%>">
<!--<input type="hidden" name="ckUserId" value="<%=request.Cookies("partner")("userid")%>">-->
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">Idx :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	        <%= oMainContents.FOneItem.Fidx %>
	        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
	        <% else %>

	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">그룹명 :</td>
	    <td>

	        <% if oMainContents.FOneItem.Fidx<>"" then %>
				<%= oMainContents.FOneItem.Fgubun %>
	        <% else %>
        		<% call DrawGroupGubunCombo ("gubun", gubun, "onchange='ChangeGroup(this);'") %>
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">구분명 :</td>
	    <td>

	        <% if oMainContents.FOneItem.Fidx<>"" then %>
				<%= oMainContents.FOneItem.Fposname %> (<%= oMainContents.FOneItem.Fposcode %>)
				<input type="hidden" name="poscode" value="<%= oMainContents.FOneItem.Fposcode %>">
	        <% else %>
        		<% call DrawMainPosCodeCombo("poscode", poscode, "onChange='ChangeGubun(this);'",gubun) %>
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">이미지정렬우선순위 :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
					<select name="image_order">
						<option>선택</option>
						<% for ix = 1 to 50 %>
							<option value="<%=ix%>" <% if cint(oMainContents.FOneItem.fimage_order) = cint(ix) then response.write " selected"%>><%= ix %></option>
						<% next %>
					</select>
	        <% else %>
	            <% if poscode<>"" then %>
					<select name="image_order">
						<option>선택</option>
						<% for ix = 1 to 50 %>
							<option value="<%=ix%>"><%= ix %></option>
						<% next %>
					</select>
					실서버 적용시 숫자가 작을경우 우선노출
	            <% else %>
	            <font color="red">구분을 먼저 선택하세요</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">링크구분 :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	        <%= oMainContents.FOneItem.fimagetype %>
	        <% else %>
	            <% if poscode<>"" then %>
	            <%= oposcode.FOneItem.fimagetype %>
	            <% else %>
	            <font color="red">구분을 먼저 선택하세요</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">이미지 :</td>
	  <td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <br><img src="<%=imgFingers%>/main/<%= oMainContents.FOneItem.fimagepath %>" border="0">
	  <br><%=imgFingers%>/main/<%= oMainContents.FOneItem.fimagepath %>
	  <% end if %>
	  </td>
	</tr>

	<% if poscode="999" then %>
		<tr bgcolor="#FFFFFF">
		  <td width="150" align="center">이미지 컬러 코드 :</td>
		  <td>
		  <% if oMainContents.FOneItem.Fidx<>"" then %>
		  	 <input type="text" name="leftimagecolor" value="<%= oMainContents.FOneItem.fleftimagecolor %>" size="10" maxlength="10">
		  	 (#제외)
		  <% else %>
		  	 <input type="text" name="leftimagecolor" value="" size="10" maxlength="10">
		  	 (#제외)
		  <% end if %>
		  </td>
		</tr>

		<tr bgcolor="#FFFFFF">
		  <td width="150" align="center">TEXT 이미지(선택) :</td>
		  <td><input type="file" name="file2" value="" size="32" maxlength="32" class="file">
		  <% if oMainContents.FOneItem.Fidx<>"" then %>
			  <% if oMainContents.FOneItem.fimagepath_etc <> "" then %>
				  <br><img src="<%=imgFingers%>/main/<%= oMainContents.FOneItem.fimagepath_etc %>" border="0">
				  <br><%=imgFingers%>/main/<%= oMainContents.FOneItem.fimagepath_etc %>
				<% end if %>
		  <% end if %>
		  </td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" align="center">기간</td>
			<td colspan="2">
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<input type="text" id="sDt" name="sdate" size="10" value="<%=sDt%>" />
					<input type="hidden" name="sTm" size="8" value="<%=sTm%>" /> ~
					<input type="text" id="eDt" name="edate" size="10" value="<%=eDt%>" />
					<input type="hidden" name="eTm" size="8" value="<%=eTm%>" />
				<% else %>
					<input type="text" id="sDt" name="sdate" size="10" value="<%=stdt%>" />
					<input type="hidden" name="sTm" size="8" value="<%=sTm%>" /> ~
					<input type="text" id="eDt" name="edate" size="10" value="<%=eddt%>" />
					<input type="hidden" name="eTm" size="8" value="<%=eTm%>" />
				<% end if %>
			</td>
		</tr>
	<% end if %>

	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">사용할 이미지수 :</td>
	  <td>
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <input type="text" name="imagehcount" value="<%= oMainContents.FOneItem.fimagecount %>" size="2" maxlength="2">
	  <% else %>
	        <% if poscode<>"" then %>
	        <%= oposcode.FOneItem.fimagecount %>
	        <% else %>
	        <font color="red">구분을 먼저 선택하세요</font>
	        <% end if %>
	  <% end if %>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="150"  align="center">이미지Width :</td>
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
	  <td width="150" align="center">이미지Height :</td>
	  <td>
	  <% if oMainContents.FOneItem.Fidx<>""  then %>
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
	<%
	'// 포스코드 800이나 900 일경우 입력 받음
	if poscode = "800" or oMainContents.FOneItem.Fposcode = "800" or poscode = "900" or oMainContents.FOneItem.Fposcode = "900" then
	%>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">강좌코드 :</td>
	    <td>
	        <%
	        '// 수정
	        if oMainContents.FOneItem.Fidx<>"" then
	        %>
	        	<input type="text" name="relation_itemcode" value="<%= oMainContents.FOneItem.frelation_itemcode %>" >
	        <%
	        '//신규
	        else
	        %>
	            <% if poscode<>"" then %>
	            	<input type="text" name="relation_itemcode" value="" >
	            <% else %>
	            <font color="red">구분을 먼저 선택하세요</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
	<% end if %>
	<%
	'// 포스코드 510, 800, 801 일경우 입력 받음
	if poscode = "800" or oMainContents.FOneItem.Fposcode = "800" or poscode = "801" or oMainContents.FOneItem.Fposcode = "801" or poscode = "510" or oMainContents.FOneItem.Fposcode = "510" Or poscode = "908" or oMainContents.FOneItem.Fposcode = "908" Or poscode = "911" or oMainContents.FOneItem.Fposcode = "911" Or poscode = "916" or oMainContents.FOneItem.Fposcode = "916" Or poscode = "917" or oMainContents.FOneItem.Fposcode = "917" then
	%>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">제목 :</td>
	    <td>
	        <%
	        '// 수정
	        if oMainContents.FOneItem.Fidx<>"" then
	        %>
	        	<textarea name="relation_itemtitle" cols="60" rows="2"><%= oMainContents.FOneItem.frelation_itemtitle %></textarea>
	        <%
	        '//신규
	        else
	        %>
	            <% if poscode<>"" then %>
	            	<textarea name="relation_itemtitle" cols="60" rows="2"></textarea>
	            <% else %>
	            <font color="red">구분을 먼저 선택하세요</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
	<% end if %>
	<%
	'// 포스코드 917 일경우 입력 받음
	if poscode = "917" or oMainContents.FOneItem.Fposcode = "917" then
	%>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">강사코드 :</td>
	    <td>
	        <%
	        '// 수정
	        if oMainContents.FOneItem.Fidx<>"" then
	        %>
	        	<textarea name="relation_itemtitle2" cols="60" rows="2"><%= oMainContents.FOneItem.frelation_itemtitle2 %></textarea>
	        <%
	        '//신규
	        else
	        %>
	            <% if poscode<>"" then %>
	            	<textarea name="relation_itemtitle2" cols="60" rows="2"></textarea>
	            <% else %>
	            <font color="red">구분을 먼저 선택하세요</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
	<% end if %>
	<%
	'// 포스코드800 일경우 입력 받음
	if poscode = "800" or oMainContents.FOneItem.Fposcode = "800"  Or poscode = "908" or oMainContents.FOneItem.Fposcode = "908"  Or poscode = "917" or oMainContents.FOneItem.Fposcode = "917" Or poscode = "999" or oMainContents.FOneItem.Fposcode = "999" then
	%>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">설명 :</td>
	    <td>
	        <%
	        '// 수정
	        if oMainContents.FOneItem.Fidx<>"" then
	        %>
	        	<textarea name="relation_itemcontents" cols="60" rows="2"><%= oMainContents.FOneItem.frelation_itemcontents %></textarea>
	        <%
	        '//신규
	        else
	        %>
	            <% if poscode<>"" then %>
	            	<textarea name="relation_itemcontents" cols="60" rows="2"></textarea>
	            <% else %>
	            <font color="red">구분을 먼저 선택하세요</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
	<% end if %>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">링크값 :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	            <% if oMainContents.FOneItem.fimagetype="map" then %>
	            <textarea name="linkpath" cols="60" rows="6"><%= oMainContents.FOneItem.flinkpath %></textarea>
	            <% else %>
	            <input type="text" name="linkpath" value="<%= oMainContents.FOneItem.flinkpath %>" maxlength="256" size="60">
	            <% end if %>
	        <% else %>
	            <% if poscode<>"" then %>
	                <% if oposcode.FOneItem.fimagetype="map" then %>
	                    <textarea name="linkpath" cols="60" rows="6"><map name='Map<%=poscode%>'></map></textarea>
	                    <br>(이미지맵 변수값 변경 금지)
	                <% else %>
	                    <input type="text" name="linkpath" value="" maxlength="256" size="60">
	                    <br>(상대경로로 표시해 주세요  ex: /culturestation/culturestation_event.asp?evt_code=7)
	                <% end if %>
	            <% else %>
	            <font color="red">구분을 먼저 선택하세요</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">등록일 :</td>
	    <td>
	        <%= oMainContents.FOneItem.Fregdate %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">사용여부 :</td>
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
	    <td width="150" align="center">디자인 작업자 :</td>
	    <td>
	        <input type="text" name="designer" value="<%= oMainContents.FOneItem.fdesigner %>" maxlength="32" size="60">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td  align="center" colspan=2>
	    	<input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);" class="button">
			<% If idx <> 0 Then %>&nbsp;&nbsp;<input type="button" value=" 미리보기 " onClick="PreViewContents(frmcontents);" class="button"><% End If %>
	    </td>
	</tr>
</form>
</table>
<%
set oposcode = Nothing
set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->