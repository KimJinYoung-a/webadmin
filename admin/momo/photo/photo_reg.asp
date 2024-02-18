<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 감성포토
' Hieditor : 2009.10.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i 
dim photoid, photoword, mainimage, regdate, isusing , detailimage
dim ingimage , wordimage ,wordovimage
	photoid = requestcheckvar(request("photoid"),8)

'// 이벤트 리스트
set ocontents = new cphoto_list
	ocontents.frectphotoid = photoid
	
	'//수정일경우에만 쿼리
	if photoid <> "" then
	ocontents.fphoto_oneitem()
	end if
	
	if ocontents.ftotalcount > 0 then
		photoid = ocontents.FOneItem.fphotoid
		photoword = ocontents.FOneItem.fphotoword
		mainimage = ocontents.FOneItem.fmainimage
		regdate = ocontents.FOneItem.fregdate
		isusing = ocontents.FOneItem.fisusing		
		detailimage = ocontents.FOneItem.fdetailimage
		wordimage = ocontents.FOneItem.fwordimage		
		ingimage = ocontents.FOneItem.fingimage
		wordovimage = ocontents.FOneItem.fwordovimage		
	end if
%>

<script language="javascript">

document.domain = "10x10.co.kr";

function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){

	window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imgdel';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

//저장
function reg(){
	if (frm.photoword.value==''){
	alert('키워드를 입력해주세요');
	frm.photoword.focus();
	return;
	}
	if (frm.isusing.value==''){
	alert('사용여부를 선택해주세요');
	return;
	}
	
	frm.action='/admin/momo/photo/photo_process.asp';
	frm.mode.value='edit';	
	frm.submit();
}
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode" >
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>키워드ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= photoid %><input type="hidden" name="photoid" value="<%= photoid %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>키워드</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="photoword" value="<%= photoword %>" size=20>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>진행중이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('ingimgdiv','ingimg','ing','2000','800','true');"/>		
		<input type="hidden" name="ingimg" value="<%= ingimage %>">
		<div align="right" id="ingimgdiv"><% IF ingimage<>"" THEN %><img src="<%=webImgUrl%>/momo/photo/ing/<%= ingimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>메인이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('mainimgdiv','mainimg','main','2000','800','true');"/>		
		<input type="hidden" name="mainimg" value="<%= mainimage %>">
		<div align="right" id="mainimgdiv"><% IF mainimage<>"" THEN %><img src="<%=webImgUrl%>/momo/photo/main/<%= mainimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>리스트이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('wordimgdiv','wordimg','word','2000','250','true');"/>		
		<input type="hidden" name="wordimg" value="<%= wordimage %>">
		<div align="right" id="wordimgdiv"><% IF wordimage<>"" THEN %><img src="<%=webImgUrl%>/momo/photo/word/<%= wordimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>리스트오버이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('wordovimgdiv','wordovimg','wordov','2000','250','true');"/>		
		<input type="hidden" name="wordovimg" value="<%= wordovimage %>">
		<div align="right" id="wordovimgdiv"><% IF wordovimage<>"" THEN %><img src="<%=webImgUrl%>/momo/photo/wordov/<%= wordovimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>설명이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('detailimgdiv','detailimg','detail','2000','800','true');"/>		
		<input type="hidden" name="detailimg" value="<%= detailimage %>">
		<div align="right" id="detailimgdiv"><% IF detailimage<>"" THEN %><img src="<%=webImgUrl%>/momo/photo/detail/<%= detailimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>사용여부</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2><input type="button" onclick="reg();" value="저장"></td>
</tr>
</form>
</table>

<form name="imginputfrm" method="post" action="">
<input type="hidden" name="YearUse" value="2009">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
<%
	set ocontents = nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->