<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.19 김진영 생성
'			2013.08.29 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/TENBYTENCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim olist, page, i, makerid, isusing, research
dim catecode, standardCateCode, mduserid, brandgubun
	catecode	= request("catecode")
	standardCateCode	= request("standardCateCode")
	mduserid	= request("mduserid")
	brandgubun	= request("brandgubun")	
	page	= request("page")
	makerid	= request("makerid")
	isusing	= request("selectisusing")
	menupos	= request("menupos")
	research	= request("research")	
	
If page = ""	Then page = 1
if research ="" and isusing="" then isusing = "Y"

SET olist = new cTENBYTEN
	olist.FCurrPage		= page
	olist.FPageSize		= 50
	olist.FMakerid		= makerid
	olist.FIsusing		= isusing
	olist.Frectbrandgubun		= brandgubun
	olist.Frectcatecode = catecode
	olist.FrectstandardCateCode = standardCateCode
	olist.Frectmduserid = mduserid	
	olist.sbTENBYTENlist
%>

<script language="javascript">

var ichk = 1;

function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.type=="checkbox")) {
			e.checked = blnChk ;
		}
	}
}

function jsIsusingChg(selv) {
    var frm, blnChk;
	frm = document.fitem;
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.isusing.length;i++){
			frm.isusing[i].value=selv;
		}
	}else{
		frm.isusing.value=selv;
	}
}

//순서,사용여부 저장
function jsSortIsusing() {
	var frm;
	var sValue, sortNo, isusing;
	var makerid;
	makerid = "<%=makerid%>";
	frm = document.fitem;
	sValue = "";
	sortNo = "";
	isusing = "";
	chkSel	= 0;

	if(makerid == ''){
		alert('순서&사용여부 지정은 브랜드를 검색하신후 사용가능합니다.');
		document.frm.makerid.focus();
		return;
	}

	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked) chkSel++;

			if(!IsDigit(frm.sortNo[i].value)){
				alert("순서지정은 숫자만 가능합니다.");
				frm.sortNo[i].focus();
				return;
			}
			if (frm.chkI[i].checked){
				if (sValue==""){
					sValue = frm.chkI[i].value;
				}else{
					sValue =sValue+","+frm.chkI[i].value;
				}
				// 정렬순서
				if (sortNo==""){
					sortNo = frm.sortNo[i].value;
				}else{
					sortNo =sortNo+","+frm.sortNo[i].value;
				}

				// 사용여부
				if (isusing==""){
					isusing = frm.isusing[i].value;
				}else{
					isusing =isusing+","+frm.isusing[i].value;
				}
			}
		}
	}else{
		if(frm.chkI.checked) chkSel++;
		if(frm.chkI.checked){
			sValue = frm.chkI.value;
			if(!IsDigit(frm.sortNo.value)){
				alert("순서지정은 숫자만 가능합니다.");
				frm.sortNo.focus();
				return;
			}
			sortNo 	=  frm.sortNo.value;
			isusing =  frm.isusing.value;
		}
	}
	if(chkSel<=0) {
		alert("선택한 이미지가 없습니다.");
		return;
	}
	document.frmSortIsusing.itemidarr.value = sValue;
	document.frmSortIsusing.sortarr.value = sortNo;
	document.frmSortIsusing.isusingarr.value = isusing;
	document.frmSortIsusing.submit();
}

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

//리얼적용
function AssignReal(upfrm, imagecount){
	var idx = "";
	var chkSel = 0;
	var makerid;
	makerid = "<%=makerid%>";

	if(makerid == ''){
		alert('리얼적용은 브랜드를 검색하신후 사용가능합니다.');
		document.frm.makerid.focus();
		return;
	}
	
	var chkI = document.getElementsByName("chkI")
	var isusing = document.getElementsByName("isusing")
	var orgisusing = document.getElementsByName("orgisusing")
	var sortNo = document.getElementsByName("sortNo")
	var orgsortNo = document.getElementsByName("orgsortNo")
	
	for (var i=0;i<chkI.length;i++){
		if (chkI[i].checked){
			if (isusing[i].value != orgisusing[i].value){
				alert('사용여부를 저장하시고 시도해주세요.');
				isusing[i].focus();
				return;
			}
			if (isusing[i].value=='N'){
				alert('사용여부가 N 인것이 있습니다.');
				isusing[i].focus();				
				return;
			}
			if (sortNo[i].value != orgsortNo[i].value){
				alert('순서를 저장하시고 시도해주세요.');
				sortNo[i].focus();
				return;
			}

			chkSel++;
			idx = idx + chkI[i].value + ",";
		}
	}

	if(chkSel<=0) {
		alert("선택상품이 없습니다.");
		return;
	}

	var AssignpopReal;
	AssignpopReal = window.open("<%=wwwUrl%>/chtml/street/tenbytenandmake.asp?idx=" + idx + "&imagecount=" + imagecount + "&makerid=" + makerid, "AssignpopReal","width=400,height=300,scrollbars=yes,resizable=yes");
	AssignpopReal.focus();
}

</script>

<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>TENBYTEN</b>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : 
		<%' TENBYTEN_ID_with_Name "makerid",makerid %>
		<% drawSelectBoxDesignerwithName "makerid",makerid %>
		&nbsp;&nbsp;
		* 브랜드구분 : <% drawSelectBoxbrandgubun "brandgubun",brandgubun , " onchange=""gosubmit('');""" %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 대표카테고리 : 
		기능<% SelectBoxBrandCategory "catecode", catecode %>
		전시<%= fnStandardDispCateSelectBox(1,"", "standardCateCode", standardCateCode, "")%>
		&nbsp;&nbsp;
		* 담당MD : <% drawSelectBoxCoWorker_OnOff "mduserid", mduserid, "on" %>
		&nbsp;&nbsp;
		* 사용유무 :
		<% drawSelectBoxUsingYN "selectisusing", isusing %>		
	</td>
</tr>
</table>
</form>
<form name="frmSortIsusing" method="post" action="TENBYTENSortIsusingProcess.asp" style="margin:0px;">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="sortarr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="makerid" value="<%=makerid%>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<!-- 검색 끝 -->

<br><b><font color="red" size=4>기획 변경예정 입니다. 등록하지 마세요. 2013 리뉴얼이후 차후 오픈</font></b>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% if olist.fresultcount >0 then %>
	    	<!--<a href="javascript:AssignReal(frm,'9');"><img src="/images/refreshcpage.gif" border="0">Real적용</a>-->
	    <% end if %>
	</td>
	<td align="right">
		<% if olist.fresultcount >0 then %>	
			<input class="button" type="button" id="btnEditSel" value="순서/사용여부 수정" onClick="jsSortIsusing();">
			&nbsp;&nbsp;
	    <% end if %>		
		<input type="button" value="신규등록" onclick="javascript:location.href='TENBYTEN_write.asp?menupos=<%=menupos%>';" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=olist.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= olist.FTotalPage %></b>		
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
	<td>브랜드ID</td>
	<td>이미지/동영상</td>
	<td>순서</td>
	<td>사용여부
		<select name="selisusing" onchange="jsIsusingChg(this.value)" class="select">
			<option value="N">N</option>
			<option value="Y">Y</option>
		</select>
	</td>
	<td>등록일</td>
	<td>비고</td>
</tr>
<% if olist.fresultcount >0 then %>
<% For i = 0 to olist.fresultcount -1 %>
<tr height="25" bgcolor="FFFFFF">
	<td align="center"><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= olist.FItemlist(i).FIdx %>"></td>
	<td align="center"><%=olist.FItemlist(i).FMakerid%></td>
	<td align="center">
		<% If olist.FItemlist(i).FImgurl <> "" AND olist.FItemlist(i).FFlag = "1" Then %>
			<img src="<%=uploadUrl%>/brandstreet/TENBYTEN/<%=olist.FItemlist(i).FImgurl%>" width="50" height="50" />
		<% ElseIf olist.FItemlist(i).FFlag = "2" Then %>
			동영상
		<% End If %>
	</td>
	<td align="center">
		<input type="hidden" value="<%=olist.FItemList(i).FSortNo%>" name="orgsortNo">
		<input type="text" size="2" maxlength="2" name="sortNo" value="<%=olist.FItemList(i).FSortNo%>" class="text">
	</td>
	<td align="center">
		현재 : <%=olist.FItemList(i).FIsusing%>&nbsp;&nbsp;&nbsp;/&nbsp;&nbsp;&nbsp;
		<input type="hidden" value="<%=olist.FItemList(i).FIsusing%>" name="orgisusing">
		변경 :
		<select name="isusing" class="select">
			<option value="N" <%=Chkiif(olist.FItemList(i).FIsusing = "N","selected","")%> >N</option>
			<option value="Y" <%=Chkiif(olist.FItemList(i).FIsusing = "Y","selected","")%> >Y</option>
		</select>
	</td>
	<td align="center"><%=left(olist.FItemlist(i).FRegdate,10)%></td>
	<td align="center">
		<input type="button" onclick="javascript:location.href='TENBYTEN_write.asp?idx=<%=olist.FItemlist(i).FIdx%>&menupos=<%=menupos%>';" value="수정" class="button">
	</td>
</tr>
<% Next %>

<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If olist.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= olist.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + olist.StartScrollPage to olist.StartScrollPage + olist.FScrollCount - 1 %>
			<% If (i > olist.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(olist.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If olist.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</form>
</table>

<%
SET olist = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->