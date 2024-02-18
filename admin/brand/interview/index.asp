<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/interviewCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim idx, olist, page, i, title, makerid, isusing, research
dim catecode, standardCateCode, mduserid, brandgubun
	catecode	= request("catecode")
	standardCateCode	= request("standardCateCode")
	mduserid	= request("mduserid")
	brandgubun	= request("brandgubun")	
	page	= request("page")
	idx		= request("idx")
	title	= request("title")
	makerid	= request("makerid")
	isusing	= request("isusing")
	research	= request("research")
	menupos	= request("menupos")
	
If page = ""	Then page = 1
if research ="" and isusing="" then isusing = "Y"

SET olist = new cinterview
	olist.FCurrPage		= page
	olist.FPageSize		= 20
	olist.FrectMakerid		= makerid
	olist.Frecttitle		= title
	olist.frectisusing = isusing
	olist.Frectbrandgubun		= brandgubun
	olist.Frectcatecode = catecode
	olist.FrectstandardCateCode = standardCateCode
	olist.Frectmduserid = mduserid	
	olist.finterviewmain_list
%>

<script language="javascript">

var ichk;
ichk = 1;

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

// 이미지 클릭시 원본 크기로 팝업 보기
function doImgPop(img){
	img1= new Image();
	img1.src=(img);
	imgControll(img);
}

function imgControll(img){
	if((img1.width!=0)&&(img1.height!=0)){
		viewImage(img);
	}else{
		controller="imgControll('"+img+"')";
		intervalID=setTimeout(controller,20);
	}
}

function viewImage(img){
	W=img1.width;
	H=img1.height;
	O="width="+W+",height="+H+",scrollbars=yes";
	imgWin=window.open("","",O);
	imgWin.document.write("<html><head><title>:*:*:*: 이미지상세보기 :*:*:*:*:*:*:</title></head>");
	imgWin.document.write("<body topmargin=0 leftmargin=0>");
	imgWin.document.write("<img src="+img+" onclick='self.close()' style='cursor:pointer;' title ='클릭하시면 창이 닫힙니다.'>");
	imgWin.document.close();
}

function goView(idx, makerid){
	location.href = "interviewModify.asp?mode=U&idx="+idx+"&makerid="+makerid+"&menupos=<%=menupos%>";
}

//순서 저장
function jsSort() {
	var frm;
	var sValue, sortNo;
	frm = document.fitem;
	sValue = "";
	sortNo = "";
	chkSel	= 0;
	var makerid;
	makerid = "<%=makerid%>";
	
	if(makerid == ''){
		alert('사용여부 지정은 브랜드를 검색하신후 사용가능합니다.');
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
			sortNo =  frm.sortNo.value; 
		}
	}
	if(chkSel<=0) {
		alert("선택한 이미지가 없습니다.");
		return;
	}
	document.frmSortImgSize.itemidarr.value = sValue;
	document.frmSortImgSize.sortnoarr.value = sortNo;
	document.frmSortImgSize.mode.value = 'sortedit';
	document.frmSortImgSize.submit();
}

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

</script>

<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>INTERVIEW</b>

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
		<%' drawinterview_ID_with_Name "makerid",makerid , " onchange='gosubmit("""");'" %>
		<% drawSelectBoxDesignerwithName "makerid",makerid %>
		&nbsp;&nbsp;
		* 브랜드구분 : <% drawSelectBoxbrandgubun "brandgubun",brandgubun , " onchange=""gosubmit('');""" %>		
		&nbsp;&nbsp;		
		제목 : <input type="text" name="title" value="<%=title%>" size="40" maxlength="40" class="text">
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
		<% drawSelectBoxUsingYN "isusing", isusing %>		
	</td>
</tr>
</table>
</form>
<form name="frmSortImgSize" method="post" action="/admin/brand/interview/interviewSortNoProcess.asp" style="margin:0px;">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="sortnoarr" value="">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="masteridx" value="<%=idx%>">
	<input type="hidden" name="menupos" value="<%= menupos %>">	
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<!--<input class="button" type="button" id="btnEditSel" value="노출순서수정" onClick="jsSort();">-->
	</td>
	<td align="right">
		<input type="button" value="신규등록" class="button" onclick="javascript:location.href='/admin/brand/interview/interviewModify.asp?menupos=<%=menupos%>';">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortnoarr" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=olist.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= olist.FTotalPage %></b>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<!--<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>-->
	<td>번호</td>
	<td>브랜드</td>
	<td>시작일</td>
	<td>메인<Br>이미지</td>
	<td>제목</td>
	<td>사용<Br>여부</td>	
	<td>최근수정</td>
	<td>비고</td>
</tr>
<% if olist.fresultcount > 0 then %>
<% For i = 0 to olist.fresultcount -1 %>
<% if olist.FItemList(i).fisusing="Y" then %>
	<tr height="25" bgcolor="FFFFFF"  align="center">
<% else %>
	<tr height="25" bgcolor="f1f1f1"  align="center">
<% end if %>	
	<!--<td align="center"><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= olist.FItemList(i).FmainIdx %>"></td>-->
	<td align="center"><%= olist.FItemList(i).FmainIdx %></td>
	<td align="center"><%= olist.FItemList(i).FMakerid %></td>
	<td align="center"><%= left(olist.FItemList(i).Fstartdate,10) %></td>
	<td align="center">
		<img src="<%=olist.FItemList(i).fmainimg%>" width="50" height="50" title="클릭하시면 원본크기로 보실 수 있습니다." style="cursor: pointer;" onclick="doImgPop('<%=olist.FItemList(i).fmainimg%>')"/>
	</td>
	<td>
		<%= olist.FItemList(i).Ftitle %>
	</td>
	<td>
		<%=olist.FItemList(i).FIsusing%>
	</td>	
	<td>
		<%= olist.FItemList(i).Flastupdate %>
		<Br>(<%= olist.FItemList(i).Flastadminid %>)
	</td>
	<td>
		<input type="button" onClick="goView('<%=olist.FItemList(i).FmainIdx%>', '<%=olist.FItemList(i).FMakerid%>')" value="수정" class="button">
	</td>	
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If olist.HasPreScroll Then %>
			<span class="olist_link"><a href="javascript:gosubmit('<%= olist.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + olist.StartScrollPage to olist.StartScrollPage + olist.FScrollCount - 1 %>
			<% If (i > olist.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(olist.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="olist_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If olist.HasNextScroll Then %>
			<span class="olist_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
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