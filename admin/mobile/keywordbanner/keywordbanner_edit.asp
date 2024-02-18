<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : 모바일 keywordbanner
' History : 2013.12.16 한용민
'###############################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incsessionadmin.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/mobile/keywordbanner_cls.asp" -->

<%
dim idx, keywordtype, keyword, imagepath, linkpath, isusing, orderno, regdate
dim lastdate, regadminid, lastadminid, keywordtypename, YearUse, menupos, imgalt
Dim srcSDT , srcEDT
Dim sDt, sTm, eDt, eTm
Dim mainStartDate , mainEndDate

	srcSDT = request("sDt")
	srcEDT = request("eDt")

	YearUse = "2013"
	idx = request("idx")
	menupos = request("menupos")
	
dim okeyword, i
set okeyword = new ckeywordbanner
	okeyword.frectidx = idx
	
	if idx <> "" then
		okeyword.getkeywordbanner_one()
		
		if okeyword.ftotalcount > 0 then
			idx = okeyword.FOneItem.fidx
			keywordtype = okeyword.FOneItem.fkeywordtype
			keyword = okeyword.FOneItem.fkeyword
			imagepath = okeyword.FOneItem.fimagepath
			linkpath = okeyword.FOneItem.flinkpath
			isusing = okeyword.FOneItem.fisusing
			orderno = okeyword.FOneItem.forderno
			regdate = okeyword.FOneItem.fregdate
			lastdate = okeyword.FOneItem.flastdate
			regadminid = okeyword.FOneItem.fregadminid
			lastadminid = okeyword.FOneItem.flastadminid
			keywordtypename = okeyword.FOneItem.fkeywordtypename
			imgalt = okeyword.FOneItem.fimgalt
			mainStartDate = okeyword.FOneItem.fstartdate
			mainEndDate = okeyword.FOneItem.fenddate
		end if
	end if
	
if orderno="" then orderno=99
if isusing="" then isusing="Y"


if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = date
	end if
	sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = date
	end if
	eTm = "23:59:59"
end if

%>

<script language="javascript">

function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/mobile/keywordbanner/keywordbanner_img_input.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

function keywordbannerproc(){
	if (frm.keywordtype.value==''){
		alert('타입을 선택해 주세요.');
		frm.keywordtype.focus();
		return;
	}
	if (frm.isusing.value==''){
		alert('사용여부를 선택해 주세요.');
		frm.isusing.focus();
		return;
	}
	if (frm.orderno.value==''){
		alert('정렬순위를 입력해 주세요.');
		frm.orderno.focus();
		return;
	}
	if (!IsDouble(frm.orderno.value)){
		alert('정렬순위는 숫자만 가능합니다.');
		frm.orderno.focus();
		return;
	}	
	
	frm.submit();	
}
	
function OnOffkeywordbanner(keywordtype){
	var keywordtype1 = document.getElementById("keywordtype1");
	var keywordtype2 = document.getElementById("keywordtype2");
		
	if (keywordtype == '1'){
		keywordtype1.style.display="";
		keywordtype2.style.display="none";
	}else if (keywordtype == '2'){
		keywordtype1.style.display="none";
		keywordtype2.style.display="";
	}
}
	
</script>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
		function jsSubmit(){
		var frm = document.frm;
	
		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/onair/";
	}
	$(function(){
	//달력대화창 설정
	var arrDayMin = ["일","월","화","수","목","금","토"];
	var arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
	var sTime = document.frm.sTm.value;
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
    	showCurrentAtPos: 1,
      	showOn: "button",
      	<% if Idx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
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
      	<% if Idx<>"" then %>minDate: "<%=sDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
});

function putLinkText(key) {
	var frm = document.frm;
	switch(key) {
		case 'search':
			frm.linkpath.value='/search/search_item.asp?rect=검색어';
			break;
		case 'event':
			frm.linkpath.value='/event/eventmain.asp?eventid=이벤트번호';
			break;
		case 'itemid':
			frm.linkpath.value='/category/category_itemprd.asp?itemid=상품코드';
			break;
		case 'category':
			frm.linkpath.value='/category/category_list.asp?disp=카테고리';
			break;
		case 'brand':
			frm.linkpath.value='/street/street_brand.asp?makerid=브랜드아이디';
			break;
	}
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ KEYWORDBANNER 등록
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post" action="/admin/mobile/keywordbanner/keywordbanner_process.asp">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="keywordbanneredit">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="imagepath" value="<%=imagepath%>">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td width="120" align="center"><B>노출기간</B></td>
    <td bgcolor="#FFFFFF" colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width=120><b>타입</b><br></td>
	<td bgcolor="#FFFFFF">
		<% if idx<>"" then %>
			<%= keywordtypename %>
			<input type="hidden" name="keywordtype" value="<%=keywordtype%>">
		<% else %>
			<% drawSelectBoxkeywordtype "keywordtype", keywordtype , " onchange='OnOffkeywordbanner(this.value)'" %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" id='keywordtype2' style='display:<% if keywordtype<>2 then Response.write "none" %>'>
	<td align="center"><b>키워드</b><br></td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="keyword" value="<%= keyword %>" size="30" maxlength="30">
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" id='keywordtype1' style='display:<% if keywordtype<>1 then Response.write "none" %>'>
	<td align="center"><b>이미지</b></td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('keywordbanner','<%= imagepath %>','imagepath','spanban')" class="button">
		<div id="spanban" style="padding: 5 5 5 5">
			<% IF imagepath <> "" THEN %>
				<img src="<%=imagepath%>" border="0" width="259" height="360">
				<a href="javascript:jsDelImg('imagepath','spanban');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
		alt : <input type="text" name="imgalt" value="<%= imgalt %>" size="50">
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>링크경로</b><br></td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="linkpath" value="<%= linkpath %>" size=100 maxlength=100>
		<br/>
		<font color="#707070">
			- <span style="cursor:pointer" onClick="putLinkText('search')">검색결과 링크 : /search/search_item.asp?rect=<font color="darkred">검색어</font></span><br>
			- <span style="cursor:pointer" onClick="putLinkText('event')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
			- <span style="cursor:pointer" onClick="putLinkText('itemid')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
			- <span style="cursor:pointer" onClick="putLinkText('category')">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br>
			- <span style="cursor:pointer" onClick="putLinkText('brand')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>사용여부</b><br></td>
	<td bgcolor="#FFFFFF">
		<% drawSelectBoxisusingYN "isusing", isusing, "" %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>정렬순위</b><br></td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="orderno" value="<%= orderno %>" size=3 maxlength=3>
	</td>
</tr>
<% if lastadminid<>"" then %>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center"><b>최근수정</b><br></td>
		<td bgcolor="#FFFFFF">
			<%= lastdate %>
			<Br>(<%= lastadminid %>)
		</td>
	</tr>
<% end if %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" value="저장" onclick="keywordbannerproc();" class="button">
	</td>
</tr>
</form>	
<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="YearUse" value="<%= YearUse %>">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="maxFileheight" value="">	
	<input type="hidden" name="makeThumbYn" value="">
</form>	
</table>

<%
set okeyword = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
