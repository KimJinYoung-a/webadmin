<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : 핑거스 모바일 상단 메인 배너 등록 페이지
'	History		: 2016.07.29 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/academy/mobile/main_banner/academy_mobile_mainbannerCls.asp"-->

<%
Dim i, mode
Dim sDt, sTm, eDt, eTm
Dim sdate, edate, gubun, art_text, makerid
Dim srcSDT , srcEDT, stdt, eddt
Dim idx, isusing, sortnum, regdate, linkurl_etc, layerpopurl, SearchGubun, map, linknum
Dim sqlstr, sqlsearch, arrlist, resultcount
Dim cEvtCont
	idx			=	RequestCheckvar(request("idx"),10)
	map			=	RequestCheckvar(request("map"),16)
	mode			=	RequestCheckvar(request("mode"),16)
	srcSDT			=	RequestCheckvar(request("sDt"),10)
	srcEDT			=	RequestCheckvar(request("eDt"),10)
	sdate			=	RequestCheckvar(request("sdate"),10)
	edate			=	RequestCheckvar(request("edate"),10)
	gubun			=	RequestCheckvar(request("gubun"),10)
	makerid		=	RequestCheckvar(request("makerid"),32)
	isusing		=	RequestCheckvar(request("isusing"),1)
	regdate		=	RequestCheckvar(request("regdate"),10)
	sortnum		=	RequestCheckvar(request("sortnum"),10)
	linknum		=	RequestCheckvar(request("linknum"),32)
	art_text		=	request("art_text")
	linkurl_etc	=	request("linkurl_etc")
	layerpopurl	=	RequestCheckvar(request("layerpopurl"),32)
  	if art_text <> "" then
		if checkNotValidHTML(art_text) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if linkurl_etc <> "" then
		if checkNotValidHTML(linkurl_etc) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
dim opart, con_viewthumbimg
	set opart = new CAcademyMobileMainBanner
		opart.fnGetAcademyMobileMainBannerList

if idx = "" then 
	mode="NEW"
else
	mode="EDIT"
end if

if mode="EDIT" then
	if idx <> "" then
		sqlsearch = sqlsearch & " and idx="& idx &""
	end if
		
		sqlstr = "select top 1"
		sqlstr = sqlstr & " idx, linkurl_etc, sdate, edate, isusing, sortnum, gubun, con_viewthumbimg, linknum, art_text, makerid"
		sqlstr = sqlstr & " from db_academy.dbo.tbl_academy_mobile_mainbanner_list"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by idx desc"

		rsACADEMYget.Open sqlstr, dbACADEMYget, 1
		resultcount = rsACADEMYget.recordcount
		
	if not rsACADEMYget.EOF then
		'suserid = userid
		arrlist = rsACADEMYget.getrows()
	end if
	
	rsACADEMYget.close
	
	idx = arrlist(0,0)
	linkurl_etc = arrlist(1,0)
	sdate = arrlist(2,0)
	edate = arrlist(3,0)
	isusing = arrlist(4,0)
	sortnum = arrlist(5,0)
	gubun = arrlist(6,0)
	con_viewthumbimg = arrlist(7,0)
	linknum = arrlist(8,0)
	art_text = arrlist(9,0)
	makerid = arrlist(10,0)
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
	eTm = "23:59"
end If

IF sortnum = "" then sortnum = "99"
%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	
function frmedit(){
	if(frm.SearchGubun.value==""){
		alert("종류를 선택해 주세요");
		frm.SearchGubun.focus();
		return;
	}

	if(frm.con_viewthumbimg.value==""){
		alert("이미지를 등록해 주세요");
		return;
	}

	<% if gubun<>"5" then %>
		if(frm.linknum.value==""){
			<% if gubun="1" then %>
				alert("강좌 번호를 입력해 주세요.");
	
			<% elseif gubun="2" then %>
				alert("상품 번호를 입력해 주세요.");
	
			<% elseif gubun="3" then %>
				alert("매거진 번호를 입력해 주세요.");
	
			<% elseif gubun="4" then %>
				alert("강사/작가 ID를 입력해 주세요.");

			<% end if %>
			frm.linknum.focus();
			return;
		}
	<% else %>
		if(frm.linkurl_etc.value==""){
			alert("기타 링크를 입력해 주세요.");
			frm.linkurl_etc.focus();
			return;
		}
	<% end if %>

	if(frm.art_text.value==""){
		alert("알트값을 입력해 주세요");
		frm.art_text.focus();
		return;
	}

	if(frm.StartDate.value==""){
		alert("시작일을 선택해 주세요");
		frm.StartDate.focus();
		return;
	}

	if(frm.EndDate.value==""){
		alert("종료일을 선택해 주세요");
		frm.EndDate.focus();
		return;
	}
	
	if(frm.sortnum.value==""){
		alert("우선순위를 입력해 주세요");
		frm.sortnum.focus();
		return;
	}
	frm.submit();
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

function chghicprogbn(comp){
    var frm=comp.form;
	location.href="/academy/mobile/main_banner/academy_mobile_mainbanner_write.asp?idx=<%= idx %>&gubun="+comp;
}

//이미지 확대화면 새창으로 보여주기
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/itemmaster/colortrend_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}
//이미지 삭제
function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	eval("document.all."+sName).value = "";
	eval("document.all."+sSpan).style.display = "none";
	}
}
//이미지 등록
function jsSetImg(sImg, sName, sSpan){	
	document.domain ="10x10.co.kr";	
	var winImg;
	winImg = window.open('/academy/mobile/main_banner/academy_mobile_mainbanner_imagereg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}
//한글 입력 안되게
function onlyNumDecimalInput(){
	var code = window.event.keyCode; 
	
	if ((code >= 48 && code <= 57) || (code >= 96 && code <= 105) || code == 110 || code == 190 || code == 8 || code == 9 || code == 13 || code == 46){ 
		window.event.returnValue = true; 
		return; 
	} 
	window.event.returnValue = false; 
}

//링크 텍스트박스 강조(색상)
function clearFieldColor(field) {
  if (field.value == field.defaultValue) {
      field.style.backgroundColor = "#FFFFFF";
  }
}
function checkFieldColor(field) {
  if (!field.value) {
      field.style.backgroundColor = "#FFDDDD";
  }
} 
</script>

<img src="/images/icon_arrow_link.gif"> <b>아카데미 모바일 메인배너 등록</b>
<form name="frm" method="post" action="academy_mobile_mainbanner_proc.asp">
<input type = "hidden" name = "idx" value = "<%=idx %>">
<input type = "hidden" name = "mode" value = "<%=mode %>">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="con_viewthumbimg" value="<%= con_viewthumbimg %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if mode = "EDIT"  then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">번호</td>
		<td colspan="2"><%=idx%></td>
	</tr>
	<% end if %>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">종류</td>
		<td colspan="2">
			<select name="SearchGubun" onChange='chghicprogbn(this.value)'>
				<option value ="" style="color:blue">종 류</option>
				<option value="1" <% If "1" = cstr(gubun) Then %> selected <% End if %>>강좌</option>
				<option value="2" <% If "2" = cstr(gubun) Then %> selected <% End if %>>상품</option>
				<option value="3" <% If "3" = cstr(gubun) Then %> selected <% End if %>>매거진</option>
				<option value="4" <% If "4" = cstr(gubun) Then %> selected <% End if %>>강사/작가</option>
				<option value="5" <% If "5" = cstr(gubun) Then %> selected <% End if %>>기타</option>
			</select>
		<% if mode = "NEW" then %>
			<font color="red">※종류를 꼭 먼저 선택해 주세요!!</font>
		<% end if %>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnmainbannerimg" value="이미지등록" onClick="jsSetImg('<%= con_viewthumbimg %>','con_viewthumbimg','maindiv')" class="button">
			<div id="maindiv" style="padding: 5 5 5 5">
				<% IF con_viewthumbimg <> "" THEN %>			
					<img src="<%=con_viewthumbimg%>" border="0" width=300 height=300 onclick="jsImgView('<%=con_viewthumbimg %>');" alt="누르시면 확대 됩니다">
					<!--
					<a href="javascript:jsDelImg('main','maindiv');"><img src="/images/icon_delete2.gif" border="0"></a>
					-->
				<% END IF %>
			</div>
		</td>
	</tr>

<% If "1" = cstr(gubun) Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" style="background-color:#FFDDDD;" align="center">강좌 번호</td>
		<td colspan="2">
			<input type="text" name="linknum" value="<%=linknum%>" onBlur="checkFieldColor(this);" onFocus="clearFieldColor(this);"/>
		</td>
	</tr>
<% elseif "2" = cstr(gubun) Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" style="background-color:#FFDDDD;"  align="center">상품 번호</td>
		<td colspan="2">
			<input type="text" name="linknum" value="<%=linknum%>" onBlur="checkFieldColor(this);" onFocus="clearFieldColor(this);"/>
		</td>
	</tr>
<% elseif "3" = cstr(gubun) Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" style="background-color:#FFDDDD;" align="center">매거진 번호</td>
		<td colspan="2">
			<input type="text" name="linknum" value="<%=linknum%>" onBlur="checkFieldColor(this);" onFocus="clearFieldColor(this);"/>
		</td>
	</tr>
<% elseif "4" = cstr(gubun) Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" style="background-color:#FFDDDD;" align="center">강사/작가 ID</td>
		<td colspan="2">
			<input type="text" name="linknum" value="<%=linknum%>" onBlur="checkFieldColor(this);" onFocus="clearFieldColor(this);"/>
		</td>
	</tr>
<% elseif "5" = cstr(gubun) Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" style="background-color:#FFDDDD;" align="center">기타링크</td>
		<td colspan="2">
			<input type="text" name="linkurl_etc" style="width:100%; background-color:#FFDDDD;" value="<%= trim(linkurl_etc) %>" onBlur="checkFieldColor(this);" onFocus="clearFieldColor(this);"/>
		</td>
	</tr>
<% end if %>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">알트값(필수)</td>
		<td colspan="2">
			<input type="text" name="art_text" value="<%=art_text%>" onBlur="checkFieldColor(this);" onFocus="clearFieldColor(this);"/>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">기간</td>
		<td colspan="2">
			<% if mode = "NEW" then %>
				<input type="text" id="sDt" name="StartDate" size="10" value="<%=stdt%>" />
				<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
				<input type="text" id="eDt" name="EndDate" size="10" value="<%=eddt%>" />
				<input type="text" name="eTm" size="8" value="<%=eTm%>" />
			<% else %>
				<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
				<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
				<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
				<input type="text" name="eTm" size="8" value="<%=eTm%>" />
			<% end if %>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center"> 사용여부 </td>
		<td colspan="2">
			<input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; 
			<input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">우선순위</td>
		<td colspan="2"><input type="text" name="sortnum" size="10" value="<%=sortnum%>" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled"/></td>
	</tr>
	
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="3">
		<% if gubun <> "" then %>
			<% if mode = "EDIT" or mode = "NEW" then %>
				<input type="button" class="button" uname="editsave" value="저장" onclick="frmedit()" />	
			<% end if %>
				<input type="button" class="button" name="editclose" value="취소" onclick="self.close()" />
		<% end if %>
		</td>
	</tr>
</table>
</form>
<% set opart = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->