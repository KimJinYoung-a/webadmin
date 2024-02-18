<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : HITCHHIKER ADMIN(페이지관리->이슈영역)
'	History		: 2014.07.09 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/about_hitchhikerCls.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim i, mode
Dim sDt, sTm, eDt, eTm
Dim srcSDT , srcEDT, stdt, eddt, todaybanner
Dim sdate, edate, gubun, vol1, vol2
Dim idx, evt_title, isusing, sortnum, regdate, imghtmltext, SearchGubun, issueimg
Dim sqlstr, sqlsearch, arrlist, resultcount
Dim cEvtCont
	idx = request("idx")
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	sdate = request("sdate")
	edate = request("edate")
	mode = request("mode")
	gubun = request("gubun")
	isusing = request("isusing")
	regdate = request("regdate")
	evt_title = request("evt_title")
	sortnum = request("sortnum")
	issueimg = request("issueimg")
	
dim opart
	set opart = new CAbouthitchhiker
		opart.fnGetHitchhikerList
				
'만약 idx값이 없을경우(신규등록) NEW, 아닐경우(수정) EDIT	
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
		sqlstr = sqlstr & " idx, imghtmltext, hic_title, sdate, edate, isusing, sortnum, gubun, vol1, vol2 , issueimg"
		sqlstr = sqlstr & " from db_sitemaster.dbo.tbl_hitchhiker_list"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by idx desc"
		
		rsget.Open sqlstr, dbget, 1
		
		resultcount = rsget.recordcount
		
	if not rsget.EOF then
		'suserid = userid
		arrlist = rsget.getrows()
	end if
	
		rsget.close
		
		idx = arrlist(0,0)
		imghtmltext = arrlist(1,0)
		evt_title = arrlist(2,0)
		sdate = arrlist(3,0)
		edate = arrlist(4,0)
		isusing = arrlist(5,0)
		sortnum = arrlist(6,0)
		gubun = arrlist(7,0)
		vol1 = arrlist(8,0)
		vol2 = arrlist(9,0)
		issueimg = arrlist(10,0)
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
	
	<% If "2" = cstr(gubun) Then %>
		if(frm.vol1.value==""){
			alert("공모전 번호를 입력해 주세요");
			frm.vol1.focus();
			return;
		}
		
		if(frm.vol2.value==""){
			alert("공모전 번호를 입력해 주세요");
			frm.vol2.focus();
			return;
		}
	<% end if %>
	
	if(frm.evt_title.value==""){
		alert("타이틀을 입력해 주세요");
		frm.evt_title.focus();
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

function chghicprogbn(comp){
    var frm=comp.form;
	location.href="/admin/hitchhiker/issuearea/about_hitchhiker_write.asp?idx=<%= idx %>&gubun="+comp;
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

//이미지 확대화면 새창으로 보여주기
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/itemmaster/colortrend_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	eval("document.all."+sName).value = "";
	eval("document.all."+sSpan).style.display = "none";
	}
}

function jsSetImg(sImg, sName, sSpan){	
	document.domain ="10x10.co.kr";	
	var winImg;
	winImg = window.open('/admin/hitchhiker/issuearea/issue_imagereg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function onlyNumDecimalInput(){  //한글 입력 안되게
	var code = window.event.keyCode; 
	
	if ((code >= 48 && code <= 57) || (code >= 96 && code <= 105) || code == 110 || code == 190 || code == 8 || code == 9 || code == 13 || code == 46){ 
		window.event.returnValue = true; 
		return; 
	} 
	window.event.returnValue = false; 
}
</script>
<img src="/images/icon_arrow_link.gif"> <b>히치하이커 이슈영역 등록</b>
<form name="frm" method="post" action="about_hitchhiker_proc.asp">
<input type = "hidden" name = "idx" value = "<%=idx %>">
<input type = "hidden" name = "mode" value = "<%=mode %>">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="issueimg" value="<%= issueimg %>">

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
				<option value ="" style="color:blue">구 분</option>
				<option value="1" <% If "1" = cstr(gubun) Then%> selected <%End if%>>발간</option>
				<option value="2" <% If "2" = cstr(gubun) Then%> selected <%End if%>>에디터모집</option>
				<option value="3" <% If "3" = cstr(gubun) Then%> selected <%End if%>>기타</option>
			</select>
		<% if mode = "NEW" then %>
			<font color="red">※종류를 꼭 먼저 선택해 주세요!!</font>
		<% end if %>
		</td>
	</tr>
	
	<% If "2" = cstr(gubun) Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">ESSAY 공모전No.</td>
		<td colspan="2"><input type="text" name="vol1" size="10" value="<%= vol1 %>" maxlength="7" />
		<font color="red">[ON]이벤트관리>>공모전리스트에 등록된 공모전No.(예:con45)</font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">PHOTO STICKER 공모전No.</td>
		<td colspan="2"><input type="text" name="vol2" size="10" value="<%= vol2 %>" maxlength="7" />
		<font color="red">[ON]이벤트관리>>공모전리스트에 등록된 공모전No.(예:con46)</font></td>
	</tr>
	<% end if %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">타이틀</td>
		<td colspan="2">
			<% if mode = "NEW" then %>			
				<input type="text" name="evt_title" size="50" value=""/>
			<% else %>
				<input type="text" name="evt_title" size="50" value="<%=evt_title%>"/>
			<% end if %>
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
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnhicthumbimg" value="이미지등록" onClick="jsSetImg('<%= issueimg %>','issueimg','issueimgdiv')" class="button">
			<div id="issueimgdiv" style="padding: 5 5 5 5">
				<% IF issueimg <> "" THEN %>			
					<img src="<%=issueimg%>" border="0" width=100 height=100 onclick="jsImgView('<%=issueimg %>');" alt="누르시면 확대 됩니다">
					<a href="javascript:jsDelImg('issueimg','issueimgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
				<% END IF %>
			</div>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">HTML소스</td>
		<td colspan="2">
		<font color="red">
			※ map id="hitchhikerissue" name="hitchhikerissue" 절대 변경 하지 마세요.<br>
			※ 에디터 모집 소스 입력시 공모전No(예:con48)를 사용할 공모전No로 수정하세요.<br>
			※ 에세이, 포토스티커 모집 공모전No는 다릅니다.[예:에세이=('1','con47') , 포토=('2','con48')]
		</font>
		<textarea name="imghtmltext" class="textarea" style="width:100%; height:150px;"><%= trim(imghtmltext)%></textarea>
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
				<% if mode = "EDIT" or mode = "NEW" then %>
					<input type="button" class="button" name="editsave" value="저장" onclick="frmedit()" />	
				<% end if %>
					<input type="button" class="button" name="editclose" value="취소" onclick="self.close()" />
		</td>
	</tr>
</table>
</form>
<%
set opart = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->