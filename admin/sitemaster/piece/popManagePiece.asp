<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : piece 등록 팝업
' Hieditor : 2017.08.11 유태욱 생성
' Hieditor : 2017.09.05 원승현 추가/수정
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/piece/piececls.asp"-->
<%
Dim i, mode
Dim gubun	'1 : 조각, 2 :파이, 3 : 베스트키워드, 4: 배너, 5:회원조각
Dim startdate, enddate
Dim idx, con_title, isusing, sortnum, regdate, con_detail, con_movieurl
dim shorttext, listtext, tagtext, noticeYN, listtitle
Dim cEvtCont, loginUserId, oPieceView, eDt, sDt
Dim tt, tmptagtext, oRelationItemList, rr, arritemid, pieceidx
dim opart, con_viewthumbimg, appendNumberPrv, maxlengthshorttext, maxlengthlisttitle, bannergubun, etclink
Dim admintext , state , nickname , lastadminid , lastupdate , occupation , adminid , page
Dim starttime

Dim SearchDeal , SearchOpen , SearchState

SearchDeal = requestCheckvar(request("SearchDeal"), 1) '// 검색 parameter
SearchOpen = requestCheckvar(request("SearchOpen"), 1) '// 검색 parameter
SearchState = requestCheckvar(request("SearchState"), 1) '// 검색 parameter

idx = requestCheckvar(request("idx"), 50)
gubun = requestCheckvar(request("gubun"), 50)
bannergubun = requestCheckvar(request("bannergubun"), 50)

page = requestCheckvar(request("page"), 10)

loginUserId = session("ssBctId")
if gubun = "" then gubun = "1"
appendNumberPrv = 0

'만약 idx값이 없을경우(신규등록) NEW, 아닐경우(수정) EDIT
if Trim(idx) = "" then
	mode="NEW"
else
	mode="EDIT"
end If


If Trim(mode)="EDIT" Then
	'// Piece View 데이터를 가져온다.
	set oPieceView = new Cgetpiece
		oPieceView.FRectIdx = idx
		oPieceView.getPieceview()
		gubun = oPieceView.FOnePiece.Fgubun
		If Trim(bannergubun)="" Then
			bannergubun = oPieceView.FOnePiece.Fbannergubun
		End If
		con_viewthumbimg = oPieceView.FOnePiece.Flistimg
		listtitle = oPieceView.FOnePiece.Flisttitle
		isusing = oPieceView.FOnePiece.Fisusing
		sortnum = oPieceView.FOnePiece.Ffidx
		regdate = oPieceView.FOnePiece.Fregdate
		shorttext = oPieceView.FOnePiece.Fshorttext
		listtext = oPieceView.FOnePiece.Flisttext
		noticeYN = oPieceView.FOnePiece.Fnoticeyn
		TagText = oPieceView.FOnePiece.Ftagtext
		arritemid = oPieceView.FOnePiece.FItemid
		pieceidx = oPieceView.FOnePiece.Fpieceidx
		etclink = oPieceView.FOnePiece.Fetclink
		startdate = oPieceView.FOnePiece.Fstartdate
		enddate = oPieceView.FOnePiece.Fenddate
		admintext = oPieceView.FOnePiece.Fadmintext
		state = oPieceView.FOnePiece.FState
		nickname = oPieceView.FOnePiece.FNickname
		lastadminid = oPieceView.FOnePiece.Flastadminid
		lastupdate = oPieceView.FOnePiece.Flastupdate
		occupation = oPieceView.FOnePiece.Foccupation
		adminid = oPieceView.FOnePiece.Fadminid
End If

if Not(startdate="" or isNull(startdate)) Then
	starttime = Num2Str(hour(startdate),2,"0","R") &":"& Num2Str(minute(startdate),2,"0","R") &":"& Num2Str(second(startdate),2,"0","R")
else
	starttime = "00:00:00"
end if

If Trim(noticeYN="") Then
	noticeYN = "N"
End If

If bannergubun = "" Then bannergubun = "1"

'// 최대 글자 제한
Select Case Trim(gubun)
	Case "1"
		maxlengthshorttext = 21
	Case "2"
		maxlengthlisttitle = 20
		maxlengthshorttext = 20
	Case "3"
		maxlengthlisttitle = 20
	Case "4"
		maxlengthlisttitle = 40
End Select
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
body {background-color:#fff;}
</style>
</head>
<body>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/tag-it.min.js"></script>
<script type='text/javascript'>
document.domain = "10x10.co.kr";

function frmedit(){
	var frm  = document.frm;
	var gubun_value = "";
	for ( var i = 0; frm.gubun.length ; i++ ){
		if (frm.gubun[i].checked == true){
			gubun_value = frm.gubun[i].value;
			break;
		}
	}

	if(gubun_value == "")
	{
		alert('구분값을 선택해 주세요');
		return;
	}
	var tmpgubun = gubun_value;

	<% if trim(gubun)="4" then %>
		var tmpbannergubun = frm.bannergubun.value;
	<% end if %>

	if(tmpgubun == "1")
	{ //구분값이 조각 일때 체크해야 될 것들
		$("#tagtext").val("");
		if($("#tagtext").val().length < 1 ){
			$("input[name=tags]").each(function(idx){
				// 해당 체크박스의 Value 가져오기
				var value = $("#tagtext").val();
				var eqValue = $("input[name=tags]:eq(" + idx + ")").val() ;
					if($("#tagtext").val().length < 1 ){
					$("#tagtext").val(eqValue);
					console.log(value + "," + eqValue) ;
				}else{
					$("#tagtext").val(value + "," + eqValue);
				}
			});
		}

//		if(frm.con_viewthumbimg.value=="")
//		{
//			alert('이미지를 추가해 주세요');
//			frm.con_viewthumbimg.focus();
//			return;
//		}

		if(frm.listtext.value=="")
		{
			alert('내용을 입력해 주세요');
			frm.listtext.focus();
			return;
		}

		if ($("#tagtext").val()=="")
		{
			alert("태그를 입력해주세요");
			return;
		}

		if(frm.itemid.value=="")
		{
			alert('연관상품을 등록해주세요');
			return;
		} else {
			// 2018-01-23, skyer9
			var itemidArr = frm.itemid.value.split(",");
			for (var i = 0; i < itemidArr.length; i++) {
				if (itemidArr[i].length >= 10) {
					alert("===================================\n\n에러 : 시스템팀 문의\n\n상품코드 : " + itemidArr[i] + "\n\n===================================");
					return;
				}
			}
		}

		if(frm.startdate.value=="")
		{
			alert('시작일을 입력해주세요');
			frm.startdate.focus();
			return;
		}
	}
	else if (tmpgubun == "2")
	{ //구분값이 파이 일때 체크해야 될 것들
		if(frm.listtitle.value=="")
		{
			alert('제목을 입력해주세요');
			frm.listtitle.focus();
			return;
		}

		if(frm.listtext.value=="")
		{
			alert('내용을 입력해 주세요');
			frm.listtext.focus();
			return;
		}

		if(frm.pieceidx.value=="")
		{
			alert('연관조각을 입력해주세요');
			frm.pieceidx.focus();
			return;
		}

		if(frm.startdate.value=="")
		{
			alert('시작일을 입력해주세요');
			frm.startdate.focus();
			return;
		}
	}
	else if (tmpgubun == "3")
	{ //구분값이 베스트키워드 일때 체크해야 될 것들
		if(frm.listtitle.value=="")
		{
			alert('제목을 입력해주세요');
			frm.listtitle.focus();
			return;
		}

		if(frm.startdate.value=="")
		{
			alert('시작일을 입력해주세요');
			frm.startdate.focus();
			return;
		}
	}
	else if (tmpgubun == "4")
	{ //구분값이 배너 일때 체크해야 될 것들
		if(frm.listtitle.value=="")
		{
			alert('제목을 입력해주세요');
			frm.listtitle.focus();
			return;
		}

		if(frm.etclink.value=="")
		{
			alert('링크를 입력해주세요');
			frm.etclink.focus();
			return;
		}

		if(frm.startdate.value=="")
		{
			alert('시작일을 입력해주세요');
			frm.startdate.focus();
			return;
		}

		if(frm.enddate.value=="")
		{
			alert('종료일을 입력해주세요');
			frm.enddate.focus();
			return;
		}

		if (tmpbannergubun=="2")
		{
			if(frm.con_viewthumbimg.value=="")
			{
				alert('이미지를 추가해 주세요');
				frm.con_viewthumbimg.focus();
				return;
			}
		}
	}

//	if(frm.startdate.value==""){
//		alert('시작일을 입력해 주세요');
//		frm.startdate.focus();
//		return;
//	}

//	if(frm.enddate.value==""){
//		alert('종료일을 입력해 주세요');
//		frm.enddate.focus();
//		return;
//	}

	var tmpisusing = "";
	for(var i = 0;  i < frm.isusing.length; i++)
	{
		if(frm.isusing[i].checked==true){
		tmpisusing = frm.isusing[i].value;
		}
	}

	if(tmpisusing == "")
	{
		alert('사용여부를 선택하세요');
		return;
	}

	if(frm.isusing.value=="")
	{
		alert('노출 여부를 선택해 주세요.');
		frm.isusing.focus();
		return;
	}

	frm.submit();
}

// 오프닝 체크
function openChk()
{
	if ($("#noticeYN").prop("checked"))
	{
		if(confirm("새로운 오프닝 지정 시 이전에 지정되었던 오프닝은\n일반 조각으로 전환됩니다.")){
			$("#noticeYN").val('Y');
		}
		else
		{
			$("#noticeYN").attr('checked', false) ;
			$("#noticeYN").val('N');
		}
	}
	else
	{
		$("#noticeYN").val('N');
	}
}

$(function()
{
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
		<% if idx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
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
		<% if idx<>"" then %>minDate: "<%=sDt%>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
		}
	});
});

//구분값 변경
function chghicprogbn(comp)
{
    var frm=comp.form;
	location.href="/admin/sitemaster/piece/popManagePiece.asp?idx=<%= idx %>&gubun="+comp;
}

//배너구분값 변경
function chbannerghicprogbn(comp)
{
    var frm=comp.form;
	location.href="/admin/sitemaster/piece/popManagePiece.asp?idx=<%= idx %>&gubun=<%=gubun%>&bannergubun="+comp;
}

//이미지 확대화면 새창으로 보여주기
function jsImgView(sImgUrl)
{
	var wImgView;
	wImgView = window.open('/admin/itemmaster/colortrend_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

function jsDelImg(sName, sSpan)
{
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	eval("document.all."+sName).value = "";
	eval("document.all."+sSpan).style.display = "none";
	}
}

function jsSetImg(sImg, sName, sSpan)
{
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/sitemaster/piece/piece_imagereg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}


function checkLength(objname, maxlength)
{
	var objstr = objname.value;
	var objstrlen = objstr.length

	var maxlen = maxlength;
	var i = 0;
	var bytesize = 0;
	var strlen = 0;
	var onechar = "";
	var objstr2 = "";

	for (i = 0; i < objstrlen; i++)
	{
		onechar = objstr.charAt(i);

		if (escape(onechar).length > 4)
		{
			bytesize += 2;
		}
		else
		{
			bytesize++;
		}

		if (bytesize <= maxlen)
		{
			strlen = i + 1;
		}
	}

	if (bytesize > maxlen)
	{
		alert("허용된 문자열을 초과하였습니다.\n한글 기준 최대 "+maxlength/2+"자 까지 작성할 수 있습니다.");
		objstr2 = objstr.substr(0, strlen);
		objname.value = objstr2;
	}
	objname.focus();
}
//링크값선택
function showDrop(g){
	$("#selectLink"+g+"").show();
}

function linkcopy(g){
	var val = $("#etclink").val();
	$("#selectLink"+g+"").css("display","none");
}
function populateTextBox(v,g){
	var val = v;
	$("#etclink").val(val);
	$("#selectLink"+g+"").css("display","none");
}

</script>
<%' 팝업 사이즈 : 750*800 %>
<form name="frm" method="post" action="piece_contents_proc.asp">
<input type="hidden" name="mode" value="<%=mode %>">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="idx" value="<%=idx %>">
<input type="hidden" name="con_viewthumbimg" value="<%= con_viewthumbimg %>">
<input type="hidden" name="usertype" value=1>
<input type="hidden" name="adminid" value="<%=loginUserId%>">
<input type="hidden" name="page" value="<%=page%>">

<input type="hidden" name="SearchDeal" value="<%=SearchDeal%>">
<input type="hidden" name="SearchOpen" value="<%=SearchOpen%>">
<input type="hidden" name="SearchState" value="<%=SearchState%>">

	<div class="popWinV17">
		<h1>
			<% If Trim(mode)="EDIT" Then %>
				수정
			<% Else %>
				등록
			<% End If %>
		</h1>
		<div class="popContainerV17 pad30">
			<table class="tbType1 writeTb">
				<colgroup>
					<col width="18%" /><col width="" />
				</colgroup>
				<tbody>
				<% If Trim(mode)="EDIT" Then %>
				<tr>
					<th><div>번호(idx) <strong class="cRd1"></strong></div></th>
					<td><%=idx%></td>
				</tr>
				<% End If %>
				<tr>
					<th><div>구분 <strong class="cRd1">*</strong></div></th>
					<td>
						<label class="rMar20"><input type="radio" class="formRadio" name="gubun" <% If Trim(mode)="EDIT" Then %><% if gubun="1" then %>checked="checked"<% Else %>disabled<% end if %><% Else %><% if gubun="1" then %>checked="checked"<% End If %><% End If %> value="1" onChange="chghicprogbn(this.value);" /> 조각</label>
						<label class="rMar20"><input type="radio" class="formRadio" name="gubun" <% If Trim(mode)="EDIT" Then %><% if gubun="2" then %>checked="checked"<% Else %>disabled<% end if %><% Else %><% if gubun="2" then %>checked="checked"<% End If %><% End If %> value="2" onChange="chghicprogbn(this.value);" /> 파이</label>
						<label class="rMar20"><input type="radio" class="formRadio" name="gubun" <% If Trim(mode)="EDIT" Then %><% if gubun="3" then %>checked="checked"<% Else %>disabled<% end if %><% Else %><% if gubun="3" then %>checked="checked"<% End If %><% End If %> value="3" onChange="chghicprogbn(this.value);" /> 베스트키워드</label>
						<label class="rMar20"><input type="radio" class="formRadio" name="gubun" <% If Trim(mode)="EDIT" Then %><% if gubun="4" then %>checked="checked"<% Else %>disabled<% end if %><% Else %><% if gubun="4" then %>checked="checked"<% End If %><% End If %> value="4" onChange="chghicprogbn(this.value);" /> 배너</label>
					</td>
				</tr>
				<% If Trim(gubun)="4" Then %>
					<tr>
						<th><div>타입 <strong class="cRd1">*</strong></div></th>
						<td>
							<label class="rMar20"><input type="radio" class="formRadio" name="bannergubun" <% if bannergubun="1" then %>checked="checked"<% End If %> value="1" onChange="chbannerghicprogbn(this.value);" /> 텍스트</label>
							<label class="rMar20"><input type="radio" class="formRadio" name="bannergubun" <% if bannergubun="2" then %>checked="checked"<% End If %> value="2" onChange="chbannerghicprogbn(this.value);" /> 이미지</label>
						</td>
					</tr>
				<% End If %>
				<% If Trim(gubun)="2" Or Trim(gubun)="3" Or Trim(gubun)="4" Then %>
					<tr>
						<th><div><%=chkiif(Trim(gubun)="2","여는말","제목")%> <strong class="cRd1">*</strong></div></th>
						<td>
							<p><input type="text" name="listtitle" id="listtitle" class="formTxt" style="width:100%;" value="<%= listtitle %>" onKeyup="checkLength(this, <%=maxlengthlisttitle*2%>);" /></p>
							<p class="tPad05 fs11 cGy1">- 한글 기준 최대 <%=maxlengthlisttitle%>자까지 입력 가능합니다.</p>
						</td>
					</tr>
				<% End If %>
				<% If Trim(gubun)="4" Then %>
					<tr>
						<th><div>링크 <strong class="cRd1">*</strong></div></th>
						<td>
							<div class="selectLink">
								<p><input type="text" name="etclink" id="etclink" class="formTxt" style="width:100%;" value="<%= etclink %>" onclick="showDrop('m');" onkeyup="linkcopy('m');" maxlength="200" /></p>
								<ul style="display:none;" id="selectLinkm">
									<li onclick="populateTextBox('<%=CHKIIF(etclink="","",etclink)%>','m');">선택안함</li>
									<li onclick="populateTextBox('/category/category_itemPrd.asp?itemid=상품코드','m');">/category/category_itemPrd.asp?itemid=상품코드</li>
									<li onclick="populateTextBox('/category/category_list.asp?disp=카테고리','m');">/category/category_list.asp?disp=카테고리</li>
									<li onclick="populateTextBox('/street/street_brand.asp?makerid=브랜드아이디','m');">/street/street_brand.asp?makerid=브랜드아이디</li>
									<li onclick="populateTextBox('/event/eventmain.asp?eventid=이벤트코드','m');">/event/eventmain.asp?eventid=이벤트코드</li>
									<li onclick="populateTextBox('/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드','m');">/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드</li>
									<li onclick="populateTextBox('/gift/gifttalk/','m');">기프트</li>
									<li onclick="populateTextBox('/wish/index.asp','m');">위시</li>
								</ul>
							</div>
						</td>
					</tr>
					<tr>
						<th><div>기간 <strong class="cRd1">*</strong></div></th>
						<td>
								시작일 : <input type="text" id="sDt" name="startdate" size="10" value="<%=startdate%>" />&emsp; &emsp;
								종료일 : <input type="text" id="eDt" name="enddate" size="10" value="<%=enddate%>" />
						</td>
					</tr>
				<% End If %>

				<% If Trim(gubun)="1" Or Trim(gubun)="2" Then %>
					<tr>
						<th><div><%=chkiif(Trim(gubun)="2","제목","여는말")%> <% If Trim(gubun) = "2" Then %><strong class="cRd1">*</strong><% End If %></div></th>
						<td>
<!-- 							<p><input type="text" name="shorttext" id="shorttext" class="formTxt" style="width:100%;" value="<%= shorttext %>" onKeyup="checkLength(this, <%=maxlengthshorttext*2%>);" /></p> -->
							<p><textarea name="shorttext" id="shorttext" class="formTxt" style="width:100%;height:40px;" onKeyup="checkLength(this, <%=maxlengthshorttext*2%>);"><%= shorttext %></textarea></p>
							<p class="tPad05 fs11 cGy1">- 한글 기준 최대 <%=maxlengthshorttext%>자까지 입력 가능합니다.</p>
						</td>
					</tr>
				<% End If %>
				<% If Trim(gubun)="1" Or Trim(gubun)="2" Or Trim(gubun)="3" Then %>
					<tr>
						<th><div><% If gubun="1" Then %>이미지<% End If %><% If gubun="2" Then %>배경<% End If %><% If gubun="3" Then %>배경<% End If %></div></th>
						<td>
							<div class="inTbSet">
								<div>
									<p><input type="file" name="btnhicthumbimg" onClick="jsSetImg('<%= con_viewthumbimg %>','con_viewthumbimg','con_viewthumbimgdiv')" class="formFile" style="width:90%;" /></p>
									<% If gubun="1" Then %>
										<p class="tPad05 fs11 cGy1">- 이미지는 두가지 사이즈만 첨부 가능합니다. (<b>1,000x1,000</b>px / <b>1,000x1,266</b>px)</p>
									<% End If %>
									<% If gubun="2" Then %>
										<p class="tPad05 fs11 cGy1">- 이미지 사이즈 (<b>1,000x1,400</b>px)</p>
									<% End If %>
									<% If gubun="3" Then %>
										<p class="tPad05 fs11 cGy1">- 이미지 사이즈 (<b>1,000x1,000</b>px)</p>
									<% End If %>
								</div>
								<div style="width:120px;" id="con_viewthumbimgdiv">
									<% IF con_viewthumbimg <> "" THEN %>
										<p class="registImg">
											<button type="button" onclick="jsDelImg('con_viewthumbimg','con_viewthumbimgdiv');">X</button>
											<img src="<%= con_viewthumbimg %>" onclick="jsImgView('<%= con_viewthumbimg %>');" alt="누르시면 확대 됩니다" style="width:120px;" />
										</p>
									<% end if %>
								</div>
							</div>
						</td>
					</tr>
				<% End If %>
				<% If Trim(gubun)="4" Then %>
					<% If Trim(bannergubun)="2" Then %>
						<tr>
							<th><div>이미지</div></th>
							<td>
								<div class="inTbSet">
									<div>
										<p><input type="file" name="btnhicthumbimg" onClick="jsSetImg('<%= con_viewthumbimg %>','con_viewthumbimg','con_viewthumbimgdiv')" class="formFile" style="width:90%;" /></p>
										<p class="tPad05 fs11 cGy1">- 이미지 권장 가로 사이즈 (<b>1,000x198</b>px)</p>
									</div>
									<div style="width:120px;" id="con_viewthumbimgdiv">
										<% IF con_viewthumbimg <> "" THEN %>
											<p class="registImg">
												<button type="button" onclick="jsDelImg('con_viewthumbimg','con_viewthumbimgdiv');">X</button>
												<img src="<%= con_viewthumbimg %>" onclick="jsImgView('<%= con_viewthumbimg %>');" alt="누르시면 확대 됩니다" style="width:120px;" />
											</p>
										<% end if %>
									</div>
								</div>
							</td>
						</tr>
					<% End If %>
				<% End If %>
				<% If Trim(gubun)="1" Or Trim(gubun)="2" Then %>
					<tr>
						<th><div>내용 <strong class="cRd1">*</strong></div></th>
						<td>
							<p><textarea class="formTxtA" name="listtext" style="width:100%; height:80px;" onKeyup="checkLength(this, 1600);"><%= listtext %></textarea></p>
							<p class="tPad05 fs11 cGy1">- 한글기준 최대 800자까지 입력 가능합니다.</p>
						</td>
					</tr>
				<% End If %>
				<% If gubun="1" Then %>
					<tr>
						<input type="hidden" name="tagtext" id="tagtext" value="" >
						<th><div>태그 <strong class="cRd1">*</strong></div></th>
						<td>
							<ul id="singleFieldTags">
								<%
									If Trim(tagtext) <> "" Then
										tmptagtext = Split(tagtext, ",")
										For tt = 0 To UBound(tmptagtext)
								%>
											<li><%=tmptagtext(tt)%></li>
								<%
										Next
									End If
								%>
							</ul>
							<p class="tPad05 fs11 cGy1">- 태그는 전체 100자까지 입력가능하며, 태그에는 공백을 사용할 수 없습니다.</p>
						</td>
					</tr>
					<tr>
						<input type="hidden" name="itemid" id="itemid" value="<%=arritemid%>">
						<th><div>연관상품 <strong class="cRd1">*</strong></div></th>
						<td>
							<div class="pdtLinkWrap">
								<div class="pdtAdd"><a href="" class="btn-append">상품추가</a></div>
								<div class="swiper-container">
									<div class="swiper-wrapper">
										<% If Trim(mode)="EDIT" Then %>
											<%
												set oRelationItemList = new Cgetpiece
												oRelationItemList.FRectIdx = idx
												oRelationItemList.GetRelationItemList()
												appendNumberPrv = oRelationItemList.FResultCount
											%>
											<% If oRelationItemList.FResultCount > 0 Then %>
												<% For rr=0 To oRelationItemList.Fresultcount-1 %>
													<div class="swiper-slide" id="itemidimgdiv<%=rr+1%>">
														<button type="button" onclick="fndelitemid(this.value,<%=rr+1%>);return false;" name="additemid<%=rr+1%>" value="<%=oRelationItemList.FRelationItemlist(rr).FRItemid%>">X</button>
														<span style="position:absolute;opacity:0.8;background-color:#FFFFFF"><strong>&nbsp;&nbsp;<%=oRelationItemList.FRelationItemlist(rr).FRItemid%>&nbsp;&nbsp;</strong></span>
														<span style="position:absolute;bottom:5%;opacity:0.8;background-color:#FFFFFF"><%=fnGetLastPrice(oRelationItemList.FRelationItemlist(rr).FSellcash,oRelationItemList.FRelationItemlist(rr).FOrgprice,oRelationItemList.FRelationItemlist(rr).FSaleYN,oRelationItemList.FRelationItemlist(rr).FItemcouponYN,oRelationItemList.FRelationItemlist(rr).FItemcouponValue,oRelationItemList.FRelationItemlist(rr).FitemcouponType)%></span>
														<a href="javascript:void(0);">
															<img id="img<%=rr+1%>" src="<%=oRelationItemList.FRelationItemlist(rr).FRlistimage%>" alt="연관상품">
														</a>
													</div>
												<% Next %>
											<% End If %>
										<% End If %>
									</div>
								</div>
							</div>
							<p class="tPad05 fs11 cGy1">- 연관상품은 최대 10건까지 등록 가능합니다.</p>
						</td>
					</tr>
				<% End If %>
				<% If Trim(gubun)="2" Then %>
					<tr>
						<th><div>연관조각 <strong class="cRd1">*</strong></div></th>
						<td>
							<p><input type="text" name="pieceidx" id="pieceidx" class="formTxt" style="width:100%;" value="<%= pieceidx %>"/></p>
							<p class="tPad05 fs11 cGy1">- 조각번호를 입력해주세요. 최대 10건까지 입력 가능하며, 쉼표(,)로 구분해주세요.</p>
						</td>
					</tr>
				<% End If %>
				<% If Trim(gubun)="1" Or Trim(gubun)="2" Or Trim(gubun)="3" Then %>
					<tr>
						<th><div>시작일 <strong class="cRd1">*</strong></div></th>
						<td>
							<input type="text" id="sDt" name="startdate" size="10" value="<%=Left(startdate,10)%>" /> <input type="text" name="starttime" size="8" value="<%=starttime%>" />
						</td>
					</tr>
				<% End If %>
				<tr>
					<th>작업자 지시사항</th>
					<td ><textarea name="admintext" rows="6" class="formTxtA" style="width:100%; height:80px;" /><%=admintext%></textarea></td>
				</tr>
				<tr>
					<th><div>노출 <strong class="cRd1">*</strong></div></th>
					<td>
						<span class="tPad05 col2">
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="N" <% if isusing<>"Y" then %>checked="checked"<% end if %> /> 비공개</label>
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="Y" <% if isusing="Y" then %>checked="checked"<% end if %> /> 공개</label>
						</span>
						<% If gubun="1" Then %>
							<span class="pad05 col2 bgGry2">
								<label class="lPad10"><input type="checkbox" name="noticeYN" id="noticeYN" value="<%= noticeYN %>" <% if noticeYN="Y" then %>checked="checked"<% end if %> class="formCheck" onclick="openChk();" /> 오프닝</label>
							</span>
						<% End If %>
					</td>
				</tr>
				<tr>
					<th><div>진행상태 <strong class="cRd1">*</strong></div></th>
					<td>
						<select class="formSlt" id="state" name="state" title="옵션 선택">
							<option value="1" <% If state = "1" Then %> selected <% End If %>>등록대기</option>
							<option value="2" <% If state = "2" Then %> selected <% End If %>>이미지 등록요청</option>
							<option value="3" <% If state = "3" Then %> selected <% End If %>>디자인 작업중</option>
							<option value="4" <% If state = "4" Then %> selected <% End If %>>오픈요청</option>
							<option value="7" <% If state = "7" Then %> selected <% End If %>>오픈</option>
							<option value="8" <% If state = "8" Then %> selected <% End If %>>보류</option>
							<option value="9" <% If state = "9" Then %> selected <% End If %>>종료</option>
						</select>
					</td>
				</tr>
				<% If mode = "EDIT" then%>
				<tr>
					<th><div>등록정보</div></th>
					<td>
						<span class="tPad05 col2"><%=occupation%>&nbsp;<%=nickname%> (<%=fnGetMyname(adminid)%>)<br/><%=regdate%></span>
					</td>
				</tr>
				<% If lastadminid <> "" Then %>
				<tr>
					<th><div>최종수정</div></th>
					<td>
						<span class="tPad05 col2 cRd1"><%=LastUpdateAdmin(lastadminid)%> (<%=fnGetMyname(lastadminid)%>)<br/><%=lastupdate%></span>
					</td>
				</tr>
				<% End If %>
				<% End If %>
				</tbody>
			</table>
		</div>
		<div class="popBtnWrap">
			<input type="button" value="취소" onclick="self.close();" style="width:100px; height:30px;" />
			<% if mode = "EDIT"then %>
				<input type="button" value="수정" onclick="frmedit();" class="cRd1" style="width:100px; height:30px;" />
			<% Else %>
				<input type="button" value="등록" onclick="frmedit();" class="cRd1" style="width:100px; height:30px;" />
			<% end if %>

		</div>
	</div>
</form>
<script>
var appendNumber = <%=appendNumberPrv%>;
var swiper;
	swiper = new Swiper('.pdtLinkWrap .swiper-container', {
		slidesPerView:'auto',
		freeMode:true,
		freeModeSticky:true
	});

$(function(){
	var ajaxtagtext
	var sampleTags
	var str = $.ajax({
		type: "POST",
		url: "/admin/sitemaster/piece/ajax_tag.asp",
		data: "mode=admin",
		dataType: "text",
		contentType: "application/x-www-form-urlencoded; charset=UTF-8",
		async: false
	}).responseText;
		var str1 = str.split("||");
		if (str1[0] == "OK"){
			ajaxtagtext =  unescape(str1[1]);
			sampleTags = ajaxtagtext.split(',');
			$('#singleFieldTags').tagit({
				availableTags: sampleTags,
				placeholderText: "#제외 입력",
			});
		}else if (str1[0] == "ERR"){
			alert(str1[1]);
			return false;
		}else{
			alert('오류가 발생했습니다.');
		return false;
		}
<% if false then %>
//	var sampleTags = ['c++', 'java', 'php', 'coldfusion', 'javascript', 'asp', 'ruby', 'python', 'c', 'scala', 'groovy', 'haskell', 'perl', 'erlang', 'apl', 'cobol', 'go', 'lua', 'piece'];
//	var tmpsampleTags = tmpsampleTags = ['c++', 'java', 'php', 'coldfusion', 'javascript', 'asp', 'ruby', 'python', 'c', 'scala', 'groovy', 'haskell', 'perl', 'erlang', 'apl', 'cobol', 'go', 'lua', 'piece'];
//	var tmpstr = new Array;
//	tmpstr = ajaxtagtext;//agtext;///	tmpstr = ajaxtagtext.toString();
//	var sampleTags =tmpstr;

//alert(sampleTags);
//	$('#singleFieldTags').tagit({
//		availableTags: sampleTags,
//		placeholderText: "#제외 입력",
////		Usage : https://github.com/aehlke/tag-it 참고
////		autocomplete: {delay: 0, minLength: 2},
////		singleField: true,
////		singleFieldNode: $('#mySingleField')
//	});
<% end if %>
});
document.querySelector('.btn-append').addEventListener('click', function(e) {
	e.preventDefault();
	appendNumber = ++appendNumber
	if (appendNumber>10)
	{
		alert("연관상품은 최대 10개까지 등록 가능합니다.");
		return false;
	}
	popItemWindow('frm',appendNumber, $("#itemid").val());
	swiper.appendSlide('<div class="swiper-slide" id="itemidimgdiv'+(appendNumber)+'" ><button type="button" onclick="fndelitemid(this.value,'+appendNumber+');" name="additemid'+(appendNumber)+'" value="" >X</button><a href=""><img id="img'+(appendNumber)+'" src="" alt="연관상품" /></a></div>');
	swiper.update();
});


String.prototype.replaceAll = function(org, dest) {
    return this.split(org).join(dest);
}


function fndelitemid(delitemid,appendno){
	var itemid = $('#itemid').val().replace(/ /g, '');

	var str = itemid;
	str = str.replaceAll(delitemid+",","");
	str = str.replaceAll(","+delitemid,"");
	str = str.replaceAll(delitemid,"");
	$('#itemid').val(str);
	$('#itemidimgdiv'+appendno).remove();
	appendNumber = --appendNumber;
}

function replaceAll(content,before,after){
    return content.split(before).join(after);
}

function popItemWindow(tgf,ipindex, itemarr){
	var popup_item = window.open("/admin/sitemaster/piece/pop_singleItemSelect.asp?target=" + tgf + "&ptype=piece" + "&ipindex=" + ipindex + "&itemarr=" + itemarr, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
}
</script>
</body>
</html>
<%
set oPieceView = Nothing
Set oRelationItemList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
