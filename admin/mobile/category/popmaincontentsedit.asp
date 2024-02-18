<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : popMainContentsEdit.asp
' Discription : 모바일 사이트 메인 아이템 작성/수정
' History : 2010.02.23 허진원 생성
'           2012.02.14 허진원 - 미니달력 교체
'           2012.12.14 이종화 - alt 값 추가
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/mobile/mo_catetoryMainManageCls.asp" -->
<%
Dim makerid, eC, sqlStr, imgURL
dim idx, poscode, reload
	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	eC = request("eC")
	if idx="" then idx=0

	if reload="on" then
	    response.write "<script>opener.location.reload(); window.close();</script>"
	    dbget.close()	:	response.End
	end if

	dim oMainContents
		set oMainContents = new CMainContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneMainContents
		imgURL = oMainContents.FOneItem.Fmain_image

dim orderidx
	if oMainContents.FOneItem.fview_order = "" then
		orderidx = 99
	else
		orderidx = oMainContents.FOneItem.fview_order
	end if

dim dateOption
dateOption = request("dateoption")	

if dateOption = "" then
	dateOption = date
end if

%>
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="http://m.10x10.co.kr/lib/css/main.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
.text-bnr .headline {padding:0; background-color:transparent; border:none; color:#0d0d0d;}
.text-bnr .thumbnail img {width:100%;}
</style>
<script src="http://code.jquery.com/jquery-latest.min.js"></script>
<script src="http://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
	$(function(){
		initiateDate();
	})

	function initiateDate(){
		var date = '<%=dateOption%>'
		var startDateEle = document.getElementById('start_date');
		var endDateEle = document.getElementById('end_date');

		if(date != '' && startDateEle.value == '' && endDateEle.value == '' ){		
			startDateEle.value = date;
			endDateEle.value = date;
		}		
	}	

	function SaveMainContents(frm){

	    if (frm.catecode.value.length<1){
	        alert('카테고리를 먼저 선택 하세요.');
	        frm.catecode.focus();
	        return;
	    }

        if (frm.imgURL.value.length<1){
	        alert('이미지를 업로드 해 주세요.');
	        frm.imgURL.focus();
	        return;
	    }

        if (frm.main_text.value.length<1){
	        alert('메인 카피를 입력 하세요.');
	        frm.main_text.focus();
	        return;
	    }

	    if (frm.banner_link.value.length<1){
	        alert('링크 값을 입력 하세요.');
	        frm.banner_link.focus();
	        return;
	    }

	    if (frm.start_date.value.length!=10){
	        alert('시작일을 입력  하세요.');
	        return;
	    }

	    if (frm.end_date.value.length!=10){
	        alert('종료일을 입력  하세요.');
	        return;
	    }

	    var vstartdate = new Date(frm.start_date.value.substr(0,4), (1*frm.start_date.value.substr(5,2))-1, frm.start_date.value.substr(8,2));
	    var venddate = new Date(frm.end_date.value.substr(0,4), (1*frm.end_date.value.substr(5,2))-1, frm.end_date.value.substr(8,2));

	    if (vstartdate>venddate){
	        alert('종료일이 시작일보다 빠르면 안됩니다.');
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

	function ChangeGubun(comp){
	    location.href = "?poscode=" + comp.value;
	    // nothing;
	}

	// poscode 2071 추가 작업
	function chkopt(v){
		if (v == "2"){
			document.getElementById("culopt").style.display = "";
			document.getElementById("playopt").style.display = "none"; //리뉴얼시 주석
			document.getElementById("callcontents").style.display = "";
		}else if (v == "3"){
			document.getElementById("culopt").style.display = "none";
			document.getElementById("playopt").style.display = ""; //리뉴얼시 주석
			document.frmcontents.maincopy.value = "PLAYing";
		}else{
			document.getElementById("culopt").style.display = "none";
			document.getElementById("playopt").style.display = "none"; //리뉴얼시 주석
			document.frmcontents.maincopy.value = "HITCHHIKER";
		}
	}
</script>
<script type="text/javascript">
	function fileInfo(f){
		var file = f.files; // files 를 사용하면 파일의 정보를 알 수 있음

		var reader = new FileReader(); // FileReader 객체 사용
		reader.onload = function(rst){ // 이미지를 선택후 로딩이 완료되면 실행될 부분
			$('#img_box').empty().html('<img src="' + rst.target.result + '">'); // append 메소드를 사용해서 이미지 추가
			// 이미지는 base64 문자열로 추가
			// 이 방법을 응용하면 선택한 이미지를 미리보기 할 수 있음
            $("#imgURL").val(rst.target.result);
		}
		reader.readAsDataURL(file[0]); // 파일을 읽는다, 배열이기 때문에 0 으로 접근
	}

	// typing
	function textclone(k,v){
		var frmtext = $("#"+k);
		frmtext.bind("keyup",function(){
			var msg = $(this).val();
			if ($(this).val().length > 0){
				msg = msg.replace(/(?:\r\n|\r|\n)/g, '<br>');
				$("#"+v).html(msg);
			}else{
				$("#"+v).html("");
			}
		});
	}

	$(function() {
		$('input:radio[name="salediv"]').click(function(){
			if($('input:radio[name="salediv"]:checked').val()==1)
			{
				$("#saleinfo1").show();
				//$("#saleinfo2").hide();
			}
			else
			{
				$("#saleinfo1").hide();
				alert('이벤트 코드를 입력 해주세요');
				$("#saleinfo2").focus();
				//$("#saleinfo2").show();
			}
		});
		<% if eC <> "" then %>
		chkopt("2");
		<% end if %>
	});

	function cultureloadpop(){
		winLast = window.open('/admin/sitemaster/lib/pop_culturelist.asp?gubun=MC&poscode=<%=poscode%>&pidx=<%=idx%>','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}
</script>
</head>
<body>
<div class="popWinV17">
	<h1>등록</h1>
	<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/doMobileCateRollingBannerReg.asp" onsubmit="return false;" enctype="multipart/form-data">
		<div class="popContainerV17 pad30">
			<div class="ftLt col6">
				<table class="tbType1 writeTb">
					<tr>
						<th width="160">Idx</th>
						<td>
							<% if oMainContents.FOneItem.Fidx<>"" then %>
								<%= oMainContents.FOneItem.Fidx %>
								<input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
							<% else %>
							<% end if %>
						</td>
					</tr>
					<tr>
						<th width="160">카테고리</th>
						<td>
							<% DrawSelectBoxDispCateLarge "catecode", oMainContents.FOneItem.Fcatecode, "" %>
						</td>
					</tr>
					<tr>
						<th width="160">우선순위</th>
						<td>
							<input type="text" name="orderidx" class="formTxt" size=5 value="<%= oMainContents.FOneItem.Fview_order %>">
						</td>
					</tr>
					<tr>
						<th width="160">이미지</th>
						<td>
							<input type="file" name="file1" value="" size="32" maxlength="32" class="formFile" accept="image/*" onchange="fileInfo(this)">
							<input type="hidden" name="imgURL" id="imgURL" value="<%= imgURL %>">
							<% if oMainContents.FOneItem.Fidx<>"" then %>
							<br>
								<% If imgURL = "" Then %>
									<img src="/images/admin_login_logo2.png" alt="" />
								<% else%>
									<img src="<%= imgURL %>" width="500" alt="" />
								<br> <%= imgURL %>
								<% End If %>
							<% end if %>
						</td>
					</tr>
					<tr>
					  <th width="160">메인카피</th>
					  <td><input type="text" name="main_text" value="<%=oMainContents.FOneItem.Fmain_text%>" class="formTxt" style="width:94%" maxlength="128"> </td>
					</tr>
					<tr>
						<th width="160">링크값</th>
						<td>
							<input type="text" class="formTxt" name="banner_link" value="<%= oMainContents.FOneItem.Fbanner_link %>" maxlength="128" style="width:94%">
						</td>
					</tr>
					<tr>
						<th width="160">반영시작일</th>
						<td>
							<span class="rMar10">
							<input id="start_date" name="start_date" value="<%=Left(oMainContents.FOneItem.Fstart_date,10)%>" class="formTxt" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="start_date_trigger" border="0" style="cursor:pointer; vertical-align:top; margin-left:5px;" /></span>
							<span class="rMar10">
								<input type="text" class="formTxt" name="dummy0" value="00:00:00" size="8" readonly />
							</span>
						</td>
					</tr>
					<tr>
						<th width="160">반영종료일</th>
						<td>
							<span class="rMar10"><input id="end_date" name="end_date" value="<%=Left(oMainContents.FOneItem.Fend_date,10)%>" class="formTxt"  size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="end_date_trigger" border="0" style="cursor:pointer; vertical-align:top; margin-left:5px;" /></span>
							<span class="rMar10">
								<input type="text" class="formTxt" name="dummy1" value="23:59:59" size="8" readonly />
							</span>
							<script type="text/javascript">
								var CAL_Start = new Calendar({
									inputField : "start_date", trigger    : "start_date_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_End.args.min = date;
										CAL_End.redraw();
										this.hide();
									}, bottomBar: true, dateFormat: "%Y-%m-%d"
								});
								var CAL_End = new Calendar({
									inputField : "end_date", trigger    : "end_date_trigger",
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
					<tr>
						<th width="160">등록일</th>
						<td>
							<%= oMainContents.FOneItem.Fregdate %> (<%= oMainContents.FOneItem.Freguserid %>)
						</td>
					</tr>
					<tr>
						<th width="160">사용여부</th>
						<td>
							<% if oMainContents.FOneItem.Fview_yn="0" then %>
							<span class="rMar10"><input type="radio" name="view_yn" class="formRadio" value="1">사용함</span>
							<span class="rMar10"><input type="radio" name="view_yn" class="formRadio" value="0" checked >사용안함</span>
							<% else %>
							<span class="rMar10"><input type="radio" name="view_yn" class="formRadio" value="1" checked >사용함</span>
							<span class="rMar10"><input type="radio" name="view_yn" class="formRadio" value="0">사용안함</span>
							<% end if %>
						</td>
					</tr>
				</table>
			</div>
			<div style="position:fixed;left:62%;top:70px;">
				<div class="lPad30 vTop" id="today_preview">
					<%'타입별 템플릿 %>
					<%'rolling image%>
					<div class="text-bnr">
					<section class="" style="width:375px;">
						<div class="thumbnail" id="img_box">
							<% If imgURL="" Then %>
							<img src="/images/admin_login_logo2.png" alt="" />
							<% Else %>
							<img src="<%=imgURL%>" alt="" />
							<% End If %>
						</div>
					</section>
					</div>
				</div>
			</div>
		</div>
		<div class="popBtnWrap">
			<input type="button" value="저장" onClick="SaveMainContents(frmcontents);" class="cRd1" style="width:100px; height:30px;">
		</div>
	</form>
</div>
<%
set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->