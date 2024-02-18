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
<!-- #include virtual="/lib/classes/mobile/main_ContentsManageCls.asp" -->
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
		imgURL = oMainContents.FOneItem.GetImageUrl

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

if eC <> "" then
	sqlStr = "SELECT e.evt_code, d.eventtype_mo as evt_type, e.evt_name, e.evt_subcopyk, d.evt_mainimg" + vbcrlf
	sqlStr = sqlStr & " from db_event.dbo.tbl_event as e" + vbcrlf
	sqlStr = sqlStr & " LEFT JOIN [db_event].[dbo].[tbl_event_display] as d on d.evt_code=e.evt_code"
	sqlStr = sqlStr & " where e.evt_using='Y'" + vbcrlf
	sqlStr = sqlStr & " and e.evt_code="& eC		
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		oMainContents.FOneItem.Fculopt = rsget("evt_type")
		oMainContents.FOneItem.Fmaincopy = rsget("evt_name")
		oMainContents.FOneItem.Fsubcopy	= rsget("evt_subcopyk")
		imgURL = rsget("evt_mainimg")
		oMainContents.FOneItem.Flinkurl = "/culturestation/culturestation_event.asp?evt_code=" & eC
		oMainContents.FOneItem.Fcgubun = "2"
	end if
	rsget.close
end if

dim dateOption
dateOption = request("dateoption")	

if dateOption = "" then
	dateOption = date
end if

%>
<!DOCTYPE html>
<html xmlns="https://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="https://m.10x10.co.kr/lib/css/main.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
.text-bnr .headline {padding:0; background-color:transparent; border:none; color:#0d0d0d;}
.text-bnr .thumbnail img {width:100%;}
</style>
<script src="https://code.jquery.com/jquery-latest.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
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
		var startDateEle = document.getElementById('startdate');
		var endDateEle = document.getElementById('enddate');

		if(date != '' && startDateEle.value == '' && endDateEle.value == '' ){		
			startDateEle.value = date;
			endDateEle.value = date;
		}		
	}	

	function SaveMainContents(frm){
		<% If poscode = "2069" or oMainContents.FOneItem.Fposcode = "2069" Then %>
		if (frm.makerid.value.length<1){
	        alert('브랜드ID를 등록하세요');
	        frm.makerid.focus();
	        return;
	    }
		<% end if %>

	    if (frm.poscode.value.length<1){
	        alert('구분을 먼저 선택 하세요.');
	        frm.poscode.focus();
	        return;
	    }

		<% if poscode = "2071" or poscode = "2082" then %>
		if (frm.cgubun.value.length<1){
	        alert('컨텐츠구분을 먼저 선택 하세요.');
	        frm.cgubun.focus();
	        return;
	    }

		if (frm.cgubun.value == "3" && frm.backColor.value.length<1){
			alert("배경 컬러를 입력 해주세요");
			frm.backColor.focus();
			return;
	    }
		<% end if %>

		<% if poscode = "2075" Or oMainContents.FOneItem.Fposcode = "2075" Or poscode = "2076" Or oMainContents.FOneItem.Fposcode = "2076" Or poscode = "2077" Or oMainContents.FOneItem.Fposcode = "2077" Or poscode = "2079" Or oMainContents.FOneItem.Fposcode = "2079" Or poscode = "2080" Or oMainContents.FOneItem.Fposcode = "2080" Then %>
	    if (frm.maincopy.value.length<1){
	        alert('메인카피를 입력 하세요.');
	        frm.maincopy.focus();
	        return;
	    }

	    if (frm.subcopy.value.length<1){
	        alert('서브카피를 입력 하세요.');
	        frm.subcopy.focus();
	        return;
	    }
		<% end if%>

		<% if poscode = "2081" Then %>

      let cgubun = $('input:radio[name="cgubun"]:checked').val();
      let evt_code = frm.evt_code.value;
      if(cgubun != '' && evt_code.length < 1) {
          if(cgubun == 'I') alert('상품코드를 입력해주세요');
          else if(cgubun == 'E') alert('기획전/이벤트 코드를 입력해주세요');
          return;
      }

      if(cgubun == '') {
          frm.evt_code.value = '';
      }

		<% end if%>


	    if (frm.linkurl.value.length<1){
	        alert('링크 값을 입력 하세요.');
	        frm.linkurl.focus();
	        return;
	    }

		if (frm.linkurl.value.indexOf("카테고리") > 0 || frm.linkurl.value.indexOf("이벤트번호") > 0 || frm.linkurl.value.indexOf("상품코드") > 0 || frm.linkurl.value.indexOf("브랜드아이디") > 0){
			alert("링크 값을 확인 해주세요.");
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
				frm.linkurl.value='/search/search_item.asp?rect=검색어';
				break;
			case 'event':
				frm.linkurl.value='/event/eventmain.asp?eventid=이벤트번호';
				break;
			case 'itemid':
				frm.linkurl.value='/category/category_itemprd.asp?itemid=상품코드';
				break;
			case 'category':
				frm.linkurl.value='/category/category_detail2020.asp?disp=카테고리';
				break;
			case 'brand':
				frm.linkurl.value='/brand/brand_detail2020.asp?brandid=브랜드아이디';
				break;
			case 'showbanner':
				frm.linkurl.value='/showbanner/show_view.asp?showidx=쇼배너아이디';
				break;
			case 'culture':
				frm.linkurl.value='/culturestation/culturestation_event.asp?evt_code=이벤트아이디';
				break;
			case 'ground':
				frm.linkurl.value='/play/playGround.asp?idx=그라운드번호&contentsidx=컨텐츠번호';
				break;
			case 'styleplus':
				frm.linkurl.value='/play/playStylePlus.asp?idx=스타일플러스번호&contentsidx=컨텐츠번호';
				break;
			case 'fingers':
				frm.linkurl.value='/play/playDesignFingers.asp?idx=핑거스번호&contentsidx=컨텐츠번호';
				break;
			case 'tepisode':
				frm.linkurl.value='/play/playTEpisode.asp?idx=티에피소드번호&contentsidx=컨텐츠번호';
				break;
			case 'gift':
				frm.linkurl.value='/gift/gifttalk/';
				break;
			case 'wish':
				frm.linkurl.value='/wish/index.asp';
				break;
			case 'playing':
				frm.linkurl.value='/playing/view.asp?didx=플레잉번호';
				break;
			case 'couponbook':
				frm.linkurl.value='/my10x10/couponbook.asp'
				break;
		}
	}

	// poscode 2071 추가 작업
	function chkopt(v){
		if (v == "2"){
			document.getElementById("culopt").style.display = "";
			document.getElementById("playopt").style.display = "none"; //리뉴얼시 주석
			//document.frmcontents.maincopy.value = "CULTURE<br/>STATION";
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

    <% If poscode = "2081" Then %>
      $("input[name='cgubun']:radio").change(function () {
          if(this.value == ''){
              $('#evt_code').val('');
              $('#evt_code').attr('disabled',true);
          } else {
              $('#evt_code').attr('disabled',false);
          }
      });
    <% End If %>
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
	<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/doMobileMainContentsReg.asp" onsubmit="return false;" enctype="multipart/form-data">
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
						<th width="160">구분명</th>
						<td>
							<% if oMainContents.FOneItem.Fidx<>"" then %>
							<%= oMainContents.FOneItem.Fposname %> (<%= oMainContents.FOneItem.Fposcode %>)
							<input type="hidden" name="poscode" value="<%= oMainContents.FOneItem.Fposcode %>">
							<% else %>
							<% call DrawMainPosCodeCombo("poscode", poscode, "onChange='ChangeGubun(this);'") %>
							<% end if %>
							<% If poscode = "2069" Or oMainContents.FOneItem.Fposcode = "2069" Then %>
								<%	drawSelectBoxDesignerWithName "makerid", oMainContents.FOneItem.Fmakerid %>
							<% End If %>
						</td>
					</tr>
					<% If poscode = "2071" Or oMainContents.FOneItem.Fposcode = "2071" Or poscode = "2082" Or oMainContents.FOneItem.Fposcode = "2082" Then %>
					<tr>
						<th width="160">컨텐츠구분</th>
						<td>
							<select name="cgubun" onchange="chkopt(this.value);">
								<option value="">구분선택</option>
								<option value="1" <%=chkiif(oMainContents.FOneItem.Fcgubun="1","selected","")%>>히치하이커</option>
								<option value="2" <%=chkiif(oMainContents.FOneItem.Fcgubun="2","selected","")%>>컬처스테이션</option>
								<option value="3" <%=chkiif(oMainContents.FOneItem.Fcgubun="3","selected","")%>>플레잉</option>
							</select>&nbsp;&nbsp;&nbsp;<span id="callcontents" style="display:none"><a href="javascript:cultureloadpop();">불러오기</a></span>
						</td>
					</tr>
					<tr id="culopt" style="display:<%=chkiif(oMainContents.FOneItem.Fcgubun="2","","none")%>;">
						<th width="160">분류</th>
						<td>
							<input type="radio" class="formRadio" name="culopt" id="a1" value="1" <%=chkiif(oMainContents.FOneItem.Fculopt="1" Or oMainContents.FOneItem.Fculopt="","checked","")%>/><label for="a1">연극</label>
							<input type="radio" class="formRadio" name="culopt" id="a2" value="2" <%=chkiif(oMainContents.FOneItem.Fculopt="2","checked","")%>/><label for="a2">뮤지컬</label>
							<input type="radio" class="formRadio" name="culopt" id="a3" value="3" <%=chkiif(oMainContents.FOneItem.Fculopt="3","checked","")%>/><label for="a3">공연</label>
							<input type="radio" class="formRadio" name="culopt" id="a4" value="4" <%=chkiif(oMainContents.FOneItem.Fculopt="4","checked","")%>/><label for="a4">전시</label>
							<input type="radio" class="formRadio" name="culopt" id="a5" value="5" <%=chkiif(oMainContents.FOneItem.Fculopt="5","checked","")%>/><label for="a5">도서</label>
							<input type="radio" class="formRadio" name="culopt" id="a6" value="6" <%=chkiif(oMainContents.FOneItem.Fculopt="6","checked","")%>/><label for="a6">영화</label>
							<input type="radio" class="formRadio" name="culopt" id="a7" value="7" <%=chkiif(oMainContents.FOneItem.Fculopt="7","checked","")%>/><label for="a7">음반</label>
						</td>
					</tr>
					<tr id="playopt" style="display:<%=chkiif(oMainContents.FOneItem.Fcgubun="3","","none")%>;">
						<th width="160">BG컬러</th>
						<td>
							<input type="text" name="backColor" value="<%=oMainContents.FOneItem.FbackColor%>"/>#붙여주세요 ex)#000000
						</td>
					</tr><%'//2017 리뉴얼시 주석처리 %>
					<% End If %>
					<tr>
						<th width="160">링크구분</th>
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
					<tr>
						<th width="160">적용구분(반영주기)</th>
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
					<tr>
						<th width="160">우선순위</th>
						<td>
							<% if oMainContents.FOneItem.Fidx<>"" then %>
								<% if oMainContents.FOneItem.Flinktype="X" or poscode="2085" then %>
								<input type="text" name="orderidx" class="formTxt" size=5 value="<%= orderidx %>">
								<% end if %>
							<% else %>
								<% if poscode<>"" then %>
									<input type="text" name="orderidx" class="formTxt" size=5 value="<%= orderidx %>">
								<% else %>
									<font color="red">구분을 먼저 선택하세요</font>
								<% end if %>
							<% end if %>
						</td>
					</tr>
					<tr>
						<th width="160">이미지</th>
						<td>
							<input type="file" name="file1" value="" size="32" maxlength="32" class="formFile" accept="image/*" onchange="fileInfo(this)">
							<input type="hidden" name="imgURL" value="<%= imgURL %>">
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
						<th width="160">알트명 (필수)</th>
						<td><input type="text" name="altname" value="<%=oMainContents.FOneItem.Faltname%>" class="formTxt" size="50" maxlength="50"> </td>
					</tr>
					<% '2022 Today개선 1차, 이미지배너C %>
					<% If poscode = "2081" Then %>
						<tr>
							<th width="160">메인카피</th>
							<td>
                <textarea name="maincopy" cols="50" rows="2" class="formTxtA" maxlength="60"/><%=oMainContents.FOneItem.Fmaincopy%></textarea>
							</td>
						</tr>
						<tr>
							<th width="160">컨텐츠 연결</th>
							<td>
							  <span class="rMar10"><input type="radio" name="cgubun" value="" <% If oMainContents.FOneItem.Fcgubun="" Then Response.write " checked" %>>사용안함 </span>
                <span class="rMar10"><input type="radio" name="cgubun" value="I" <% If oMainContents.FOneItem.Fcgubun="I" Then Response.write " checked" %>>상품코드 </span>
                <span class="rMar10"><input type="radio" name="cgubun" value="E" <% If oMainContents.FOneItem.Fcgubun="E" Then Response.write " checked" %>>기획전/이벤트 코드 </span>
                <p class="tPad05">
							    <input type="text" name="evt_code" id="evt_code" value="<%=oMainContents.FOneItem.Fevt_code%>" class="formTxt" <% If oMainContents.FOneItem.Fcgubun="" Then Response.write " disabled" %>>
                </p>
							</td>
						</tr>
					<% End If %>
					<% '2022 Today개선 1차, 마케팅배너 %>
					<% If poscode = "2089" Then %>
						<tr>
							<th width="160">라벨</th>
							<td><input type="text" name="subcopy" value="<%=oMainContents.FOneItem.Fsubcopy%>" class="formTxt"></td>
						</tr>
						<tr>
							<th width="160">라벨&배너 배경색</th>
							<td><input type="text" name="backColor" value="<%=oMainContents.FOneItem.fbackColor%>" class="formTxt"></td>
						</tr>
						<tr>
							<th width="160">메인카피</th>
							<td><input type="text" name="maincopy" value="<%=oMainContents.FOneItem.Fmaincopy%>" class="formTxt"></td>
						</tr>
						<tr>
							<th width="160">메인카피 글자색</th>
							<td><input type="text" name="maincopy2" value="<%=oMainContents.FOneItem.Fmaincopy2%>" class="formTxt"></td>
						</tr>
					<% End If %>
					<% '//2022 Today개선 1차, 마케팅배너 %>
					<%'2016 사용%>
					<% If poscode = "2070" Or oMainContents.FOneItem.Fposcode = "2070" Or poscode = "2071" Or oMainContents.FOneItem.Fposcode = "2071" Or poscode = "2082" Or oMainContents.FOneItem.Fposcode = "2082" Then %>
					<tr>
					  <th width="160">메인카피</th>
					  <td><input type="text" name="maincopy" value="<%=oMainContents.FOneItem.Fmaincopy%>" class="formTxt"size="20" maxlength="20"> </td>
					</tr>
					<tr>
					  <th width="160">서브카피</th>
					  <td><input type="text" name="subcopy" value="<%=oMainContents.FOneItem.Fsubcopy%>" class="formTxt"  size="50" maxlength="50"> </td>
					</tr>
					<% End If %>
					<%'2016 사용%>
					<%'2017 사용%>
					<% if poscode = "2075" Or oMainContents.FOneItem.Fposcode = "2075" Or poscode = "2076" Or oMainContents.FOneItem.Fposcode = "2076" Or poscode = "2077" Or oMainContents.FOneItem.Fposcode = "2077" Or poscode = "2079" Or oMainContents.FOneItem.Fposcode = "2079" Or poscode = "2080" Or oMainContents.FOneItem.Fposcode = "2080" Then %>
					<tr>
						<th width="160">메인카피</th>
						<td><input type="text" name="maincopy" id="maincopy" class="formTxt" value="<%=oMainContents.FOneItem.Fmaincopy%>" size="50" maxlength="20" onclick="textclone('maincopy','T_maincopy1');"><br/><strong>※ 최대 20자 제한 ※</strong></td>
					</tr>
					<tr>
						<th width="160">메인카피2</th>
						<td><input type="text" name="maincopy2" id="maincopy2" class="formTxt" value="<%=oMainContents.FOneItem.Fmaincopy2%>" size="50" maxlength="11" onclick="textclone('maincopy2','T_maincopy2');"><br/><strong>※ 최대 11자 제한 ※</strong></td>
					</tr>
					<tr>
						<th width="160">서브카피</th>
						<td ><textarea name="subcopy" cols="50" rows="4" id="subcopy" class="formTxtA" onclick="textclone('subcopy','T_subcopy');" maxlength="60"/><%=oMainContents.FOneItem.Fsubcopy%></textarea><br/><strong>※ 최대 60자 제한 ※</strong></td>
					</tr>
					<% End If %>
					<%
						If poscode = "2075" Or poscode = "2076" Or poscode = "2077" Or poscode = "2079" Or poscode = "2080" Or oMainContents.FOneItem.Fposcode = "2075" Or oMainContents.FOneItem.Fposcode = "2076" Or oMainContents.FOneItem.Fposcode = "2077" Or oMainContents.FOneItem.Fposcode = "2079" Or oMainContents.FOneItem.Fposcode = "2080" Then
					%>
					<tr>
					  <th width="160">태그</th>
						<td>
							<p>
								<span class="rMar10"><input type="checkbox" name="tag_only" id="tag_only" value="Y" <%=chkiif(oMainContents.FOneItem.Ftag_only = "Y","checked","")%>/> <label for="tag_gift">단독</label></span>							
								<span class="rMar10"><input type="checkbox" name="tag_gift" id="tag_gift" value="Y" <%=chkiif(oMainContents.FOneItem.Ftag_gift = "Y","checked","")%>/> <label for="tag_gift">GIFT</label></span>
								<span class="rMar10"><input type="checkbox" name="tag_plusone" id="tag_plusone" value="Y" <%=chkiif(oMainContents.FOneItem.Ftag_plusone = "Y","checked","")%>/> <label for="tag_plusone">1+1</label></span>
								<span class="rMar10"><input type="checkbox" name="tag_launching" id="tag_launching" value="Y" <%=chkiif(oMainContents.FOneItem.Ftag_launching = "Y","checked","")%>/> <label for="tag_launching">런칭</label></span>
								<span class="rMar10"><input type="checkbox" name="tag_actively" id="tag_actively" value="Y" <%=chkiif(oMainContents.FOneItem.Ftag_actively = "Y","checked","")%>/> <label for="tag_actively">참여(코멘트, 게시판 , 상품후기)</label>&nbsp;</span>
							</p>
							<p class="tPad05">
								<span class="rMar10"><strong>※ 우선순위 : 단독 > GIFT > 1+1 > 런칭 > 참여</strong><br/>&nbsp;&nbsp;&nbsp;한가지만 선택 하세요.</strong></span>
							</p>
						</td>
					</tr>
					<tr>
						<th width="160">할인/쿠폰</th>
						<td>
							<p>
							<input type="radio" name="salediv" value="1" <% If oMainContents.FOneItem.Fsalediv="1" Then Response.write " checked" %>>직접입력
							&nbsp;&nbsp;
							<input type="radio" name="salediv" value="2" <% If oMainContents.FOneItem.Fsalediv="2" Then Response.write " checked" %>>할인율 자동 노출(A타입)
							&nbsp;&nbsp;
							<input type="radio" name="salediv" value="3" <% If oMainContents.FOneItem.Fsalediv="3" Then Response.write " checked" %>>할인율 자동 노출(B타입)
							</p>
							<span id="saleinfo1" style="display:<% If oMainContents.FOneItem.Fsalediv="2" Then Response.write "none" %>">
							<p class="tPad05"><input type="text" class="formTxt" name="sale_per" value="<%=oMainContents.FOneItem.Fsale_per%>"> : 할인율 ex)<strong class="cRd1">~45%</strong></p>
							<p class="tPad05"><input type="text" class="formTxt" name="coupon_per" value="<%=oMainContents.FOneItem.Fcoupon_per%>"> : 쿠폰할인율 ex)<strong class="cGn1">10%</strong></p>
						</td>
					</tr>
					<tr>
						<th width="160">이벤트코드</th>
						<td>
							<span><p class="tPad05"><input type="text" id="saleinfo2" class="formTxt" name="evt_code" value="<%=oMainContents.FOneItem.Fevt_code%>"></p></span>
							<p class="tPad05">
								<span class="rMar10"><strong>※ 할인율 자동 노출 및 이벤트 상태 체크 (대기중 , 종료) 노출 X ※</strong></span>
							</p>
						</td>
					</tr>
					<%
						End If
					%>
					<%'2017 사용%>
					<tr>
						<th width="160">이미지Width</th>
						<td>
							<% if oMainContents.FOneItem.Fidx<>"" then %>
							<input type="text" class="formTxt" name="imagewidth" value="<%= oMainContents.FOneItem.Fimagewidth %>" size="8" maxlength="16">
							<% else %>
								<% if poscode<>"" then %>
									<%= oposcode.FOneItem.Fimagewidth %>
								<% else %>
									구분을 먼저 선택하세요
								<% end if %>
							<% end if %>
						</td>
					</tr>
					<tr>
						<th width="160">이미지Height</th>
						<td>
							<% if oMainContents.FOneItem.Fidx<>"" then %>
								<input type="text" class="formTxt" name="imageheight" value="<%= oMainContents.FOneItem.Fimageheight %>" size="8" maxlength="16">
								<% else %>
								<% if poscode<>"" then %>
									<%= oposcode.FOneItem.Fimageheight %>
								<% else %>
									구분을 먼저 선택하세요
								<% end if %>
							<% end if %>
						</td>
					</tr>
					<tr>
						<th width="160">작업자 지시사항</th>
						<td ><textarea name="ordertext" rows="8" class="formTxtA" style="width:94%;" /><%=oMainContents.FOneItem.Fordertext%></textarea></td>
					</tr>
					<tr>
						<th width="160">링크값</th>
						<td>
							<% if oMainContents.FOneItem.Fidx<>"" then %>
								<% if oMainContents.FOneItem.FLinkType="M" then %>
								<textarea name="linkurl" cols="60" rows="6" class="formTxtA"><%= oMainContents.FOneItem.Flinkurl %></textarea>
								<% else %>
								<input type="text" class="formTxt" name="linkurl" value="<%= oMainContents.FOneItem.Flinkurl %>" maxlength="128" style="width:94%">
								<% end if %>
							<% else %>
								<% if poscode<>"" then %>
									<% if oposcode.FOneItem.FLinkType="M" then %>
										<textarea name="linkurl" cols="60" rows="6" class="formTxtA"><%= defaultMapStr %></textarea>
										<br>(이미지맵 변수값 변경 금지)
									<% else %>
										<p>
											<input type="text" name="linkurl" value="" maxlength="128" class="formTxt" style="width:94%;">
										</p>
										<p class="tPad05">
										<% If poscode = "2061" or poscode = "2070" or poscode = "2071" or poscode = "2082" Then %>
										<font color="#707070">
										- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
										- <span style="cursor:pointer" onClick="putLinkText('culture')">컬처스테이션 : /culturestation/culturestation_event.asp?evt_code=<font color="darkred">이벤트코드</font></span><br>
										- <span style="cursor:pointer" onClick="putLinkText('gift')">기프트 : /gift/gifttalk/</span><br>
										- <span style="cursor:pointer" onClick="putLinkText('playing')">플레잉 : /playing/view.asp?didx=<font color="darkred">플레잉번호</font></span><br>
										- <span style="cursor:pointer" onClick="putLinkText('wish')">위시 : /wish/index.asp</span><br>
										- <span style="cursor:pointer" onClick="putLinkText('event')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
										</font>
										<% Else %>
										<font color="#707070">
										<% If poscode="2086" Then %>
											- <font color="red"><strong>mobile 전용(앱메인배너는 [APP]앱관리->앱구동배너 에서 설정하실 수 있습니다.)</strong></font> - <br/>
										<% Else %>
											- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
										<% End If %>
										- <span style="cursor:pointer" onClick="putLinkText('event')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
										- <span style="cursor:pointer" onClick="putLinkText('itemid')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
										- <span style="cursor:pointer" onClick="putLinkText('category')">카테고리 링크 : /category/category_detail2020.asp?disp=<font color="darkred">카테고리</font></span><br/>
										- <span style="cursor:pointer" onClick="putLinkText('brand')">브랜드아이디 링크 : /brand/brand_detail2020.asp?brandid=<font color="darkred">브랜드아이디</font></span><br/>
										</font>
										- <span style="cursor:pointer" onClick="putLinkText('couponbook')">쿠폰북 : /my10x10/couponbook.asp</span><br/>
										</font>
										<% End If %>
										</p>
									<% end if %>
								<% else %>
								<font color="red">구분을 먼저 선택하세요</font>
								<% end if %>
							<% end if %>
						</td>
					</tr>
					<%
					'// 포스코드 1003일때
					If poscode = "1003" or oMainContents.FOneItem.Fposcode = "1003" Then
					%>
					<tr>
						<th width="160">배경색</th>
						<td>
							<% If oMainContents.FOneItem.Fidx<>"" then %>
								<input type="text" name="backColor" value="<%= oMainContents.FOneItem.fbackColor %>" >#붙여주세요 ex)#000
							<% Else %>
								<% If poscode<>"" Then %>
								<input type="text" name="backColor" value="" >#붙여주세요 ex)#000
								<% Else %>
								<font color="red">구분을 먼저 선택하세요</font>
							<%	   End If
							   End if %>
						</td>
					</tr>
					<% End If %>
					<tr>
						<th width="160">반영시작일</th>
						<td>
							<span class="rMar10">
							<input id="startdate" name="startdate" value="<%=Left(oMainContents.FOneItem.Fstartdate,10)%>" class="formTxt" size="10" maxlength="10" /><img src="https://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer; vertical-align:top; margin-left:5px;" /></span>
							<span class="rMar10">
							<% if oMainContents.FOneItem.Fidx<>"" then %>
								<% if oMainContents.FOneItem.Ffixtype="R" then %>
								<input type="text" class="formTxt" name="startdatetime" size="2" maxlength="2" value="<%= Format00(2,Hour(oMainContents.FOneItem.Fstartdate)) %>" />(시 00~23)
								<input type="text" class="formTxt" name="dummy0" value="00:00" size="6" readonly />
								<% else %>
								<input type="text" class="formTxt" name="dummy0" value="00:00:00" size="8" readonly />
								<% end if %>
							</span>
							<span class="rMar10">
							<% else %>
								<% if poscode<>"" then %>
									<% if oposcode.FOneItem.Ffixtype="R" then %>
									<input type="text" class="formTxt" name="startdatetime" size="2" maxlength="2" value="<%= Format00(2,Hour(oMainContents.FOneItem.Fstartdate)) %>" />(시 00~23)
									<input type="text" class="formTxt" name="dummy0" value="00:00" size="6" readonly />
									<% else %>
									<input type="text" class="formTxt" name="dummy0" value="00:00:00" size="8" readonly />
									<% end if %>
								<% end if %>
							</span>
							<% end if %>
						</td>
					</tr>
					<tr>
						<th width="160">반영종료일</th>
						<td>
							<span class="rMar10"><input id="enddate" name="enddate" value="<%=Left(oMainContents.FOneItem.Fenddate,10)%>" class="formTxt"  size="10" maxlength="10" /><img src="https://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer; vertical-align:top; margin-left:5px;" /></span>
							<span class="rMar10">
							<% if oMainContents.FOneItem.Fidx<>"" then %>
								<% if oMainContents.FOneItem.Ffixtype="R" then %>
								<input type="text" class="formTxt" name="enddatetime" size="2" maxlength="2" value="<%= ChkIIF(oMainContents.FOneItem.Fenddate="","23",Format00(2,Hour(oMainContents.FOneItem.Fenddate))) %>">(시 00~23)
								<input type="text" class="formTxt" name="dummy1" value="59:59" size="6" readonly />
								<% else %>
								<input type="text" class="formTxt" name="dummy1" value="23:59:59" size="8" readonly />
								<% end if %>
							</span>
							<span class="rMar10">
							<% else %>
								<% if poscode<>"" then %>
									<% if oposcode.FOneItem.Ffixtype="R" then %>
									<input type="text" class="formTxt" name="enddatetime" size="2" maxlength="2" value="<%= ChkIIF(oMainContents.FOneItem.Fenddate="","23",Format00(2,Hour(oMainContents.FOneItem.Fenddate))) %>">(시 00~23)
									<input type="text" class="formTxt" name="dummy1" value="59:59" size="6" readonly />
									<% else %>
									<input type="text" class="formTxt" name="dummy1" value="23:59:59" size="8" readonly />
									<% end if %>
								<% end if %>
							</span>
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
					<tr>
						<th width="160">등록일</th>
						<td>
							<%= oMainContents.FOneItem.Fregdate %> (<%= oMainContents.FOneItem.Freguserid %>)
						</td>
					</tr>
					<tr>
						<th width="160">사용여부</th>
						<td>
							<% if oMainContents.FOneItem.Fisusing="N" then %>
							<span class="rMar10"><input type="radio" name="isusing" class="formRadio" value="Y">사용함</span>
							<span class="rMar10"><input type="radio" name="isusing" class="formRadio" value="N" checked >사용안함</span>
							<% else %>
							<span class="rMar10"><input type="radio" name="isusing" class="formRadio" value="Y" checked >사용함</span>
							<span class="rMar10"><input type="radio" name="isusing" class="formRadio" value="N">사용안함</span>
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
						<div class="desc">
						<!--<span class="label label-speech" id="T_discount"><b class="discunt">10%</b></span> -->
							<h2 class="headline"><span id="T_maincopy1"><%=oMainContents.FOneItem.Fmaincopy%></span><br/><span id="T_maincopy2"><%=oMainContents.FOneItem.Fmaincopy2%></span></h2>
							<p class="subcopy" id="T_subcopy"><%=nl2br(oMainContents.FOneItem.Fsubcopy)%></p>
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
set oposcode = Nothing
set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->