<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_mobile_slide.asp
' Discription : 모바일 slide insert
' History : 2016-02-16 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim eCode , title , eFolder , topimg , btmimg , topaddimg 'floating img
Dim videoSize, videoLink '동영상 추가
Dim slideimg
Dim mode , idx , strSql , sqlStr , isarrow

	eCode = requestCheckvar(request("eC"),16)
	title = "슬라이드 등록 팝업(M)"

	eFolder = eCode

	If eCode <> "" Then 
		strSql = "SELECT idx , topimg , btmimg , topaddimg, videosize, videolink , isarrow " + vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_event_slide_template where evt_code = '"& eCode &"' and device = 'M'" 
		rsget.Open strSql,dbget,1
		IF Not rsget.Eof Then
			idx			= rsget("idx")
			topimg		= rsget("topimg")
			btmimg		= rsget("btmimg")
			topaddimg	= rsget("topaddimg")
			videosize	= rsget("videosize")
			videolink	= rsget("videolink")
			isarrow		= rsget("isarrow")
		End If
		rsget.close()
	End If 

	If idx = "" Then
		mode = "I"
	Else
		mode = "U"
	End If

%>
<!-- #include virtual="/admin/lib/popheaderslide.asp"-->
<link rel="stylesheet" type="text/css" href="http://webadmin.10x10.co.kr/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="http://webadmin.10x10.co.kr/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
.vod-wrap .vod {overflow:hidden; position:relative; width:100%; height:100%; padding-bottom:100%; /padding-bottom:70.4%;/}
.vod-wrap .vod iframe {position:absolute; top:0; left:0; bottom:0; width:100%; height:100%;}
.shape-rtgl .vod {padding-bottom:56.25%;}
</style>
<script type="text/javascript" src="http://m.10x10.co.kr/lib/js/jquery.swiper-3.1.2.min.js"></script>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 소재를 선택해주세요.");
		return;
	}
	if(confirm("지정하신 목록의 선택 정보를 저장하시겠습니까?")) {
		document.frmList.action="pop_slide_proc.asp";
		document.frmList.submit();
	}
}

//'아이템 삭제
function slideimgDel(v){
	if (confirm("상품이 삭제됩니다 삭제 하시겠습니까?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.submit();
	}
}

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<script type="text/javascript">
$(function(){
	dfslide(); //좌측 슬라이드 로딩
	
	//드래그
	$( "#subList").sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});

});

function dfslide(){
	var str = $.ajax({
		type: "GET",
		url: "pop_mobile_slide_ajax.asp",
		data: "eC=<%=eCode%>",
		dataType: "text",
		async: false
		}).responseText;
	if (str != ""){
		$('#preview_ajax').append(str);
	}
}

//링크값선택
function showDrop(){
	$(".selectLink ul").show();
}

//선택입력
function populateTextBox(v){
	var val = v;
	$("#mlinkurl").val(val);
	$("#linkurl").val(val);
	$(".selectLink ul").css("display","none");
}

function linkcopy(){
	var val = $("#mlinkurl").val();
	$("#linkurl").attr("value",val);
	$(".selectLink ul").css("display","none");
}

function closel(){
	$(".selectLink ul").css("display","none");
}

function simgsubmit(){
	//슬라이드이미지입력
	var frm = document.slideimgfrm;
	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}
function mimgsubmit(){
	//기타 상하 플로팅
	var frm = document.slidefrm;
	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}

function videosizeins(v){
	var val = v;
	$("#videosize").val(v);
}

function chkisarrow(v){
	var val = v;
	$("#isarrow").val(v);
}

function videolinkins(v){
	var val = v;
	$("#videolink").val(v);
}


$(function(){
    var currentPosition = parseInt($("#preview_ajax").css("top"));
    $(window).scroll(function() {  
         var position = $(window).scrollTop(); // 현재 스크롤바의 위치값을 반환합니다.
		 if (position > 0){
			$("#preview_ajax").stop().animate({"top":position+"px"},500);
		 }else{
			$("#preview_ajax").stop().animate({"top":position+currentPosition+"px"},500);  
		 }
         
    });
});  
</script>
</head>
<body>
<div class="slideRegister adminMob">
	<h1>슬라이드 및 동영상 등록 (MOBILE)</h1>
	<%'슬라이드 화면 불러오기%>
	<div class="preview <% If topaddimg <> "" Then %>txtFix<% End If %>" id="preview_ajax"></div>
	<%'슬라이드 화면 불러오기%>
	<div class="register">
		<h2>컨텐츠 등록</h2>
		<dl>
			<dt>- 상단<span>(생략가능)</span></dt>
			<dd><input type="button" value="이미지 불러오기" name="mtopimg" onClick="jsSetImg('<%=eFolder%>','','topimg','spantopimg')"/><%IF topimg <> "" THEN %><div><br/><%=topimg%>&nbsp;<a href="javascript:jsDelImg('topaddimg','spantopaddimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a></div><%END IF%></dd>
		</dl>
		<dl>
			<dt>- 슬라이드 화살표 사용 여부</dt>
			<dd>
				<input type="radio" value="1" name="isarrowyn" id="arrowYes" <%=chkiif(isarrow="" or isarrow = 1 , "checked" , "")%> onclick="chkisarrow(this.value);"/> <label for="arrowYes">사용함</label>&nbsp;&nbsp;&nbsp;<input type="radio" value="0" name="isarrowyn" id="arrowNo" <%=chkiif(isarrow = 0, "checked" , "")%> onclick="chkisarrow(this.value);"/> <label for="arrowNo">사용안함</label>
			</dd>
		</dl>
		<dl>
			<dt>- 슬라이드 <span>(필수/3~10개까지 등록)</span></dt>
			<dd>
				<p class="floatImg">플로팅 이미지(width:750px, png로 등록) :<input class="button" type="button" value="이미지 불러오기" name="mtopimg" onclick="jsSetImg('<%=eFolder%>','','topaddimg','spantopaddimg')"/></p>
				<div class="insertImg">
					<h3>슬라이드 이미지 등록<br/><span style="color:#c80a0a;line-height:2;">※ 상하단 이미지 우선 등록후 슬라이드 이미지를 등록해주세요. 또는 슬라이드 우선 등록후 상하단 이미지를 등록 해주세요. ※<br/>※ 등록버튼 클릭후 좌측 Viewer에 반영 됩니다.※</span></h3>
					<table class="tbType1 listTb tMar10">
						<colgroup>
							<col width="7%" /><col /><col width="42%" /><col width="7%" /><col width="12%" />
						</colgroup>
						<tbody>
						<tr>
							<td></td>
							<td>
								<div id="spanslideimg"></div>
								<input class="button" type="button" value="이미지 불러오기" name="mslideimg" onClick="jsSetImg('<%=eFolder%>','','slideimg','spanslideimg');"/>
							</td>
							<td>
								<div class="selectLink">
									<input type="text" value="링크값 입력(선택)" onclick="showDrop();" id="mlinkurl" onkeyup="linkcopy();" />
									<ul style="display:none;">
										<li onclick="populateTextBox('');">선택안함</li>
										<li onclick="populateTextBox('#group그룹코드');">#group그룹코드</li>
										<li onclick="populateTextBox('/event/eventmain.asp?eventid=이벤트코드');">/event/eventmain.asp?eventid=이벤트코드</li>
										<li onclick="populateTextBox('/category/category_itemprd.asp?itemid=상품코드');">/category/category_itemprd.asp?itemid=상품코드 (O)</li>
										<li onclick="populateTextBox('/category/category_list.asp?disp=카테고리');">/category/category_list.asp?disp=카테고리</li>
										<li onclick="populateTextBox('/street/street_brand.asp?makerid=브랜드아이디');">/street/street_brand.asp?makerid=브랜드아이디</li>
									</ul>
								</div>
							</td>
							<td><input type="text" value="0" /></td>
							<td>
								<input type="button" class="btn" value="등록" onclick="simgsubmit();">
							</td>
						</tr>
						</tbody>
					</table>
				</div>
				<form name="frmList" method="POST" action="" style="margin:0;">
				<input type="hidden" name="mode" value="SU"/>
				<input type="hidden" name="device" value="M"/>
				<input type="hidden" name="eventid" value="<%=eCode%>"/>
				<div class="tMar20">
					<p>
						<input type="button" class="btn" value="전체 선택" onclick="chkAllItem();">
						<input type="button" class="btn" value="상태 저장" onClick="saveList();" title="표시순서 및 사용여부를 일괄저장합니다.">
					</p>
					<table class="tbType1 listTb tMar10">
						<colgroup>
							<col width="7%" /><col /><col width="42%" /><col width="7%" /><col width="12%" />
						</colgroup>
						<thead>
						<tr>
							<th><input type="checkbox" onclick="chkAllItem();"/></th>
							<th>이미지</th>
							<th>링크(선택)</th>
							<th>순서</th>
							<th>사용여부</th>
						</tr>
						</thead>
						<tbody id="subList">
						<% 
							If eCode <> "" Then 

							sqlStr = "SELECT idx , slideimg , linkurl , sorting , isusing " + vbcrlf
							sqlStr = sqlStr & " from db_event.[dbo].[tbl_event_slide_addimage] where evt_code = '"& eCode &"' and device = 'M' " 
							sqlStr = sqlStr & " order by sorting asc , idx asc " 
							rsget.Open sqlStr,dbget,1
							if Not(rsget.EOF or rsget.BOF) Then
								Do Until rsget.eof
						%>
						<tr class="<%=chkIIF(rsget("isusing")="N","bgGry1","")%>">
							<td><input type="checkbox" name="chkIdx" value="<%=rsget("idx")%>" /></td>
							<td><img src="<%=rsget("slideimg")%>" style="width:100px;" /></td>
							<td class="lt"><input type="text" style="width:400px;" name="linkurl<%=rsget("idx")%>" value="<%=rsget("linkurl")%>"/></td>
							<td><input type="text" value="<%=rsget("sorting")%>" name="sort<%=rsget("idx")%>"/></td>
							<td>
								<span><input type="radio" <%=chkIIF(rsget("isusing")="Y","checked","")%> name="use<%=rsget("idx")%>" value="Y"/> Y</span>
								<span class="lMar10"><input type="radio" <%=chkIIF(rsget("isusing")="N","checked","")%> name="use<%=rsget("idx")%>" value="N"/> N</span>
								<br/><input type="button" class="btn" value="삭제" onclick="slideimgDel(<%=rsget("idx")%>);">
							</td>
						</tr>
						<% 
								rsget.movenext
								Loop
							End If
							rsget.close

							End If
						%>
						</tbody>
					</table>
					<p style="text-align:right;margin-top:5px;">
						<input type="button" class="btn" value="상태 저장" onClick="saveList();" title="표시순서 및 사용여부를 일괄저장합니다.">
					</p>
				</div>
				</form>
			</dd>
		</dl>
		<dl>
			<dt>- 하단<span>(생략가능)</span></dt>
			<dd><input class="button" type="button" value="이미지 불러오기" name="mbtmimg" onClick="jsSetImg('<%=eFolder%>','','btmimg','spanbtmimg')" /><%IF btmimg <> "" THEN %><div><br/><%=btmimg%>&nbsp;<a href="javascript:jsDelImg('btmimg','spanbtmimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a></div><%END IF%></dd>
		</dl>
		<dl>
			<dt>- 동영상<span>(생략가능)</span></dt>
			<dd>
				<span style="color:#c80a0a;line-height:2;">※ 비메오/유튜브 동영상 링크만 가능합니다.※</span>
				<table class="tbType1 listTb tMar10">
					<colgroup>
						<col width="30%" /><col /><col width="42%" />
					</colgroup>
					<thead>
					<tr>
						<th>사이즈</th>
						<th>링크</th>
					</tr>
					</thead>
					<tbody>
					<tr>
						<td><input type="radio" name="videosizechk" id="videosizechk" value="N" <%=chkIIF(videosize="N","checked","")%> onclick="videosizeins(this.value);">정방형&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="videosizechk" value="W" <%=chkIIF(videosize="W","checked","")%> onclick="videosizeins(this.value);">와이드</td>
						<td class="lt"><input type="text" style="width:500px;" name="videolinktxt" id="videolinktxt" value="<%=videolink%>" onkeyup="videolinkins(this.value);"/></td>
					</tr>
					</tbody>
				</table>
				<!--p style="text-align:right;margin-top:5px;">
					<input type="button" class="btn" value="동영상 저장" onClick="saveList();" title="동영상을 저장합니다.">
				</p-->			
			</dd>
		</dl>		
		<div class="btnArea">
			<input type="image" src="http://webadmin.10x10.co.kr/images/icon_save.gif" alt="저장" onclick="mimgsubmit();"/>
			<a href=""><img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" alt="취소" /></a>
		</div>
	</div>
</div>
<form name="frmdel" method="POST" action="pop_slide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SD"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="chkIdx" />
</form>
<form name="slideimgfrm" method="post" action="pop_slide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SI"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="slideimg" value=""/>
<input type="hidden" name="linkurl" id="linkurl" value=""/>
</form>
<form name="slidefrm" method="post" action="pop_slide_proc.asp" style="margin:0px;">
<input type="hidden" name="idx" value="<%=idx%>"/>
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="<%=mode%>"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="topimg" value="<%=topimg%>"/>
<input type="hidden" name="btmimg" value="<%=btmimg%>"/>
<input type="hidden" name="topaddimg" value="<%=topaddimg%>"/>
<input type="hidden" name="videosize" id="videosize" value="<%=videosize%>"/>
<input type="hidden" name="videolink" id="videolink" value="<%=videolink%>"/>
<input type="hidden" name="isarrow" id="isarrow" value="<%=isarrow%>"/>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->