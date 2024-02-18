<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_pcweb_slide.asp
' Discription : PCWEB slide insert
' History : 2016-02-16 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim eCode , title , eFolder , idx , mode
Dim strSql , sqlStr , gubuncls, menuidx
Dim topimg , topaddimg , btmYN , btmimg , btmcode , btmaddimg , pcadd1 , gubun

	eCode = requestCheckvar(request("eC"),16)
	menuidx = requestCheckvar(request("menuidx"),16)
	if menuidx="" or isnull(menuidx) then menuidx=0
	title = "이미지 템플릿 슬라이드 등록 (PC WEB)"

	eFolder = eCode

	If eCode <> "" Then 
		strSql = "SELECT idx , topimg , topaddimg , btmYN , btmimg , btmcode , btmaddimg , pcadd1 , gubun " & vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_event_slide_template where evt_code = '"& eCode &"'" & vbcrlf
		strSql = strSql & " and device = 'W' and menuidx=" & menuidx
		rsget.Open strSql,dbget,1
		IF Not rsget.Eof Then
			idx			= rsget("idx")
			topimg		= rsget("topimg")
			topaddimg	= rsget("topaddimg")
			btmYN		= rsget("btmYN")
			btmimg		= rsget("btmimg")
			btmcode		= rsget("btmcode")
			btmaddimg	= rsget("btmaddimg")
			pcadd1		= rsget("pcadd1")
			gubun		= rsget("gubun")
		End If
		rsget.close()
	End If


	If idx = "" Then
		mode = "I"
	Else
		mode = "U"
	End If

	If gubun = "" Then gubun = 1 '와이드 슬라이드

	'와이드 슬라이드 , 	'와이드 풀단 슬라이드 
	'상단 이미지 , 상단 좌우배경 이미지 
	'하단 이미지 (이미지 or html[코드 대체]), 하단 좌우배경 이미지
	'슬라이드이미지 필수 사항

	'풀단 슬라이드 
	'상단 이미지 , 하단 이미지 (이미지 or html[코드 대체])
	'슬라이드 bg 이미지 - field를 어떤거로??
	'슬라이드이미지 필수 사항

	'punit1 = 1 or 2 = 상단 좌우 배경이미지  , 슬라이드 내부 배경이미지 , 하단 좌우 배경이미지
	'punit2 = 3 슬라이드bg
%>
<!-- #include virtual="/admin/lib/popheaderslide.asp"-->
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
</style>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/academy/itemmaster/viewDIYitem/js/jquery.slides.min.js"></script>
<script type="text/javascript" src="/academy/itemmaster/viewDIYitem/js/swiper-2.1.min.js"></script>
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

function jsSetImg(sFolder, sImg, sName, sSpan , sOpt){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan+'&sOpt='+sOpt,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan , sOpt){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	   if (sOpt != ""){
		   eval("document.all."+sSpan+"_bg").style.background = "";
	   }
	}
}
</script>
<script type="text/javascript">
$(function(){
	dfslide(<%=gubun%>); //좌측 슬라이드 로딩
	<% if btmYN="Y" then %>
	$("#sel1").show();
	<% elseif btmYN="N" then %>
	$("#sel2").show();
	<% end if %>

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

function dfslide(v){
	$("#gubun").val(v);
	$('#preview_ajax').empty();
	var str = $.ajax({
		type: "GET",
		url: "pop_pcweb_slide_ajax.asp",
		data: "eC=<%=eCode%>&menuidx=<%=menuidx%>&gU="+v,
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

function tecopy(v){
	$("#btmcode").attr("value",v);
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

function chkimg(v){
	 $("#btmYN").val(v);
	 if (v == "Y"){
		$("#sel1").show();
		$("#sel2").hide();
	 }else{
 		$("#sel1").hide();
		$("#sel2").show();
	 }
}
</script>
</head>
<body>
<div class="slideRegister adminPc">
	<h1><%=title%>&nbsp;&nbsp;/&nbsp;&nbsp;<a href="/admin/eventmanage/event/v5/template/slide/pop_mobile_slide.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>" style="color:#FFFFFF">(MOBILE)</a></h1>
	<%'슬라이드 화면 불러오기%>
	<div class="preview" id="preview_ajax"></div>
	<%'슬라이드 화면 불러오기%>
	<div class="register">
		<h2>컨텐츠 등록</h2>
		<dl>
			<dt>- 상단<span>(생략가능)</span></dt>
			<dd>
				<table class="tbType1 listTb">
					<tr>
						<th width="15%">상단 이미지 :</th>
						<td class="lt">
							<input class="button" type="button" value="이미지 불러오기" name="mtopimg" onClick="jsSetImg('<%=eFolder%>','','topimg','spantopimg','')"/><%IF topimg <> "" THEN %><div><%=topimg%>&nbsp;<a href="javascript:jsDelImg('topimg','spantopimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a></div><% End If %></td>
					</tr>
					<tr class="punit1">
						<th width="15%">상단 배경 이미지 :</th>
						<td class="lt">
							<input class="button" type="button" value="이미지 불러오기" name="mtopaddimg" onclick="jsSetImg('<%=eFolder%>','','topaddimg','spantopaddimg','B')"/>
							<div id="spantopaddimg">
							<%IF topaddimg <> "" THEN %>
								<%=topaddimg%>
								<a href="javascript:jsDelImg('topaddimg','spantopaddimg','B');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
							<%END IF%>
							</div>
						</td>
					</tr>
				</table>
			</dd>
		</dl>

		<dl>
			<dt>- 슬라이드 <span>(필수/3~10개까지 등록)</span></dt>
			<dd>
				<p class="floatImg punit2">슬라이드 백그라운드 이미지 : <input class="button" type="button" value="이미지 불러오기" name="mpcadd1" onclick="jsSetImg('<%=eFolder%>','','pcadd1','spanpcadd1','B')"/>
				<div id="spanpcadd1" class="punit2">
				<%IF pcadd1 <> "" THEN %>
					<%=pcadd1%>
					<a href="javascript:jsDelImg('pcadd1','spanpcadd1','B');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
				<%END IF%>
				</div></p>
				<div class="insertImg">
					<h3>슬라이드 이미지 등록<br/><span style="color:#c80a0a;line-height:2;">※ 상하단 이미지 우선 등록후 슬라이드 이미지를 등록해주세요. 또는 슬라이드 우선 등록후 상하단 이미지를 등록 해주세요. ※<br/>※ 등록버튼 클릭후 좌측 Viewer에 반영 됩니다.※</span></h3>
					<table class="tbType1 listTb tMar10">
						<colgroup class="">
							<col width="5%" /><col width="5%" class="punit1"/><col /><col width="8%" /><col width="8%" />
						</colgroup>
						<tbody>
						<tr>
							<td>
								<div id="spanslideimg"></div>
								<input class="button" type="button" value="이미지 불러오기" name="mslideimg" onClick="jsSetImg('<%=eFolder%>','','slideimg','spanslideimg','');"/>
							</td>
							<td class="punit1">
								<div id="spanbgslideimg"></div>
								<input class="button" type="button" value="배경 이미지" name="mbgslideimg" onClick="jsSetImg('<%=eFolder%>','','bgslideimg','spanbgslideimg','');"/>
							</td>
							<td>
								<div class="selectLink">
									<input type="text" value="링크값 입력(선택)" onclick="showDrop();" id="mlinkurl" onkeyup="linkcopy();" />
									<ul style="display:none;">
										<li onclick="populateTextBox('');">선택안함</li>
										<li onclick="populateTextBox('#group그룹코드');">#group그룹코드</li>
										<li onclick="populateTextBox('#commentarea');">#commentarea(코멘트바로가기)</li>
										<li onclick="populateTextBox('#reviewarea');">#reviewarea(상품후기바로가기)</li>
										<li onclick="populateTextBox('#photocmtarea');">#photocmtarea(포토게시판바로가기)</li>
										<li onclick="populateTextBox('/event/eventmain.asp?eventid=이벤트코드');">/event/eventmain.asp?eventid=이벤트코드</li>
										<li onclick="populateTextBox('/shopping/category_prd.asp?itemid=상품코드');">/shopping/category_prd.asp?itemid=상품코드 (O)</li>
										<li onclick="populateTextBox('/shopping/category_list.asp?disp=카테고리');">/shopping/category_list.asp?disp=카테고리</li>
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
				<input type="hidden" name="device" value="W"/>
				<input type="hidden" name="eventid" value="<%=eCode%>"/>
				<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
				<div class="tMar20">
					<p>
						<input type="button" class="btn" value="전체 선택" onclick="chkAllItem();">
						<input type="button" class="btn" value="상태 저장" onClick="saveList();" title="표시순서 및 사용여부를 일괄저장합니다.">
					</p>
					<table class="tbType1 listTb tMar10">
						<colgroup>
							<col width="6%" /><col width="12%" class="punit1"/><col width="8%" /><col /><col width="7%" /><col width="11%" />
						</colgroup>
						<thead>
						<tr>
							<th><input type="checkbox" onclick="chkAllItem();"/></th>
							<th>이미지</th>
							<th class="punit1">배경이미지</th>
							<th>링크(선택)</th>
							<th>순서</th>
							<th>사용여부</th>
						</tr>
						</thead>
						<tbody id="subList">
						<% 
							If eCode <> "" Then 

							sqlStr = "SELECT idx , slideimg , linkurl , sorting , isusing , bgimg " & vbcrlf
							sqlStr = sqlStr & " from db_event.[dbo].[tbl_event_slide_addimage] where evt_code = '"& eCode &"'" & vbcrlf
							sqlStr = sqlStr & " and device ='W' and menuidx=" & menuidx & vbcrlf
							sqlStr = sqlStr & " order by sorting asc , idx asc " 
							rsget.Open sqlStr,dbget,1
							if Not(rsget.EOF or rsget.BOF) Then
								Do Until rsget.eof
						%>
						<tr class="<%=chkIIF(rsget("isusing")="N","bgGry1","")%>">
							<td><input type="checkbox" name="chkIdx" value="<%=rsget("idx")%>" /></td>
							<td><img src="<%=rsget("slideimg")%>" style="width:120px;" /></td>
							<td class="punit1"><img src="<%=rsget("bgimg")%>" style="width:20px;height:20px;" /></td>
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
			<dd>
				<table class="tbType1 listTb">
					<tr>
						<th width="15%">하단 이미지 :</th>
						<td class="lt">
							<p>
								<span><input type="radio" name="mbtmyn" value="Y" <%=chkiif(btmyn="Y" ,"checked","")%> onclick="chkimg(this.value);"/> 이미지</span>
								<span class="lMar10"><input type="radio" name="mbtmyn" value="N" <%=chkiif(btmyn="N","checked","")%> onclick="chkimg(this.value);"/> HTML</span>
								<%IF btmimg <> "" THEN %><br/><br/><%=btmimg%>&nbsp;<a href="javascript:jsDelImg('btmimg','spanbtmimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a><% End If %>
							</p>
							<div class="tMar10" id="sel1" style="display:none"><input class="button" type="button" value="이미지 불러오기" name="mbtmimg" onClick="jsSetImg('<%=eFolder%>','','btmimg','spanbtmimg','')"/></div>
							<div class="tMar10" id="sel2" style="display:none"><textarea cols="50" rows="5" style="width:100%; height:200px;" name="mbtmcode" onkeyup="tecopy(this.value);"><%=db2html(btmcode)%></textarea></div>
						</td>
					</tr>
					<tr class="punit1">
						<th width="15%">하단 배경 이미지 :</th>
						<td class="lt">
							<input class="button" type="button" value="이미지 불러오기" name="mbtmaddimg" onClick="jsSetImg('<%=eFolder%>','','btmaddimg','spanbtmaddimg','B')"/>
							<div id="spanbtmaddimg">
							<%IF btmaddimg <> "" THEN %>
								<%=btmaddimg%>
								<a href="javascript:jsDelImg('btmaddimg','spanbtmaddimg','B');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
							<%END IF%>
							</div>
						</td>
					</tr>
				</table>
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
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
<input type="hidden" name="mode" value="SD"/>
<input type="hidden" name="device" value="W"/>
<input type="hidden" name="chkIdx" />
</form>
<form name="slideimgfrm" method="post" action="pop_slide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
<input type="hidden" name="mode" value="SI"/>
<input type="hidden" name="device" value="W"/>
<input type="hidden" name="slideimg" value=""/>
<input type="hidden" name="linkurl" id="linkurl" value=""/>
<input type="hidden" name="bgslideimg" value=""/>
</form>
<form name="slidefrm" method="post" action="pop_slide_proc.asp" style="margin:0px;">
<input type="hidden" name="idx" value="<%=idx%>"/>
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
<input type="hidden" name="mode" value="<%=mode%>"/>
<input type="hidden" name="device" value="W"/>
<input type="hidden" name="topimg" value="<%=topimg%>"/>
<input type="hidden" name="btmimg" value="<%=btmimg%>"/>
<input type="hidden" name="topaddimg" value="<%=topaddimg%>"/>
<input type="hidden" name="btmYN" id="btmYN" value="<%=btmYN%>"/>
<input type="hidden" name="btmaddimg" value="<%=btmaddimg%>"/>
<input type="hidden" name="pcadd1" value="<%=pcadd1%>"/>
<input type="hidden" name="gubun" id="gubun" value="<%=gubun%>"/>
<textarea cols="0" rows="0" style="display:none;" id="btmcode" name="btmcode"><%=db2html(btmcode)%></textarea>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->