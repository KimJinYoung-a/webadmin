<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_pcweb_themeslide.asp
' Discription : PCWEB slide insert
' History : 2019-02-11 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim eCode , title , eFolder , idx , mode
Dim strSql , sqlStr , gubuncls, menuidx
Dim topimg , topaddimg , btmYN , btmimg , btmcode , btmaddimg , pcadd1 , gubun, bgubun, saveafter, mdtheme
	saveafter = requestCheckvar(request("saveafter"),2)
	eCode = requestCheckvar(request("eC"),16)
	bgubun = 3
	mdtheme = requestCheckvar(request("mdtheme"),1)
	menuidx = requestCheckvar(request("menuidx"),10)
	title = "슬라이드 등록 (PC WEB)"

	eFolder = eCode

	If saveafter <>"" Then
	Response.write "<script>window.document.domain='10x10.co.kr';opener.document.location.reload();</script>"
	End If

	If idx = "" Then
		mode = "I"
	Else
		mode = "U"
	End If

	If gubun = "" Then gubun = bgubun '와이드 슬라이드
	
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
window.document.domain = "10x10.co.kr";
function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	chkAllItem();
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 소재를 선택해주세요.");
		return;
	}
	if(confirm("지정하신 목록의 선택 정보를 저장하시겠습니까?")) {
		document.frmList.action="pop_themeslide_proc.asp";
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
	<% if mdtheme="4" then %>
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV6.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan+'&sOpt='+sOpt+'&wid=7800&hei=5000','popImg','width=370,height=150');
	<% else %>
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV6.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan+'&sOpt='+sOpt+'&wid=11400&hei=5600','popImg','width=370,height=150');
	<% end if %>
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
	dfslide(<%=bgubun%>); //좌측 슬라이드 로딩
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
			$(this).find("input[name^='viewidx']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='viewidx']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});

});

function dfslide(v){
	$('#preview_ajax').empty();
	var str = $.ajax({
		type: "GET",
		url: "pop_pcweb_themeslide_ajax.asp",
		data: "eC=<%=eCode%>&gU="+v+"&menuidx=<%=menuidx%>",
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

function fnUploadTypeSelect(type){
	if(type=="I"){
		$("#utypeI").show();
		$("#utypeV").hide();
	}
	else{
		$("#utypeI").hide();
		$("#utypeV").show();
	}
}

function fnuploadvideo(){
	if(document.frmvideo.videolink.value==""){
		alert("동영상 URL을 입력해주세요.");
		return false;
	}
	else{
		document.frmvideo.submit();
	}
}
</script>
</head>
<body>
<div class="slideRegister adminPc">
	<h1><%=title%></h1>
	<%'슬라이드 화면 불러오기%>
	<div class="preview" id="preview_ajax"></div>
	<%'슬라이드 화면 불러오기%>
	<div class="register">
		<h2>컨텐츠 등록</h2>
		<dl>
			<dt>- 슬라이드 <span>(필수/3~10개까지 등록)</span></dt>
			<dd>
				<div id="spanpcadd1" class="punit2">
				<%IF pcadd1 <> "" THEN %>
					<%=pcadd1%>
					<a href="javascript:jsDelImg('pcadd1','spanpcadd1','B');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
				<%END IF%>
				</div></p>
				<div class="insertImg">
					<h3>슬라이드 이미지 등록<br/><span style="color:#c80a0a;line-height:2;">※ 상하단 이미지 우선 등록후 슬라이드 이미지를 등록해주세요. 또는 슬라이드 우선 등록후 상하단 이미지를 등록 해주세요. ※<br/>※ 등록버튼 클릭후 좌측 Viewer에 반영 됩니다.※<br>※ 이미지 사이즈 : <% if mdtheme="4" then %>780 X 500<% else %>1140 x 560<% end if %>※</span></h3><br>
					<!--<input type="radio" name="uploadtype" value="I" onClick="fnUploadTypeSelect('I')" checked/>이미지 <input type="radio" name="uploadtype" value="V" onClick="fnUploadTypeSelect('V')"/>동영상-->
					<table class="tbType1 listTb tMar10" id="utypeI">
						<colgroup class="">
							<col width="80%" /><col width="20%" />
						</colgroup>
						<tbody>
						<tr>
							<td>
								<div id="spanslideimg"></div>
								<input class="button" type="button" value="이미지 불러오기" name="mslideimg" onClick="jsSetImg('<%=eFolder%>','','slideimg','spanslideimg','');"/>
							</td>
							<td>
								<input type="button" class="btn" value="등록" onclick="simgsubmit();">
							</td>
						</tr>
						</tbody>
					</table>
					<form name="frmvideo" method="POST" action="pop_themeslide_proc.asp" style="margin:0;">
					<input type="hidden" name="eventid" value="<%=eCode%>"/>
					<input type="hidden" name="mode" value="SV"/>
					<input type="hidden" name="device" value="W"/>
					<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
					<table class="tbType1 listTb tMar10" id="utypeV" style="display:none">
						<colgroup class="">
							<col width="10%" /><col width="70%" /><col width="20%" />
						</colgroup>
						<tbody>
						<tr>
							<td>
								동영상 URL
							</td>
							<td>
								<input type="text" name="videolink" style="width:400px"/>
							</td>
							<td rowspan="2"><input type="button" class="btn" value="등록" onclick="fnuploadvideo();"></td>
						</tr>
						<tr>
							<td>
								동영상 Type
							</td>
							<td>
								<input type="radio" name="videotype" value="1" checked>16:9 (720*405) <input type="radio" name="videotype" value="2">4:3 (720*540) <input type="radio" name="videotype" value="3">1:1 (720*720)
							</td>
						</tr>						
						</tbody>
					</table>
					</form>
				</div>
				<form name="frmList" method="POST" action="" style="margin:0;">
				<input type="hidden" name="eventid" value="<%=eCode%>"/>
				<input type="hidden" name="mode" value="SU"/>
				<input type="hidden" name="device" value="W"/>
				<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
				<div class="tMar20">
					<table class="tbType1 listTb tMar10">
						<colgroup>
							<col width="10%" /><col width="60%" class="punit1"/><col /><col width="10%" /><col width="20%" />
						</colgroup>
						<thead>
						<tr>
							<th><input type="checkbox" onclick="chkAllItem();"/></th>
							<th>이미지</th>
							<th>순서</th>
							<th>사용여부</th>
						</tr>
						</thead>
						<tbody id="subList">
						<% 
							If eCode <> "" Then
								sqlStr = "SELECT idx , imgurl, viewidx , videoLink, isusing" + vbcrlf
								sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_multi_contents] where menuidx = '"& menuidx &"' and device ='W' " 
								sqlStr = sqlStr & " order by viewidx asc, idx asc"
								rsget.Open sqlStr,dbget,1
								if Not(rsget.EOF or rsget.BOF) Then
									Do Until rsget.eof
						%>
						<tr class="<%=chkIIF(rsget("isusing")="N","bgGry1","")%>">
							<td><input type="checkbox" name="chkIdx" value="<%=rsget("idx")%>" /></td>
							<% if rsget("videoLink")<>"" then %>
							<td><%=rsget("videoLink")%></td>
							<% else %>
							<td><img src="<%=rsget("imgurl")%>" style="width:120px;" /></td>
							<% end if %>
							<td><input type="text" value="<%=rsget("viewidx")%>" name="viewidx<%=rsget("idx")%>"/></td>
							<td>
								<span><input type="radio" <%=chkIIF(rsget("isusing")="Y","checked","")%> name="use<%=rsget("idx")%>" value="Y"/> Y</span>
								<span class="lMar10"><input type="radio" <%=chkIIF(rsget("isusing")="N","checked","")%> name="use<%=rsget("idx")%>" value="N"/> N</span>
								&nbsp;&nbsp;&nbsp;<input type="button" class="btn" value="삭제" onclick="slideimgDel(<%=rsget("idx")%>);">
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
				</div>
				</form>
			</dd>
		</dl>
		<div class="btnArea">
			<input type="image" src="http://scm.10x10.co.kr/images/icon_save.gif" alt="저장" onclick="saveList();"/>
			<a href=""><img src="http://scm.10x10.co.kr/images/icon_cancel.gif" alt="취소" /></a>
		</div>
	</div>
</div>
<form name="frmdel" method="POST" action="pop_themeslide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SD"/>
<input type="hidden" name="device" value="W"/>
<input type="hidden" name="chkIdx" />
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
</form>
<form name="slideimgfrm" method="post" action="pop_themeslide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SI"/>
<input type="hidden" name="device" value="W"/>
<input type="hidden" name="slideimg" value=""/>
<input type="hidden" name="linkurl" id="linkurl" value=""/>
<input type="hidden" name="bgslideimg" value=""/>
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
</form>
<form name="slidefrm" method="post" action="pop_themeslide_proc.asp" style="margin:0px;">
<input type="hidden" name="idx" value="<%=idx%>"/>
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="<%=mode%>"/>
<input type="hidden" name="device" value="W"/>
<input type="hidden" name="topimg" value="<%=topimg%>"/>
<input type="hidden" name="btmimg" value="<%=btmimg%>"/>
<input type="hidden" name="topaddimg" value="<%=topaddimg%>"/>
<input type="hidden" name="btmYN" id="btmYN" value="<%=btmYN%>"/>
<input type="hidden" name="btmaddimg" value="<%=btmaddimg%>"/>
<input type="hidden" name="pcadd1" value="<%=pcadd1%>"/>
<input type="hidden" name="gubun" id="gubun" value="<%=bgubun%>"/>
<textarea cols="0" rows="0" style="display:none;" id="btmcode" name="btmcode"><%=db2html(btmcode)%></textarea>
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->