<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_mobile_top_slide.asp
' Discription : 모바일 top slide insert
' History : 2019-02-12 정태훈
'			2019-10-02 정태훈	템플릿 컨텐츠로 변경
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
Dim eCode , title , eFolder , topimg , btmimg , topaddimg 'floating img
Dim slideimg, cEvtCont, videoLink, videoType, menuidx
Dim mode , idx , strSql , sqlStr, smode, saveafter
smode = requestCheckvar(request("smode"),16)
saveafter = requestCheckvar(request("saveafter"),2)
eCode = requestCheckvar(request("eC"),16)
menuidx = requestCheckvar(request("menuidx"),16)
if menuidx="" or isnull(menuidx) then menuidx=0

If saveafter <>"" Then
Response.write "<script>opener.document.location.reload();</script>"
End If

If smode="SU" Then
Response.write "<script>self.close();</script>"
Response.end
End If

IF eCode <> "" THEN
    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode	'이벤트 코드
    '이벤트 화면설정 내용 가져오기
    cEvtCont.fnGetEventDisplay
    videoLink = cEvtCont.FvideoLink
    videoType = cEvtCont.FvideoType
    set cEvtCont = nothing
End If

%>
<!-- #include virtual="/admin/lib/popheaderslide.asp"-->
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
</style>
<script type="text/javascript" src="/admin/common/lib/js/jquery.swiper-3.1.2.min.js"></script>
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
	chkAllItem();
	var chk=0;
	$("input[name='chkIdx']").each(function(){
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

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV3.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan+'&wid=750&hei=528','popImg','width=370,height=150');
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
		url: "pop_mobile_top_slide_ajax.asp",
		data: "eC=<%=eCode%>&menuidx=<%=menuidx%>",
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

function fnDeleteVideo(){
	if(confirm("동영상을 삭제 하시겠습니까?")){
		document.frmvideo.mode.value="VD";
		document.frmvideo.submit();
	}
}

function fnMasterInfoSet(){
	var winMasterInfo;
	winMasterInfo = window.open('/admin/eventmanage/event/v5/popup/pop_multicontents_masterinfo.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>','menuinfo','width=1024,height=850,scrollbars=yes,resizable=yes');
	winMasterInfo.focus();
}
</script>
</head>
<body>
<div class="slideRegister adminMob">
	<h1>슬라이드 등록 (MOBILE)&nbsp;&nbsp;/&nbsp;&nbsp;<a href="/admin/eventmanage/event/v5/template/slide/pop_pcweb_top_slide.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>" style="color:#FFFFFF">(PC Web)</a></h1>
	<%'슬라이드 화면 불러오기%>
	<div class="preview <% If topaddimg <> "" Then %>txtFix<% End If %>" id="preview_ajax"></div>
	<%'슬라이드 화면 불러오기%>
	<div class="register">
		<h2>컨텐츠 등록</h2>
		<dl>
			<dt>- 슬라이드 <span>(필수/3~10개까지 등록)</span> <input type="button" class="btn" value="배경/레이아웃 설정" onclick="fnMasterInfoSet();"></dt>
			<dd>
				<div class="insertImg">
					<h3>슬라이드 이미지 등록<br/><span style="color:#c80a0a;line-height:2;">※ 상하단 이미지 우선 등록후 슬라이드 이미지를 등록해주세요. 또는 슬라이드 우선 등록후 상하단 이미지를 등록 해주세요. ※<br/>※ 등록버튼 클릭후 좌측 Viewer에 반영 됩니다.※<br/>※ 이미지 사이즈 : 750 * 528※</span></h3><br>
					<table class="tbType1 listTb tMar10" id="utypeI" style="display:<% if videoLink<>"" then %>none<% end if %>">
						<colgroup>
							<col width="50%" /><col /><col width="50%" /><col />
						</colgroup>
						<tbody>
						<tr>
							<td>
								<div id="spanslideimg"></div>
								<input class="button" type="button" value="이미지 불러오기" name="mslideimg" onClick="jsSetImg('<%=eFolder%>','','slideimg','spanslideimg');"/>
							</td>
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
				<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
				<div class="tMar20">
					<p>
						<input type="button" class="btn" value="전체 선택" onclick="chkAllItem();">
					</p>
					<table class="tbType1 listTb tMar10">
						<colgroup>
							<col width="8%" /><col /><col width="42%" /><col width="10%" /><col width="40%" />
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

							sqlStr = "SELECT idx , slideimg , linkurl , sorting , isusing " & vbcrlf
							sqlStr = sqlStr & " from db_event.[dbo].[tbl_event_top_slide_addimage] where evt_code = '"& eCode &"'" & vbcrlf
							sqlStr = sqlStr & " and device = 'M' and menuidx=" & menuidx & vbcrlf
							sqlStr = sqlStr & " order by sorting asc , idx asc " 
							rsget.Open sqlStr,dbget,1
							if Not(rsget.EOF or rsget.BOF) Then
								Do Until rsget.eof
						%>
						<tr class="<%=chkIIF(rsget("isusing")="N","bgGry1","")%>">
							<td><input type="checkbox" name="chkIdx" id="chkIdx" value="<%=rsget("idx")%>" /></td>
							<td><img src="<%=rsget("slideimg")%>" style="width:100px;" /></td>
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
				</div>
				</form>
			</dd>
		</dl>
		<div class="btnArea">
			<input type="image" src="/images/icon_save.gif" alt="저장" onclick="saveList();"/>
			<a href="" onclick="self.close();"><img src="/images/icon_cancel.gif" alt="취소" /></a>
		</div>
	</div>
</div>
<form name="frmdel" method="POST" action="pop_top_slide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
<input type="hidden" name="mode" value="SD"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="chkIdx" />
</form>
<form name="slideimgfrm" method="post" action="pop_top_slide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
<input type="hidden" name="mode" value="SI"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="slideimg" value=""/>
<input type="hidden" name="linkurl" id="linkurl" value=""/>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->