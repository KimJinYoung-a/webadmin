<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_listbanner.asp
' Discription : I형(통합형) 이벤트 기본 배너 설정 창
' History : 2019.01.24 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
Dim cEvtCont, eFolder, eregdate
Dim eCode, eEtcitemid, eEtcitemimg, ebimgMo2014


eCode = Request("eC")

IF eCode <> "" THEN
    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode	'이벤트 코드
	cEvtCont.fnGetEventCont
	eregdate = cEvtCont.FERegdate
    '이벤트 화면설정 내용 가져오기
    cEvtCont.fnGetEventDisplay
	eEtcitemid = cEvtCont.FEtcitemid
	eEtcitemimg	= cEvtCont.FEtcitemimg
	ebimgMo2014	= cEvtCont.FEBImgMoListBanner '//2014 모바일 리스트 배너 추가
	
    set cEvtCont = nothing
else
    eEtcitemid=""
    eEtcitemimg=""
end if 

eFolder = eCode
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){

	if(frm.etcitemban.value!=""){
		if(frm.banMoList.value==""){
			if(confirm("와이드 배너 등록 없이 저장 하시겠습니까?")){
				frm.submit();
			}
			return false;
		}
	}
	if(frm.banMoList.value!=""){
		if(frm.etcitemban.value==""){
			if(confirm("기본 배너 등록 없이 저장 하시겠습니까?")){
				frm.submit();
			}
			return false;
		}
	}
	frm.submit();
}

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/event/v5/lib/pop_event_uploadimg.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsShowImage(sImg){ 
	var winshowImg;
	winshowImg = window.open('/lib/showimage.asp?img='+sImg,'popImg','width=370,height=150');
	winshowImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		eval("document.all."+sName).value = "";
		eval("document.all."+sSpan).style.display = "none";
	}
}

</script>
<form name="frmEvt" method="post" style="margin:0px;" action="listbanner_process.asp">
<input type="hidden" name="imod" value="BU">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="etcitemban" value="<%=eEtcitemimg%>">
<input type="hidden" name="banMoList" value="<%=ebimgMo2014%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>기본배너</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>정방형<p class="cGy2 fs13">(750px × 750px)</p></th>
					<td>
						<button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=eEtcitemimg%>','etcitemban','etciitem');return false;">이미지 찾기</button>
						<%IF eEtcitemimg <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('etcitemban','etciitem');return false;">삭제</button><%END IF%>
						<div class="previewThumb150W tMar10" id="etciitem"><%IF eEtcitemimg <> "" THEN %><a href="javascript:jsShowImage('<%=eEtcitemimg%>')"><img src="<%=eEtcitemimg%>" alt=""></a><%END IF%></div>
					</td>
				</tr>
				<tr>
					<th>와이드<p class="cGy2 fs13">(750px × 512px)</p></th>
					<td>
						<button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=ebimgMo2014%>','banMoList','spanbanMoList');return false;">이미지 찾기</button>
						<%IF ebimgMo2014 <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('banMoList','spanbanMoList');return false;">삭제</button><%END IF%>
						<div class="previewThumb150H tMar10" id="spanbanMoList"><%IF ebimgMo2014 <> "" THEN %><a href="javascript:jsShowImage('<%=ebimgMo2014%>')"><img src="<%=ebimgMo2014%>" alt=""></a><%END IF%></div>
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">저장</button>
	</div>
</div>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->