<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_top_info.asp
' Discription : I형(통합형) 이벤트 기획전 상단타이틀 셋팅 설정 창
' History : 2019.01.28 정태훈
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
Dim cEvtCont, eFolder, winmode, sqlStr, contentsAlign
Dim eCode, etemp, etemp_mo, ehtml5_mo, emimg_mo, ehtml_mo
Dim eregdate, title_pc, title_mo, enamesub, blnFull, blnWide
Dim mdbntype, mdbntypemo, themecolor, themecolormo, subcopyK
Dim ehtml5, emimg, ehtml, blnExec, blnExec_mo, eExecFile, eExecFile_mo
dim evt_pc_addimg_cnt, evt_m_addimg_cnt, textbgcolor, GroupItemType
dim slide_w_flag, slide_m_flag, eventtype_pc, eventtype_mo, copyhide, ename
eCode = Request("eC")
winmode = Request("wm")

IF eCode <> "" THEN
    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode	'이벤트 코드
    cEvtCont.fnGetEventCont
	eregdate = cEvtCont.FERegdate
    ename = db2html(cEvtCont.FEName) ' 이벤트명
    enamesub = db2html(cEvtCont.FENamesub) '이벤트 타이틀 서브카피 모바일
    subcopyK =	db2html(cEvtCont.FsubcopyK) '서브카피 한글 PC
    '이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay
    etemp = cEvtCont.FETemp
    etemp_mo = cEvtCont.FETemp_mo
    emimg_mo = cEvtCont.FEMImg_mo
    emimg 		= cEvtCont.FEMImg
    ehtml_mo = db2html(cEvtCont.FEHtml_mo)
    ehtml 		= db2html(cEvtCont.FEHtml)
    if etemp="" or isnull(etemp) then etemp="10"
    if etemp_mo="" or isnull(etemp_mo) then etemp_mo="11"

    if etemp="3" then etemp="6"
    if etemp_mo="3" then etemp_mo="6"

    blnFull	= cEvtCont.FEFullYN
 	blnWide	= cEvtCont.FEWideYN
    mdbntype = cEvtCont.Fmdbntype
	mdbntypemo = cEvtCont.Fmdbntypemo
    themecolor = cEvtCont.Fthemecolor
	themecolormo = cEvtCont.Fthemecolormo
    textbgcolor = cEvtCont.Ftextbgcolor
    blnExec = cEvtCont.FEisExec
    blnExec_mo = cEvtCont.FEisExec_mo
    evt_pc_addimg_cnt = cEvtCont.FEvt_pc_addimg_cnt '// PC 추가 이미지 카운트
	evt_m_addimg_cnt = cEvtCont.FEvt_m_addimg_cnt '// 모바일 추가 이미지 카운트
    copyhide = cEvtCont.FvideoType  '//모바일 카피, 서브카피 숨김 유무
	'엠디 등록 이벤트 테마 정보
	cEvtCont.fnGetEventMDThemeInfo
    title_pc = cEvtCont.Ftitle_pc
	title_mo = cEvtCont.Ftitle_mo
    eExecFile = cEvtCont.FEexecFile
    eExecFile_mo = cEvtCont.FEexecFile_mo
    GroupItemType = cEvtCont.FGroupItemType
    contentsAlign = cEvtCont.FcontentsAlign
    slide_w_flag	= cEvtCont.FESlide_W_Flag '// 슬라이드 웹
	slide_m_flag	= cEvtCont.FESlide_M_Flag '// 슬라이드 모바일
	eventtype_pc = cEvtCont.Feventtype_pc
	eventtype_mo = cEvtCont.Feventtype_mo
    set cEvtCont = nothing

    if mdbntypemo="" or isnull(mdbntypemo) then mdbntypemo="D"
    If themecolor = "" or isnull(themecolor) Then themecolor=11
	If themecolormo = "" or isnull(themecolormo) Then themecolormo=11
    if mdbntype="" or isnull(mdbntype) then mdbntype="D"
    If textbgcolor = "" Then textbgcolor=1
    if title_mo="" then title_mo=ename
    if title_pc="" then title_pc=ename
else

end if

if GroupItemType=" " or GroupItemType="" or isnull(GroupItemType) then GroupItemType="C"
if contentsAlign=" " then contentsAlign=1

eFolder = eCode
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){
    if($("input:radio[name='evt_template_mo']:checked").val()==11){
        if(!frm.title_mo.value){
            alert("모바일 메인카피를 입력해주세요");
            frm.title_mo.focus();
            return false;
        }
        if(!frm.subsEN.value){
            alert("모바일 서브카피를 입력해주세요");
            frm.subsEN.focus();
            return false;
        }

        if(!frm.title_pc.value){
            alert("PC 메인카피를 입력해주세요");
            frm.title_pc.focus();
            return false;
        }
        if(!frm.subcopyK.value){
            alert("PC 서브카피를 입력해주세요");
            frm.subcopyK.focus();
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

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		eval("document.all."+sName).value = "";
		eval("document.all."+sSpan).style.display = "none";
	}
}

function jsAddByte(target,obj){ 
    var realText = obj.value; 
    var textBit = '';
    var textLen = 0;
    for (var i = 0 ; i < realText.length ; i++) {
        textBit = realText.charAt(i); 
        if(escape(textBit).length > 4) {
            textLen = textLen + 2;
        } else {
            textLen = textLen + 1;
        }

        if (textLen >= 140){
            realText = realText.substr(0,i);
            obj.value = realText;
            break;
        }
    }
	if(target=="1"){
		$("#etkm").html(textLen);
	}
	else if(target=="2"){
		$("#sckm").html(textLen);
    }
    else if(target=="3"){
		$("#etkp").html(textLen);
	}
	else{
		$("#sckp").html(textLen);
	}
}

//탑 배너 등록 체크
function TnThemeBannerRegCheck(d){
    if (d == "w"){
        var winpop = window.open('/admin/eventmanage/event/v5/template/slide/pop_pcweb_top_slide.asp?eC=<%=eCode%>&bgubun='+ $("input[name='chkWide']:checked").val() + '&mdtheme=' + document.frmEvt.mdtheme.value,'winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
        winpop.focus();
    }else{
        var winpop = window.open('/admin/eventmanage/event/v5/template/slide/pop_mobile_top_slide.asp?eC=<%=eCode%>&bgubun='+document.frmEvt.themetypemo.value,'winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
        winpop.focus();
    }
}

//색상코드 선택
function selColorChipMo(cd) {
    var i;
    document.frmEvt.DFcolorCDMo.value=cd;
    for(i=1;i<=27;i++) {
        $("#clinet"+i).removeClass("picked");
    }
    $("#clinet"+cd).addClass("picked");
}

//색상코드 선택
function selColorChip(cd) {
    var i;
    document.frmEvt.DFcolorCD.value=cd;
    for(i=1;i<=27;i++) {
        $("#pclinet"+i).removeClass("picked");
    }
    $("#pclinet"+cd).addClass("picked");
}

//색상코드 선택
function selColorChip2(cd) {
    var i;
    document.frmEvt.DFcolorCD.value=cd;
    for(i=1;i<=27;i++) {
        $("#pcclinet"+i).removeClass("picked");
    }
    $("#pcclinet"+cd).addClass("picked");
}

function fnChangeTemplate(div){
    if(div==0){
        $("#topmdiv1").show();
        $("#topmdiv2").show();
        $("#topmdiv3").hide();
        $("#topmdiv11").hide();
        $("#topmdiv12").show();
    }
    else if(div==1){
        $("#topmdiv1").hide();
        $("#topmdiv2").hide();
        $("#topmdiv3").hide();
        $("#topmdiv11").hide();
        $("#topmdiv12").hide();
    }
    else if(div==4){
        $("#topmdiv1").hide();
        $("#topmdiv2").hide();
        $("#topmdiv3").hide();
        $("#topmdiv11").show();
        $("#topmdiv12").hide();
    }
}

function fnChangeDivece(div){
    if(div=="M"){
        $("#deviceM").show();
        $("#deviceP").hide();
    }
    else if(div=="P"){
        $("#deviceM").hide();
        $("#deviceP").show();
    }
}
$(document).ready(function(){
    $('#title_mo').keyup(function(){
        $("#title_pc").val($("#title_mo").val());
    });
    $('#subsEN').keyup(function(){
        $("#subcopyK").val($("#subsEN").val());
    });
});

function TnDelSlideBanner(idx){
    document.ibfrm.target="ifrmProc";
    document.ibfrm.idx.value=idx;
    document.ibfrm.action="/admin/eventmanage/event/v5/template/slide/delslidebanner.asp";
    document.ibfrm.submit();
}

function poppcaddimg(){
    var winPopaddimg;
    winPopaddimg = window.open('/admin/eventmanage/event/v5/template/addbanner/pop_pc_addbanner.asp?eC=<%=eCode%>','pCal','width=1450,height=800,scrollbars=yes,resizable=yes');
    winPopaddimg.focus();
}

function popmoaddimg(){
    var winPopaddimg;
    winPopaddimg = window.open('/admin/eventmanage/event/v5/template/addbanner/pop_mobile_addbanner.asp?eC=<%=eCode%>','pCal','width=1450,height=800,scrollbars=yes,resizable=yes');
    winPopaddimg.focus();
}

//수작업 슬라이드 이미지 등록
function TnSlideImageRegCheck(d){
    if (d=="W"){
        var winslidepop = window.open('/admin/eventmanage/event/v5/template/slide/pop_pcweb_slide.asp?eC=<%=eCode%>&bgubun=3','winslidepop','width=1450,height=800,scrollbars=yes,resizable=yes');
        winslidepop.focus();
    }else{
        var winslidepop = window.open('/admin/eventmanage/event/v5/template/slide/pop_mobile_slide.asp?eC=<%=eCode%>','winslidepop','width=1450,height=800,scrollbars=yes,resizable=yes');
        winslidepop.focus();
    }
}
//수작업 슬라이드 이미지 삭제
function TnDelSlideImage(idx){
    document.ibfrm.target="ifrmProc";
    document.ibfrm.idx.value=idx;
    document.ibfrm.action="/admin/eventmanage/event/v5/template/slide/delslideimage.asp";
    document.ibfrm.submit();
}

//20181105 멀티3번 최종원
function pop_multi3_manage(){	
	var multi3Window = window.open('/admin/eventmanage/event/V5/popup/pop_manage_multi3.asp?evt_code=<%=eCode%>','multi3Window','width=700, height=900,scrollbars=yes,resizable=yes');
	multi3Window.focus();
}

function jsSetVideo(sFolder, sImg, sName, sSpan){
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName,'popImg','width=370,height=150');
	winImg.focus();
}

//색상코드 선택
function selBackGroundColorChip(cd) {
    var i;
    frmEvt.DFcolorCD2.value= cd;
    for(i=1;i<=2;i++) {
        $("#bcline"+i).removeClass("picked");
    }
    $("#bcline"+cd).addClass("picked");
}

function jsManageEventImageNew(evtcode){
    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/V2/eventManageDir_new.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function fnAlignTypeChange(objvalue){
    if(objvalue==1){
        $('input:radio[name=chkFull]:input[value=1]').attr("checked", true);
        $('input:radio[name=chkWide]:input[value=1]').attr("checked", false);
    }
    else{
        $('input:radio[name=chkFull]:input[value=1]').attr("checked", false);
        $('input:radio[name=chkWide]:input[value=1]').attr("checked", true);
    }
}
function fnFullWideCheck(objvalue){
    if(objvalue==1){
        $('input:radio[name=contentsAlign]:input[value=1]').attr("checked", true);
        $('input:radio[name=contentsAlign]:input[value=2]').attr("checked", false);
    }
    else{
        $('input:radio[name=contentsAlign]:input[value=1]').attr("checked", false);
        $('input:radio[name=contentsAlign]:input[value=2]').attr("checked", true);
    }
}

</script>
<form name="frmEvt" method="post" style="margin:0px;" action="topinfo_process.asp">
<input type="hidden" name="imod" value="TU">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="themetypemo" value="3">
<input type="hidden" name="themetype" value="2">
<input type="hidden" name="mdtheme" value="2">
<input type="hidden" name="DFcolorCD" value="<%=themecolor%>">
<input type="hidden" name="DFcolorCDMo" value="<%=themecolormo%>">
<input type="hidden" name="DFcolorCD2" value="<%=textbgcolor%>">
<input type="hidden" name="mdbntypemo" value="D">
<input type="hidden" name="slide_w_flag" value="N">
<input type="hidden" name="slide_m_flag" value="N">
<input type="hidden" name="evt_template" value="10">
<div class="popV19">
	<div class="popHeadV19">
		<h1>상단타이틀</h1>
	</div>
	<div class="popContV19" id="deviceM" style="display:<% if winmode="P" then %>none<% end if %>">
		<div class="tabV19">
			<ul>
				<li class="selected"><a href="#">Mobile / App</a></li>
				<li class="" onClick="fnChangeDivece('P');"><a href="#">PC</a></li>
			</ul>
		</div>
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th>입력방법</th>
                    <td>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="evt_template_mo" id="radio1a" value="11" onclick="fnChangeTemplate(0);"<% if etemp_mo="11" then %> checked<% end if %>>
                                템플릿 등록
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="evt_template_mo" id="radio1c" value="10" onclick="fnChangeTemplate(4);"<% if etemp_mo="10" then %> checked<% end if %>>
                                Multi3형
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>유형선택</th>
                    <td>
						<select class="formControl" name="mo_evttype">
                            <option value=""<% If eventtype_mo="" Then Response.write " selected" %>>선택</option>
                            <option value="0"<% If eventtype_mo="0" or eventtype_mo="" Then Response.write " selected" %>>MD형</option>
                            <option value="20"<% If eventtype_mo="20" Then Response.write " selected" %>>디자인형</option>
						</select>
                    </td>
                </tr>
                <!-- '템플릿 등록' 선택시 노출-->
				<tr id="topmdiv1" style="display:<% if etemp_mo="6" or etemp_mo="10" then %>none<% end if %>">
                    <th>메인카피 (기획전명)</th>
                    <td>
                        <input type="text" class="formControl formControl550" placeholder="메인카피" name="title_mo" id="title_mo" value="<%=title_mo%>" OnKeyUp="jsAddByte('1',this);">
                         <span class="lMar05 cGy1 fs12 vBtm"><span class="cPk2 vBtm" id="etkm">50</span><span class="cPk2 vBtm">byte</span>&#47;120byte</span>
                        <script type="text/javascript">
                            jsAddByte(1,frmEvt.title_mo);
                        </script>
                    </td>
                </tr>
                <tr id="topmdiv2" style="display:<% if etemp_mo="6" or etemp_mo="10" then %>none<% end if %>">
                    <th>서브카피</th>
                    <td>
                        <input type="text" class="formControl formControl550" placeholder="서브카피" name="subsEN" id="subsEN" value="<%=enamesub%>" OnKeyUp="jsAddByte('2',this);">
                        <span class="lMar05 cGy1 fs12 vBtm"><span class="cPk2 vBtm" id="sckm">50</span><span class="cPk2 vBtm">byte</span>&#47;120byte</span>
                        <script type="text/javascript">
                            jsAddByte(2,frmEvt.subsEN);
                        </script>
                    </td>
                </tr>
                <tr id="topmdiv13" style="display:<% if etemp_mo="10" then %>none<% end if %>">
                    <th>템플릿 카피 여부</th>
                    <td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="copyhide" value="1"<% if copyhide="1" then %> checked<% end if %>>
								카피 / 서브카피 숨김
								<i class="inputHelper"></i>
							</label>
						</div>
                    </td>
                </tr>
                <tr>
                    <th>이미지</th>
                    <td>
                        <div>
                            <div class="colorPicker">
                                <ul>
                                    <li<% If themecolormo="11" Or themecolormo="" Then %> class="picked"<% End If %> onClick="selColorChipMo(11);" id="clinet11"><span style="background-color:#848484;"></span></li>
                                    <li<% If themecolormo="1" Then %> class="picked"<% End If %> onClick="selColorChipMo(1);" id="clinet1"><span style="background-color:#ed6c6c;"></span></li>
                                    <li<% If themecolormo="2" Then %> class="picked"<% End If %> onClick="selColorChipMo(2);" id="clinet2"><span style="background-color:#f385af;"></span></li>
                                    <li<% If themecolormo="3" Then %> class="picked"<% End If %> onClick="selColorChipMo(3);" id="clinet3"><span style="background-color:#f3a056;"></span></li>
                                    <li<% If themecolormo="4" Then %> class="picked"<% End If %> onClick="selColorChipMo(4);" id="clinet4"><span style="background-color:#e7b93c;"></span></li>
                                    <li<% If themecolormo="5" Then %> class="picked"<% End If %> onClick="selColorChipMo(5);" id="clinet5"><span style="background-color:#8eba4a;"></span></li>
                                    <li<% If themecolormo="6" Then %> class="picked"<% End If %> onClick="selColorChipMo(6);" id="clinet6"><span style="background-color:#43a251;"></span></li>
                                    <li<% If themecolormo="7" Then %> class="picked"<% End If %> onClick="selColorChipMo(7);" id="clinet7"><span style="background-color:#50bdd1;"></span></li>
                                    <li<% If themecolormo="8" Then %> class="picked"<% End If %> onClick="selColorChipMo(8);" id="clinet8"><span style="background-color:#5aa5ea;"></span></li>
                                    <li<% If themecolormo="9" Then %> class="picked"<% End If %> onClick="selColorChipMo(9);" id="clinet9"><span style="background-color:#2672bf;"></span></li>
                                    <li<% If themecolormo="10" Then %> class="picked"<% End If %> onClick="selColorChipMo(10);" id="clinet10"><span style="background-color:#2c5a85;"></span></li>
                                </ul>
                                <ul class="tMar05">
                                    <li<% If themecolormo="12" Then %> class="picked"<% End If %> onClick="selColorChipMo(12);" id="clinet12"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_1.jpg);"></span></li>
                                    <li<% If themecolormo="13" Then %> class="picked"<% End If %> onClick="selColorChipMo(13);" id="clinet13"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_2.jpg);"></span></li>
                                    <li<% If themecolormo="14" Then %> class="picked"<% End If %> onClick="selColorChipMo(14);" id="clinet14"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_3.jpg);"></span></li>
                                    <li<% If themecolormo="15" Then %> class="picked"<% End If %> onClick="selColorChipMo(15);" id="clinet15"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_4.jpg);"></span></li>
                                    <li<% If themecolormo="16" Then %> class="picked"<% End If %> onClick="selColorChipMo(16);" id="clinet16"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_5.jpg);"></span></li>
                                    <li<% If themecolormo="17" Then %> class="picked"<% End If %> onClick="selColorChipMo(17);" id="clinet17"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_6.jpg);"></span></li>
                                    <li<% If themecolormo="18" Then %> class="picked"<% End If %> onClick="selColorChipMo(18);" id="clinet18"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_7.jpg);"></span></li>
                                    <li<% If themecolormo="19" Then %> class="picked"<% End If %> onClick="selColorChipMo(19);" id="clinet19"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_8.jpg);"></span></li>
                                    <li<% If themecolormo="20" Then %> class="picked"<% End If %> onClick="selColorChipMo(20);" id="clinet20"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_9.jpg);"></span></li>
                                    <li<% If themecolormo="21" Then %> class="picked"<% End If %> onClick="selColorChipMo(21);" id="clinet21"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_10.jpg);"></span></li>
                                    <li<% If themecolormo="22" Then %> class="picked"<% End If %> onClick="selColorChipMo(22);" id="clinet22"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_11.jpg);"></span></li>
                                    <li<% If themecolormo="23" Then %> class="picked"<% End If %> onClick="selColorChipMo(23);" id="clinet23"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_12.jpg);"></span></li>
                                    <li<% If themecolormo="24" Then %> class="picked"<% End If %> onClick="selColorChipMo(24);" id="clinet24"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_13.jpg);"></span></li>
                                    <li<% If themecolormo="25" Then %> class="picked"<% End If %> onClick="selColorChipMo(25);" id="clinet25"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_14.jpg);"></span></li>
                                    <li<% If themecolormo="26" Then %> class="picked"<% End If %> onClick="selColorChipMo(26);" id="clinet26"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_15.jpg);"></span></li>
                                    <li<% If themecolormo="27" Then %> class="picked"<% End If %> onClick="selColorChipMo(27);" id="clinet27"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_16.jpg);"></span></li>
                                </ul>
                            </div>
                            <p class="tMar05 cGy1 fs12">*선택하신 테마 정보로 배경과 기차 색상이 선택됩니다.</p>
                        </div>
                    </td>
                </tr>
                <tr id="topmdiv8">
                    <th>Exec File (개발파일)</th>
                    <td>
                        <input type="text" class="formControl" placeholder="개발파일" name="sEFP_mo" value="<%=eExecFile_mo%>">
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="rdoEF_mo" id="radio4a" value="0"<%if not blnExec_mo then%> checked<%end if%>>
                                비실행
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="rdoEF_mo" id="radio4a" value="1"<%if blnExec_mo then%>checked<%end if%>>
                                실행
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                    </td>
                </tr>
                <!--// '수작업 등록' 선택시 노출 -->

                <!-- 'Multi3형' 선택시 노출 -->
                <tr id="topmdiv11" style="display:<% if etemp_mo="6" or etemp_mo="11" then %>none<% end if %>">
                    <th>내용</th>
                    <td>
                        <button class="btn4 btnBlue1" onClick="pop_multi3_manage();return false;">내용 설정</button>
                    </td>
                </tr>
                <!--// 'Multi3형' 선택시 노출 -->

			</tbody>
        </table>
	</div>
	<div class="popContV19" id="deviceP" style="display:<% if winmode="M" then %>none<% end if %>">
		<div class="tabV19">
			<ul>
				<li class="" onClick="fnChangeDivece('M');"><a href="#">Mobile / App</a></li>
				<li class="selected"><a href="#">PC</a></li>
			</ul>
		</div>
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th>유형선택</th>
                    <td>
						<select class="formControl" name="pc_evttype">
                            <option value=""<% If eventtype_pc="" Then Response.write " selected" %>>선택</option>
                            <option value="0"<% If eventtype_pc="0" Then Response.write " selected" %>>MD형</option>
                            <option value="50"<% If eventtype_pc="50" Then Response.write " selected" %>>디자인형(와이드)</option>
                            <option value="20"<% If eventtype_pc="20" Then Response.write " selected" %>>디자인형(풀)</option>
						</select>
                    </td>
                </tr>
                <!-- '템플릿 등록' 선택시 노출-->
				<tr>
					<th>메인카피 (기획전명)</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="메인카피" name="title_pc" id="title_pc" value="<%=title_pc%>"  OnKeyUp="jsAddByte(3,this);">
                        <span class="lMar05 cGy1 fs12 vBtm"><span class="cPk2 vBtm" id="etkp">50</span><span class="cPk2 vBtm">byte</span>&#47;120byte</span>
                        <script type="text/javascript">
                            jsAddByte(3,frmEvt.title_pc);
                        </script>
					</td>
				</tr>
                <tr>
                    <th>서브카피</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="서브카피" name="subcopyK" id="subcopyK" value="<%=subcopyK%>"  OnKeyUp="jsAddByte(4,this);">
						<span class="lMar05 cGy1 fs12 vBtm"><span class="cPk2 vBtm" id="sckp">50</span><span class="cPk2 vBtm">byte</span>&#47;120byte</span>
                        <script type="text/javascript">
                            jsAddByte(4,frmEvt.subcopyK);
                        </script>
					</td>
				</tr>
                <tr>
                    <th>타이틀 정렬</th>
                    <td>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="GroupItemType" value="L"<% if GroupItemType="L" or GroupItemType="" then %> checked<% end if %>>
                                좌측
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="GroupItemType" value="C"<% if GroupItemType="C" then %> checked<% end if %>>
                                중앙
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>이미지</th>
                    <td>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="contentsAlign" value="1"<% if contentsAlign="1" or contentsAlign="" or isnull(contentsAlign) then response.write " checked"%> onclick="fnAlignTypeChange(this.value);">
                                    Full (1140 x 540px)
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="contentsAlign" value="2"<% if contentsAlign="2" then response.write " checked"%> onclick="fnAlignTypeChange(this.value);">
                                    Wide (1920 x 540px)
                                    <i class="inputHelper"></i>
                                </label>
                            </div>

                        <div class="tMar15 tPad15 topLineGrey1" id="toppcdiv6" style="display:<% if etemp="6" then %>none<% end if %>">
                            <div class="colorPicker">
                                <ul>
                                    <li<% If themecolor="11" Or themecolor="" Then %> class="picked"<% End If %> onClick="selColorChip(11);" id="pclinet11"><span style="background-color:#848484;"></span></li>
                                    <li<% If themecolor="1" Then %> class="picked"<% End If %> onClick="selColorChip(1);" id="pclinet1"><span style="background-color:#ed6c6c;"></span></li>
                                    <li<% If themecolor="2" Then %> class="picked"<% End If %> onClick="selColorChip(2);" id="pclinet2"><span style="background-color:#f385af;"></span></li>
                                    <li<% If themecolor="3" Then %> class="picked"<% End If %> onClick="selColorChip(3);" id="pclinet3"><span style="background-color:#f3a056;"></span></li>
                                    <li<% If themecolor="4" Then %> class="picked"<% End If %> onClick="selColorChip(4);" id="pclinet4"><span style="background-color:#e7b93c;"></span></li>
                                    <li<% If themecolor="5" Then %> class="picked"<% End If %> onClick="selColorChip(5);" id="pclinet5"><span style="background-color:#8eba4a;"></span></li>
                                    <li<% If themecolor="6" Then %> class="picked"<% End If %> onClick="selColorChip(6);" id="pclinet6"><span style="background-color:#43a251;"></span></li>
                                    <li<% If themecolor="7" Then %> class="picked"<% End If %> onClick="selColorChip(7);" id="pclinet7"><span style="background-color:#50bdd1;"></span></li>
                                    <li<% If themecolor="8" Then %> class="picked"<% End If %> onClick="selColorChip(8);" id="pclinet8"><span style="background-color:#5aa5ea;"></span></li>
                                    <li<% If themecolor="9" Then %> class="picked"<% End If %> onClick="selColorChip(9);" id="pclinet9"><span style="background-color:#2672bf;"></span></li>
                                    <li<% If themecolor="10" Then %> class="picked"<% End If %> onClick="selColorChip(10);" id="pclinet10"><span style="background-color:#2c5a85;"></span></li>
                                </ul>
                                <ul class="tMar05">
                                    <li<% If themecolor="12" Then %> class="picked"<% End If %> onClick="selColorChip(12);" id="pclinet12"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_1.jpg);"></span></li>
                                    <li<% If themecolor="13" Then %> class="picked"<% End If %> onClick="selColorChip(13);" id="pclinet13"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_2.jpg);"></span></li>
                                    <li<% If themecolor="14" Then %> class="picked"<% End If %> onClick="selColorChip(14);" id="pclinet14"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_3.jpg);"></span></li>
                                    <li<% If themecolor="15" Then %> class="picked"<% End If %> onClick="selColorChip(15);" id="pclinet15"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_4.jpg);"></span></li>
                                    <li<% If themecolor="16" Then %> class="picked"<% End If %> onClick="selColorChip(16);" id="pclinet16"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_5.jpg);"></span></li>
                                    <li<% If themecolor="17" Then %> class="picked"<% End If %> onClick="selColorChip(17);" id="pclinet17"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_6.jpg);"></span></li>
                                    <li<% If themecolor="18" Then %> class="picked"<% End If %> onClick="selColorChip(18);" id="pclinet18"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_7.jpg);"></span></li>
                                    <li<% If themecolor="19" Then %> class="picked"<% End If %> onClick="selColorChip(19);" id="pclinet19"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_8.jpg);"></span></li>
                                    <li<% If themecolor="20" Then %> class="picked"<% End If %> onClick="selColorChip(20);" id="pclinet20"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_9.jpg);"></span></li>
                                    <li<% If themecolor="21" Then %> class="picked"<% End If %> onClick="selColorChip(21);" id="pclinet21"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_10.jpg);"></span></li>
                                    <li<% If themecolor="22" Then %> class="picked"<% End If %> onClick="selColorChip(22);" id="pclinet22"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_11.jpg);"></span></li>
                                    <li<% If themecolor="23" Then %> class="picked"<% End If %> onClick="selColorChip(23);" id="pclinet23"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_12.jpg);"></span></li>
                                    <li<% If themecolor="24" Then %> class="picked"<% End If %> onClick="selColorChip(24);" id="pclinet24"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_13.jpg);"></span></li>
                                    <li<% If themecolor="25" Then %> class="picked"<% End If %> onClick="selColorChip(25);" id="pclinet25"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_14.jpg);"></span></li>
                                    <li<% If themecolor="26" Then %> class="picked"<% End If %> onClick="selColorChip(26);" id="pclinet26"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_15.jpg);"></span></li>
                                    <li<% If themecolor="27" Then %> class="picked"<% End If %> onClick="selColorChip(27);" id="pclinet27"><span style="background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_16.jpg);"></span></li>
                                </ul>
                            </div>
                            <p class="tMar05 cGy1 fs12">*선택하신 테마 정보로 배경과 기차 색상이 선택됩니다.</p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>텍스트 색 지정</th>
                    <td>
                        <div class="colorPicker">
                            <ul>
                                <li<% If textbgcolor="1" Or textbgcolor="" Then %> class="picked"<% End If %> onClick="selBackGroundColorChip(1);" id="bcline1"><span style="background-color:#ffffff;"></span></li>
                                <li<% If themecolor="2" Then %> class="picked"<% End If %> onClick="selBackGroundColorChip(2);" id="bcline2"><span style="background-color:#000000;"></span></li>
                            </ul>
                        </div>
                        <p class="tMar05 cGy1 fs12">*이미지 중앙의 카피 문구 배경에 불투명도 처리할 배경색을 선택해주세요.</p>
                    </td>
                </tr>
                <tr>
                    <th>Exec File (개발파일)</th>
                    <td>
                        <input type="text" class="formControl" placeholder="개발파일" name="sEFP" value="<%=eExecFile%>">
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="rdoEF" value="0" <%if not blnExec then%>checked<%end if%>>
                                비실행
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="rdoEF" value="1" <%if blnExec then%>checked<%end if%>>
                                실행
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                    </td>
                </tr>
                <!--// '수작업 등록' 선택시 노출 -->
			</tbody>
        </table>
    </div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">저장</button>
	</div>
</div>
</form>
<form method="post" name="ibfrm">
	<input type="hidden" name="idx">
	<input type="hidden" name="mode">
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->