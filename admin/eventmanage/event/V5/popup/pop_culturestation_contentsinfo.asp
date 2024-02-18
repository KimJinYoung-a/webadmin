<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_culturestation_contentsinfo.asp
' Discription : 컬쳐스테이션 컨텐츠 설정 창
' History : 2019.02.20 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/divisionCls.asp"-->
<%
Dim cEvtCont, eFolder, sqlStr, evt_type, evt_kind, themecolor, CommentTitle
Dim eCode, ename, eregdate, enamesub, subcopyK, etemp, ehtml5, ehtml5_mo

eCode = Request("eC")

IF eCode <> "" THEN
    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode	'이벤트 코드
    cEvtCont.fnGetEventCont
    ename =	db2html(cEvtCont.FEName)
	eregdate = cEvtCont.FERegdate
    enamesub = db2html(cEvtCont.FENamesub) '이벤트 타이틀 서브카피 모바일
    subcopyK =	db2html(cEvtCont.FsubcopyK) '서브카피 한글 PC
    '이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay
    etemp = cEvtCont.FEMImg
    ehtml5 = nl2blank(cEvtCont.FEHtml)
    ehtml5_mo = nl2blank(cEvtCont.FEHtml_mo)
    themecolor = cEvtCont.Fthemecolor
    evt_type = cEvtCont.Feventtype_pc
    evt_kind = cEvtCont.Feventtype_mo
    CommentTitle = cEvtCont.FECommentTitle
    set cEvtCont = nothing
    If themecolor = "" or isnull(themecolor) Then themecolor=11

end if
eFolder = eCode
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){

    if(!frm.sEN.value){
        alert("메인카피를 입력해주세요");
        frm.sEN.focus();
        return false;
    }
    if(!frm.subcopyK.value){
        alert("서브카피를 입력해주세요");
        frm.subcopyK.focus();
        return false;
    }
    var content = Editor.getContent();
    //alert(content);
    Editor.switchEditor("2");
    var content2 = Editor.getContent();
    //alert(content2);
    document.getElementById("m_cmt_desc").value = content;
    document.getElementById("m_main_content").value = content2;
	frm.submit();
}

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/event/v5/lib/pop_event_uploadimgV5.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		eval("document.all."+sName).value = "";
		eval("document.all."+sSpan).style.display = "none";
	}
}

function jsAddByte(obj,target){ 
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
    document.frmEvt.subSize.value = textLen;
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
</script>

<!-- ckeditor ------------------------->
<script type='text/javascript' src="/js/ckeditor/ckeditor.js"></script>
<!-- //ckeditor ------------------------->
<form name="frmEvt" method="post" style="margin:0px;" action="culturestation_process.asp">
<input type="hidden" name="imod" value="CU">
<input type="hidden" name="DFcolorCD" value="<%=themecolor%>">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>상단타이틀</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>진행업체</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="입력하세요." name="evt_comment" maxlength="120" value="<%=CommentTitle%>">
					</td>
				</tr>
                <tr>
                    <th>포스터<p class="cGy2 fs13">책 표지 등 대표 이미지</p></th>
                    <td>
                        <input type="hidden" name="evt_mainimg" value="<%=etemp%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=etemp%>','evt_mainimg','mainimg');return false;">이미지 찾기</button>
                        <%IF etemp <> "" THEN %>
                        <button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('evt_mainimg','mainimg');return false;">삭제</button>
                        <%END IF%>
                        <div class="previewThumb150W tMar10" id="mainimg">
                            <%IF etemp <> "" THEN %>
                            <img src="<%=etemp%>" alt="">
                            <%END IF%>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
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
                            <p class="tMar05 cGy1 fs12">*선택하신 테마 정보로 배경 색상이 선택됩니다.</p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>모바일 컨텐츠 정보입력</th>
                    <td>
                        <textarea name="maincontent" rows="15" style="width:100%"><%=ehtml5 %></textarea>
                        <script>
                        //
                        window.onload = new function(){
                            var itemContEditor = CKEDITOR.replace('maincontent',{
                                height : 400,
                                // 업로드된 파일 목록
                                //filebrowserBrowseUrl : '/browser/browse.asp',
                                // 파일 업로드 처리 페이지
                                filebrowserImageUploadUrl : '<%= ItemUploadUrl %>/linkweb/event_admin/v5/eventEditorContentUpload.asp?eventid=<%=eCode%>'
                            });
                            itemContEditor.on( 'change', function( evt ) {
                                // 입력할 때 textarea 정보 갱신
                                document.frmEvt.maincontent.value = evt.editor.getData();
                            });
                        }
                        </script>
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
<form method="post" name="ibfrm">
	<input type="hidden" name="idx">
	<input type="hidden" name="mode">
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->