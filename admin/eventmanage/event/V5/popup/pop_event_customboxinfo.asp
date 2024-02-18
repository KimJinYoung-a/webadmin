<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
'###############################################
' PageName : pop_event_customboxinfo.asp
' Discription : I형(통합형) 이벤트 사용자 정의 박스 셋팅 설정 창
' History : 2019.02.13 정태훈
'###############################################
session.codePage = 65001		'세션코드 UTF-8 강제 설정
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
Dim cEvtCont
Dim eCode, menuidx, IDX, Title, CustomContents
Dim BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin, eFolder, eregdate

eCode = Request("eC")
menuidx = requestCheckVar(Request("menuidx"),10)

IF menuidx <> "" THEN
    set cEvtCont = new ClsMultiContentsMenu
    cEvtCont.FRectEvtCode = eCode
    cEvtCont.FRectIDX = menuidx	'멀티 컨텐츠 메뉴 idx
    cEvtCont.fnGetMultiContentsCustomBox
    IDX = cEvtCont.Fidx
    Title = cEvtCont.FTitle
    BGImage = cEvtCont.FBGImage
	BGColorLeft = cEvtCont.FBGColorLeft
    BGColorRight = cEvtCont.FBGColorRight
	contentsAlign = cEvtCont.FcontentsAlign
	Margin = cEvtCont.FMargin
    CustomContents = cEvtCont.FCustomContents
    set cEvtCont = nothing

    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode	'이벤트 코드
    cEvtCont.fnGetEventCont
	eregdate = cEvtCont.FERegdate
    if contentsAlign="" or isnull(contentsAlign) then
    cEvtCont.fnGetEventMDThemeInfo
    contentsAlign = cEvtCont.FcontentsAlign
    end if
	set cEvtCont = nothing
end if
eFolder = eCode
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
$(document).ready(function(){
    $('#nocate').on('click',function(){
        if($("#nocate").is(":checked")){
            $("#disp1").attr("disabled", true);
            $("#disp2").attr("disabled", true);
        }else{
            $("#disp1").attr("disabled", false);
            $("#disp2").attr("disabled", false);
        }
    });
});

function jsEvtSubmit(frm){

    if(!frm.title.value){
        alert("타이틀을 입력해주세요.");
        frm.title.focus();
        return false;
    }

    if(!frm.customcontents.value){
        alert("내용을 작성해주세요.");
        frm.customcontents.focus();
        return false;
    }
	frm.submit();
}

function fnSSearchBrandPop(idx){
    var wBrandView;
    wBrandView = window.open("popBrandSearch.asp?frmName=frmEvt&idx="+idx,"winBrand","width=1400,height=800,scrollbars=yes,resizable=yes");
    wBrandView.focus();
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
</script>
<script type='text/javascript' src="/js/ckeditor/ckeditor.js"></script>
<form name="frmEvt" method="post" style="margin:0px;" action="customboxinfo_process.asp">
<% if IDX<>"" then %>
<input type="hidden" name="imod" value="CU">
<% else %>
<input type="hidden" name="imod" value="CI">
<% end if %>
<input type="hidden" name="ebrand">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="idx" value="<%=IDX%>">
<input type="hidden" name="menuidx" value="<%=menuidx%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>에디터영역</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th>타이틀</th>
                    <td>
                        <input type="text" class="formControl formControl550" placeholder="입력하세요" value="<%=Title%>" name="title" id="title">
                    </td>
                </tr>
 				<tr>
					<th>내용</th>
					<td>
                        <textarea name="customcontents" rows="10" style="width:100%"><%=CustomContents %></textarea>
                        <script>
                        //
                        window.onload = new function(){
                            var itemContEditor = CKEDITOR.replace('customcontents',{
                                height : 300,
                                // 업로드된 파일 목록
                                //filebrowserBrowseUrl : '/browser/browse.asp',
                                // 파일 업로드 처리 페이지
                                filebrowserImageUploadUrl : '<%= ItemUploadUrl %>/linkweb/event_admin/v5/eventEditorContentUpload.asp?eventid=<%=eCode%>'
                            });
                            itemContEditor.on( 'change', function( evt ) {
                                // 입력할 때 textarea 정보 갱신
                                document.frmEvt.customcontents.value = evt.editor.getData();
                            });
                        }
                        </script>
					</td>
				</tr>
				<tr>
					<th>백그라운드 이미지</th>
					<td>
                        <input type="hidden" name="BGImage" value="<%=BGImage%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=BGImage%>','BGImage','spanbgimg');return false;">백그라운드 이미지 등록</button>
                        <%IF BGImage <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('BGImage','spanbgimg');return false;">삭제</button><%END IF%>
                        <div class="previewThumb150W tMar10" id="spanbgimg">
                            <%IF BGImage <> "" THEN %>
                            <%IF BGImage <> "" THEN %><img src="<%=BGImage%>" width="30%" alt=""><%END IF%>
                            <%END IF%>
                        </div>
					</td>
				</tr>
				<tr>
                    <th>백그라운드 컬러</th>
                    <td>
                        좌측 : <input type="text" class="formControl formControl150" placeholder="BackGround Color" name="BGColorLeft" id="BGColorLeft" value="<%=BGColorLeft%>">
                        우측 : <input type="text" class="formControl formControl150" placeholder="BackGround Color" name="BGColorRight" id="BGColorRight" value="<%=BGColorRight%>">
                    </td>
                </tr>
                <tr>
                    <th>유형</th>
                    <td>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="contentsAlign" value="1"<% if contentsAlign="1" or contentsAlign="" then response.write " checked"%> onclick="fnAlignTypeChange(this.value);">
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
						</div>
                    </td>
                </tr>
				<tr>
                    <th>상단 여백</th>
                    <td>
                        <div class="formInline"><input type="text" class="formControl formControl550" maxlength="6" placeholder="상단 여백" name="Margin" id="Margin" value="<%=Margin%>"> px</div>
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
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<% session.codePage = 949		'세션코드 EUC-KR 원복 %>
<!-- #include virtual="/lib/db/dbclose.asp" -->