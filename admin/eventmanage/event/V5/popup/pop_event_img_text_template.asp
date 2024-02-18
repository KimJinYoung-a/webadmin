<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_img_text_template.asp
' Discription : I형(통합형) 이벤트 기획전 메인 이미지 텍스트 템플릿 설정 창
' History : 2019.10.02 정태훈
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
Dim cEvtCont, eFolder, winmode, eregdate, idx
Dim eCode, ImgURL, BrandContents, menuidx
Dim BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin

eCode = Request("eC")
menuidx = Request("menuidx")
winmode = Request("wm")
if winmode="" then winmode="M"
IF eCode <> "" THEN
    set cEvtCont = new ClsMultiContentsMenu
    cEvtCont.FRectEvtCode = eCode
    cEvtCont.FRectIDX = menuidx	'멀티 컨텐츠 메뉴 idx
    cEvtCont.FRectDevice = winmode
    cEvtCont.fnGetMultiContentsImgText
    idx = cEvtCont.Fidx
    ImgURL = cEvtCont.FImgURL
    BrandContents = cEvtCont.FBrandContents
    BGImage = cEvtCont.FBGImage
	BGColorLeft = cEvtCont.FBGColorLeft
    BGColorRight = cEvtCont.FBGColorRight
	contentsAlign = cEvtCont.FcontentsAlign
	Margin = cEvtCont.FMargin
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
function jsEvtSubmit(frm){
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

function jsManageEventImageNew(evtcode){
    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/V2/eventManageDir_new.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<form name="frmEvt" method="post" style="margin:0px;" action="imgtext_process.asp">
<% if idx<>"" then %>
<input type="hidden" name="imod" value="TU">
<% else %>
<input type="hidden" name="imod" value="TI">
<% end if %>
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="device" value="<%=winmode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menuidx" value="<%=menuidx%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>이미지 & HTML</h1>
	</div>
	<div class="popContV19">
		<div class="tabV19">
			<ul>
				<li class="<% if winmode="M" then %>selected<% end if %>"><a href="pop_event_img_text_template.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>&wm=M">Mobile / App</a></li>
				<li class="<% if winmode="W" then %>selected<% end if %>"><a href="pop_event_img_text_template.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>&wm=W">PC</a></li>
			</ul>
		</div>
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr id="topmdiv7">
                    <th>이미지 &#38; HTML</th>
                    <td>
                        <input type="hidden" name="main_mo" value="<%=ImgURL%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=ImgURL%>','main_mo','spanmain_mo');return false;" >메인 이미지 등록</button>
                        <input type="button" value="이미지관리"  onclick="jsManageEventImageNew('<%=eCode%>')" class="btn4 btnBlue1">
                        <%IF ImgURL <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('main_mo','spanmain_mo');return false;">삭제</button><%END IF%>
                        <div class="previewThumb150W tMar10" id="spanmain_mo">
                            <%IF ImgURL <> "" THEN %><img src="<%=ImgURL%>" alt=""><%END IF%>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <textarea name="tHtml_mo" rows="10" cols="50" placeholder="소스등록"><%=BrandContents%></textarea>
                        </div>
<% if winmode="M" then %>
<div class="wrapCode">
<p class="tMar05 cGy1 fs12">*이미지맵 코드 예시</p>
<pre><div class="codeArea"><strong>&lt;map name="Mainmap<%=menuidx%>"&gt;&lt;/map&gt;</strong></div></pre>
<p class="tMar05 cGy1 fs12">*앱 팝업 링크 코드 예시</p>
<pre><div class="codeArea"><strong>&lt;a href="" onclick="fnAPPpopupBrowserURL('마일리지 내역', 'https://m.10x10.co.kr/apps/appCom/wish/web2014/offshop/point/mileagelist.asp');return false;" class="mApp"&gt;</strong></div></pre>
<pre><div class="codeArea"><strong>&lt;a href="" onclick="fnAPPpopupBrowserURL('기획전','https://m.10x10.co.kr/apps/appCom/wish/web2014/event/eventmain.asp?eventid=111585');return false;" class="mApp"&gt;</strong></div></pre>
<pre><div class="codeArea"><strong>&lt;a href="" onclick="fnAPPpopupBrowserURL('개인정보수정',' https://m.10x10.co.kr/apps/appCom/wish/web2014/my10x10/userinfo/membermodify.asp');return false;" class="mApp"&gt;</strong></div></pre>
</div>
<% else %>
<div class="wrapCode">
<pre><div class="codeArea"><span class="cRd2">PC-WEB 예시</span>
&lt;map name="Mainmap<%=menuidx%>"&gt;
<strong>상품페이지 링크시</strong>
&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoProduct('상품번호');" onfocus="this.blur();"&gt;
<strong>이벤트페이지로 링크시</strong>
&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventMain('이벤트코드');" onfocus="this.blur();"&gt;
<strong>이벤트 그룹 페이지로 링크시</strong>
&lt;area shape="rect" coords="0,0,0,0" href="#mapGroup288144" onfocus="this.blur();"&gt;
<strong>이벤트 코멘트 이동</strong>
&lt;area shape="rect" coords="0,0,0,0" href="#commentarea" onfocus="this.blur();"&gt;
<strong>이벤트 리뷰 이동</strong>
&lt;area shape="rect" coords="0,0,0,0" href="#reviewarea" onfocus="this.blur();"&gt;
<strong>브랜드페이지 링크시</strong>
&lt;area shape="rect" coords="0,0,0,0" href="javascript:GoToBrandShop('브랜드아이디');" onfocus="this.blur();"&gt;
&lt;/map&gt;
<strong>카테고리 링크시</strong>
&lt;area shape="rect" coords="0,0,0,0" href="/shopping/category_list.asp?disp=카테고리번호" onfocus="this.blur();"&gt;
<strong>이미지 경로 http://webimage.10x10.co.kr/event/XXX/ 로 변경되었습니다.</strong>
<strong>*수작업 할인율 불러오기(예약어) : #[SALEPERCENT]</strong>
</div></pre>
</div>
<% end if %>
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
<form method="post" name="ibfrm">
	<input type="hidden" name="idx">
	<input type="hidden" name="mode">
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->