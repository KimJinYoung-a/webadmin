<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_customgroupitem.asp
' Discription : I형(통합형) 이벤트 커스텀 상품 등롭 타입별 
' History : 2020.12.03 이종화
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
Dim cEvtCont, ix, VideoFrameReduction
Dim eCode, menuidx, GroupItemPriceView, GroupItemCheck, GroupItemType , GroupItemViewType
dim menudiv, viewsort, isusing, ArrcGroupTemplate, eFolder, eregdate
Dim BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin, textColor 
Dim device , GroupItemBrandName , GroupItemTitleName
dim saleColor , priceColor , orgpriceColor

eCode = requestCheckVar(Request("eC"),10)
menuidx = requestCheckVar(Request("menuidx"),10)
device = requestCheckVar(Request("device"),1) '// M : 모바일웹 OR W : PC 웹


eFolder = eCode
IF menuidx <> "" THEN
    set cEvtCont = new ClsMultiContentsMenu
    cEvtCont.FRectEvtCode = eCode
    cEvtCont.FRectIDX = menuidx	'멀티 컨텐츠 메뉴 idx
	cEvtCont.fnGetMultiContentsMenu
    GroupItemPriceView = cEvtCont.FGroupItemPriceView
    GroupItemCheck = cEvtCont.FGroupItemCheck
    GroupItemType = cEvtCont.FGroupItemType
    menudiv = cEvtCont.Fmenudiv
    viewsort = cEvtCont.Fviewsort
    isusing = cEvtCont.Fisusing
    BGImage = cEvtCont.FBGImage
	BGColorLeft = cEvtCont.FBGColorLeft
    BGColorRight = cEvtCont.FBGColorRight
	contentsAlign = cEvtCont.FcontentsAlign
	Margin = cEvtCont.FMargin
    textColor = cEvtCont.FtextColor
    GroupItemTitleName = cEvtCont.FGroupItemTitleName
    GroupItemViewType = cEvtCont.FGroupItemViewType
    GroupItemBrandName = cEvtCont.FGroupItemBrandName
    saleColor = cEvtCont.FsaleColor
    priceColor = cEvtCont.FpriceColor
    orgpriceColor = cEvtCont.ForgpriceColor

    ArrcGroupTemplate=cEvtCont.fnGetMultiContentsGroupTemplateList
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

function fnGetBasicImageFullURL(basicimage,itemid)
    fnGetBasicImageFullURL=webImgUrl & "/image/basic/" + GetImageSubFolderByItemid(itemid) + "/"  + basicimage
end function

%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){
	frm.submit();
}

function TnTrainThemeItemBannerReg(){
    var winpop = window.open("/admin/eventmanage/event/V5/template/addbanner/pop_train_theme_addItems.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>","winpop","width=1450,height=800,scrollbars=yes,resizable=yes");
    winpop.focus();
}

function TnTrainThemeItemBannerDel(idx){
    document.ibfrm.target="ifrmProc";
    document.ibfrm.idx.value=idx;
    document.ibfrm.action="/admin/eventmanage/event/v5/template/addbanner/deltrainthemeitem.asp";
    document.ibfrm.submit();
}

function TnThemeBannerGroupCodeEdit(idx){
    var winG 
    winG = window.open("/admin/eventmanage/event/v5/popup/pop_event_group_select.asp?eC=<%=eCode%>&idx="+idx,"popG","width=370, height=150,scrollbars=yes,resizable=yes");
    winG.focus();
}

$(function(){
    $("#accordion").accordion();
	//드래그
	$("#subList").sortable({
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

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/event/v5/lib/pop_event_uploadimg.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan+'&sOpt='+(sName.startsWith('pcImage') ? 'P' : ''),'popImg','width=370,height=150');
	winImg.focus();
}

function jsPcSetImg(spanId , xPosition , yPosition) {
    var imgElement = document.querySelector("#"+spanId+" img");
    var imgSrc = imgElement.getAttribute('src');

    var imageNumber = spanId.substr('spanbgimg'.length);
    xPosition = xPosition == "" ? document.getElementById("xPosition"+imageNumber).value : xPosition;
    yPosition = yPosition == "" ? document.getElementById("yPosition"+imageNumber).value : yPosition;
    
    var winImg;
	winImg = window.open('pop_event_customimageset.asp?imageUrl='+imgSrc+'&xPo='+xPosition+'&yPo='+yPosition+'&imageNumber='+imageNumber,'popImg','width=600,height=400');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		eval("document.all."+sName).value = "";
		eval("document.all."+sSpan).style.display = "none";
	}
}

function jsItemDelImg(sName, sSpan) {
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		document.frmEvt[sName].value = "";
		document.getElementById(sSpan).style.display = "none";
	}
}

</script>
<form name="frmEvt" method="post" style="margin:0px;" action="customtemplate_process.asp">
<input type="hidden" name="mode" value="TU">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="menuidx" value="<%=menuidx%>">
<input type="hidden" name="device" value="<%=device%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>추천 리스트</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A" id="table">
            <div class="tabV19">
                <ul>
                    <li class="<% if device="M" then %>selected<% end if %>"><a href="?eC=<%=eCode%>&menuidx=<%=menuidx%>&device=M">Mobile / App</a></li>
                    <li class="<% if device="W" then %>selected<% end if %>"><a href="?eC=<%=eCode%>&menuidx=<%=menuidx%>&device=W">PC</a></li>
                </ul>
            </div>
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
                                <input type="radio" class="formCheckInput" name="GroupItemBrandName" value="Y" <%=chkiif(GroupItemBrandName="Y" or GroupItemBrandName="" or isnull(GroupItemBrandName),"checked","")%>>
                                브랜드명 노출
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="GroupItemBrandName" value="N" <%=chkiif(GroupItemBrandName="N","checked","")%>>
                                브랜드명 비노출
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemTitleName" value="Y" <%=chkiif(GroupItemTitleName="Y" or GroupItemTitleName="" or isnull(GroupItemTitleName),"checked","")%>>
                                    상품명 노출
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemTitleName" value="N" <%=chkiif(GroupItemTitleName="N","checked","")%>>
                                    상품명 비노출
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemPriceView" value="Y" <%=chkiif(GroupItemPriceView="Y" or GroupItemPriceView="" or isnull(GroupItemPriceView),"checked","")%>>
                                    가격 노출
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemPriceView" value="N" <%=chkiif(GroupItemPriceView="N","checked","")%>>
                                    가격 비노출
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemViewType" value="A" <%=chkiif(GroupItemViewType="A","checked","")%>>
                                    A타입
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemViewType" value="B" <%=chkiif(GroupItemViewType="B","checked","")%>>
                                    B타입
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemViewType" value="C" <%=chkiif(GroupItemViewType="C","checked","")%>>
                                    세로형A
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemViewType" value="E" <%=chkiif(GroupItemViewType="E","checked","")%>>
                                    세로형B
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemViewType" value="D" <%=chkiif(GroupItemViewType="D","checked","")%>>
                                    가로형A
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemViewType" value="F" <%=chkiif(GroupItemViewType="F","checked","")%>>
                                    가로형B
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                        </div>
                    </td>
                </tr>
				<tr>
					<th>백그라운드 이미지</th>
					<td>
                        <input type="hidden" name="BGImage" value="<%=BGImage%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=BGImage%>','BGImage','spanbgimg');return false;">백그라운드 이미지 등록</button>
                        <%IF BGImage <> "" THEN %><button class="btn4 btnGrey1 lMar05" type="button" onClick="jsItemDelImg('BGImage','spanbgimg');return false;">삭제</button><%END IF%>
                        <div class="previewThumb150W tMar10" id="spanbgimg">
                            <%IF BGImage <> "" THEN %><img src="<%=BGImage%>" width="30%" alt=""><%END IF%>
                        </div>
					</td>
				</tr>
               <tr>
                    <th>컬러지정</th>
                    <td>
                        <div class="colorPicker">
                            <input type="text" class="formControl" placeholder="ex) #FFFFFF" name="textColor" value="<%=textcolor%>" style="width:100px"> - 상품명 / 브랜드명
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="colorPicker">
                               <input type="text" class="formControl" placeholder="ex) #FFFFFF" name="saleColor" value="<%=salecolor%>" style="width:100px"> - 할인율
                            </div>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="colorPicker">
                               <input type="text" class="formControl" placeholder="ex) #FFFFFF" name="priceColor" value="<%=pricecolor%>" style="width:100px"> - 가격
                            </div>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="colorPicker">
                               <input type="text" class="formControl" placeholder="ex) #FFFFFF" name="orgpriceColor" value="<%=orgpricecolor%>" style="width:100px"> - 정가
                            </div>
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>

        <div class="tableV19BWrap tMar15 tPad25 topLineGrey2">
            <% If isArray(ArrcGroupTemplate) Then %>
            <h3 class="fs15" id="t1" >상품 정보</h3>
            <table class="tableV19A tableV19B tMar10">
                <thead>
                    <tr>
                        <th></th>
                        <th>이미지</th>
                        <th>수기 등록 이미지</th>
                        <th>상품코드</th>
                        <th>상품명&#47;카피</th>
                        <th>삭제</th>
                    </tr>
                <thead>
                <tbody id="subList">
                    <% For ix = 0 To UBound(ArrcGroupTemplate,2) %>
                    <tr>
                        <td>
                            <span class="mdi mdi-equal cBl4 fs20"></span>
                            <input type="hidden" name="bidx" value="<%=ArrcGroupTemplate(0,ix)%>">
                            <input type="hidden" name="imgurl" id="imgurl<%=ArrcGroupTemplate(0,ix)%>" value="<%=ArrcGroupTemplate(4,ix)%>">
                            <input type="hidden" name="viewidx" value="<%=ArrcGroupTemplate(8,ix)%>">
                            <input type="hidden" name="iconnew" id="iconnew<%=ArrcGroupTemplate(0,ix)%>" value="<%=ArrcGroupTemplate(6,ix)%>">
                            <input type="hidden" name="iconbest" id="iconbest<%=ArrcGroupTemplate(0,ix)%>" value="<%=ArrcGroupTemplate(7,ix)%>">
                            <input type="hidden" name="title" id="title<%=ArrcGroupTemplate(0,ix)%>" value="<%=ArrcGroupTemplate(1,ix) %>">
                            <input type="hidden" name="groupcode" value="">
                            <input type="hidden" name="brandid" value="">
                        </td>
                        <td>
                            <span class="previewThumb50W">
                            <% if ArrcGroupTemplate(4,ix)<>"" then %>
                                <img src="<%=ArrcGroupTemplate(4,ix)%>" alt="">
                            <% else %>
                                <img src="<%= fnGetBasicImageFullURL(ArrcGroupTemplate(11,ix),ArrcGroupTemplate(2,ix)) %>" alt="">
                            <% end if %>
                            </span>
                        </td>
                        <% if device = "M" then %>
                        <td>
                            <div class="previewThumb150W tMar10" id="spanbgimg<%=ix%>">
                                <%IF ArrcGroupTemplate(13,ix) <> "" THEN %><img src="<%=ArrcGroupTemplate(13,ix)%>" width="30%" alt=""><%END IF%>
                            </div>
                            <input type="hidden" name="mobileImage<%=ix%>" value="<%=ArrcGroupTemplate(13,ix) %>">
                            <button class="btnBlue1" onClick="jsSetImg('<%=eFolder%>','','mobileImage<%=ix%>','spanbgimg<%=ix%>');return false;">M 등록</button>
                            <%IF ArrcGroupTemplate(13,ix) <> "" THEN %><button class="btnGrey1" onClick="jsDelImg('mobileImage<%=ix%>','spanbgimg<%=ix%>');return false;">삭제</button><%END IF%>
                        </td>
                        <% else %>
                        <td>
                            <div class="previewThumb150W tMar10" id="spanbgimg<%=ix%>">
                                <%IF ArrcGroupTemplate(14,ix) <> "" THEN %><img src="<%=ArrcGroupTemplate(14,ix)%>" width="30%" alt="" onclick="jsPcSetImg('spanbgimg<%=ix%>','<%=ArrcGroupTemplate(15,ix)%>','<%=ArrcGroupTemplate(16,ix)%>');return false;"><%END IF%>
                            </div>
                            <input type="hidden" name="pcImage<%=ix%>" value="<%=ArrcGroupTemplate(14,ix) %>">
                            <button class="btnBlue1" onClick="jsSetImg('<%=eFolder%>','','pcImage<%=ix%>','spanbgimg<%=ix%>');return false;">PC 등록</button>
                            <%IF ArrcGroupTemplate(14,ix) <> "" THEN %><button class="btnGrey1" onClick="jsDelImg('pcImage<%=ix%>','spanbgimg<%=ix%>');return false;">삭제</button><%END IF%>
                            <div>
                                LEFT : <input type="text" name="xPosition<%=ix%>" id="xPosition<%=ix%>" size="10" value="<%=ArrcGroupTemplate(15,ix)%>" readonly>
                                TOP : <input type="text" name="yPosition<%=ix%>" id="yPosition<%=ix%>" size="10" value="<%=ArrcGroupTemplate(16,ix)%>" readonly>
                            </div>
                        </td>
                        <% end if %>
                        <td>
                            <input type="text" name="itemid" id="itemid<%=ix%>" value="<%= ArrcGroupTemplate(2,ix) %>" class="formControl" style="width:80px" readonly>
                        </td>
                        <td>
                            <input type="text" class="formControl" placeholder="상품명" name="itemname" value="<%= ArrcGroupTemplate(3,ix) %>">
                            <br/>
                            <input type="text" class="formControl" style="margin-top:10px;" placeholder="상품명2" name="itemname2" value="<%= ArrcGroupTemplate(12,ix) %>">
                        </td>
                        <td>
                            <button class="btn4 btnGrey1" onClick="TnTrainThemeItemBannerDel(<%=ArrcGroupTemplate(0,ix)%>);return false;">삭제</button>
                        </td>
                    </tr>
                    <% Next %>
                </tbody>
            </table>
            <% End If %>
        </div>
        <button class="btn4 btnBlock btnWhite2 tMar10 tPad20 bPad20 lt" onClick="TnTrainThemeItemBannerReg();return false;"><span class="mdi mdi-plus cBl4 fs15"></span> 리스트 추가</button>

	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">저장</button>
	</div>
</div>
</form>
</table>
<form method="post" name="ibfrm">
	<input type="hidden" name="idx">
	<input type="hidden" name="mode">
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->