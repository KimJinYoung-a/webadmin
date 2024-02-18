<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_customgroupitem.asp
' Discription : I형(통합형) 이벤트 커스텀 상품 등롭 타입별 
' History : 2020.12.03 이종화
' Update  : 2021.01.29 이전도
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
Dim BGImage, BGImagePC, BGColorLeft, BGColorRight, contentsAlign, Margin, textColor 
Dim device , GroupItemBrandName , GroupItemTitleName
dim saleColor , priceColor , orgpriceColor , MarginBottom , MarginColor , MarginBottomColor
Dim MarginPC, MarginBottomPC , MarginColorPC , MarginBottomColorPC

eCode = requestCheckVar(Request("eC"),10) '// 이벤트 코드
menuidx = requestCheckVar(Request("menuidx"),10) '// 컨텐츠메뉴 인덱스
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
    BGImagePC = cEvtCont.FBGImagePC
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
	MarginBottom = cEvtCont.FMarginBottom
	MarginColor = cEvtCont.FMarginColor
	MarginBottomColor = cEvtCont.FMarginBottomColor
	MarginPC = cEvtCont.FMarginPC
	MarginBottomPC = cEvtCont.FMarginBottomPC
	MarginColorPC = cEvtCont.FMarginColorPC
	MarginBottomColorPC = cEvtCont.FMarginBottomColorPC

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

Dim itemWebImageUrl : itemWebImageUrl = "http://webimage.10x10.co.kr"

function fnGetBasicImageFullURL(basicimage,itemid)
    fnGetBasicImageFullURL=itemWebImageUrl & "/image/basic/" + GetImageSubFolderByItemid(itemid) + "/"  + basicimage
end function

%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
// 이벤트 저장
function jsEvtSubmit(frm){
	frm.submit();
}

function TnTrainThemeItemBannerReg(){
    var winpop = window.open("/admin/eventmanage/event/V5/template/addbanner/pop_train_theme_addItems.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>","winpop","width=1450,height=800,scrollbars=yes,resizable=yes");
    winpop.focus();
}

// 상품정보 상품 삭제
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
    let sOpt;
    if( sName.startsWith('pcImage') ) {
        sOpt = 'P';
    } else if( sName.startsWith('BGImage') ) {
        sOpt = 'Q';
    } else {
        sOpt = '';
    }

	winImg = window.open('/admin/eventmanage/event/v5/lib/pop_event_uploadimg.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan+'&sOpt='+sOpt,'popImg','width=370,height=150');
	winImg.focus();
}

function jsPcSetImg(spanId , xPosition , yPosition) {
    var imgElement = document.querySelector("#"+spanId+" img");
    var imgSrc = imgElement.getAttribute('src');

    var imageNumber = spanId.substr('spanbgimg'.length);
    xPosition = xPosition == "" ? document.getElementById("xPosition"+imageNumber).value : xPosition;
    yPosition = yPosition == "" ? document.getElementById("yPosition"+imageNumber).value : yPosition;

    // todo : 팝업으로 전달해야 할 파라미터
    // 기본 정보 : 가격, 최종가격
    // 노출 여부 : 상품명, 브랜드명, 가격
    // 색상 코드 : 상품명/브랜드명, 할인율, 판매가격, 최종가격
    const this_type = document.frmEvt.GroupItemViewType.value;
    if( this_type !== 'A' && this_type !== 'B' )
        return false;

    const frmEvt = document.frmEvt;
    const this_tr = imgElement.closest('tr');
    const itemid = this_tr.querySelector('input[name=itemid]').value;
    const info_data = {};
    info_data.basic = {
        'type' : this_type, // 타입(A,B)
        'itemid' : itemid, // 상품ID
        'itemname' : this_tr.querySelector('input[name=itemname]').value, // 상품명
    };
    info_data.show = {
        'itemname' : frmEvt.GroupItemTitleName.value,
        'brandname' : frmEvt.GroupItemBrandName.value,
        'price' : frmEvt.GroupItemPriceView.value
    };
    info_data.color = {
        'item_and_brand_name' : frmEvt.textColor.value,
        'sale_percent' : frmEvt.saleColor.value,
        'org_price' : frmEvt.orgpriceColor.value,
        'price' : frmEvt.priceColor.value
    };
    const encoded_info = encodeURI(encodeURIComponent(JSON.stringify(info_data)));
    
    var winImg;
	winImg = window.open('pop_event_customimageset.asp?imageUrl='+imgSrc+'&xPo='+xPosition+'&yPo='+yPosition+'&imageNumber='+imageNumber+'&info='+encoded_info+'&itemid='+itemid,'popImg','width=600,height=400');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		eval("document.all."+sName).value = "";
		eval("document.all."+sSpan).style.display = "none";
	}
}

function jsItemDelImg() {
	if(confirm("배경 이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
        const sName = frmEvt.device.value === 'M' ? 'BGImage' : 'BGImagePC';
		document.frmEvt[sName].value = "";
		document.getElementById('spanbgimg').style.display = "none";
        document.querySelector('.deleteBtn').remove();
	}
}

</script>
<style>
    strong.nameTitle {display: inline-block;position: relative;margin-top: 2px;margin: 0 32px;font-size: 1.25rem;vertical-align: middle;}
    .colorPicker span.preview {display:inline-block;height:32px;width:32px;vertical-align:middle;margin-left:15px;}
</style>
<form name="frmEvt" method="post" style="margin:0px;" action="customtemplate_process.asp">
<input type="hidden" name="mode" value="TU">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="menuidx" value="<%=menuidx%>">
<input type="hidden" name="device" value="<%=device%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>가격연동 템플릿</h1>
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
				<col style="width:160px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th style="height: 121px;padding:43px 20px;">타입</th>
                    <td>
                        <strong class="nameTitle">A형</strong>
                        <div class="formInline">
                            <label class="formCheckLabel" style="font-weight:500;">
                                <input type="radio" class="formCheckInput" name="GroupItemViewType" value="A" <%=chkiif(GroupItemViewType="A" or GroupItemViewType="" or isnull(GroupItemViewType),"checked","")%>>
                                이미지형
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel" style="font-weight:500;">
                                <input type="radio" class="formCheckInput" name="GroupItemViewType" value="C" <%=chkiif(GroupItemViewType="C","checked","")%>>
                                세로형
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel" style="font-weight:500;">
                                <input type="radio" class="formCheckInput" name="GroupItemViewType" value="D" <%=chkiif(GroupItemViewType="D","checked","")%>>
                                가로형
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <strong class="nameTitle">B형</strong>
                            <div class="formInline">
                                <label class="formCheckLabel" style="font-weight:500;">
                                    <input type="radio" class="formCheckInput" name="GroupItemViewType" value="B" <%=chkiif(GroupItemViewType="B","checked","")%>>
                                    이미지형
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel" style="font-weight:500;">
                                    <input type="radio" class="formCheckInput" name="GroupItemViewType" value="E" <%=chkiif(GroupItemViewType="E","checked","")%>>
                                    세로형
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel" style="font-weight:500;">
                                    <input type="radio" class="formCheckInput" name="GroupItemViewType" value="F" <%=chkiif(GroupItemViewType="F","checked","")%>>
                                    가로형
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                        </div>
                    </td>
                </tr>
                <tr>
                    <th style="height: 121px;padding:43px 20px;">입력방법</th>
                    <td>
                        <strong class="nameTitle" style="min-width:60px;">브랜드명</strong>
                        <div class="formInline">
                            <label class="formCheckLabel" style="font-weight:500;">
                                <input type="radio" class="formCheckInput" name="GroupItemBrandName" value="Y" <%=chkiif(GroupItemBrandName="Y" or GroupItemBrandName="" or isnull(GroupItemBrandName),"checked","")%>>
                                노출
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel" style="font-weight:500;">
                                <input type="radio" class="formCheckInput" name="GroupItemBrandName" value="N" <%=chkiif(GroupItemBrandName="N","checked","")%>>
                                비노출
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <strong class="nameTitle" style="min-width:60px;">상품명</strong>
                            <div class="formInline">
                                <label class="formCheckLabel" style="font-weight:500;">
                                    <input type="radio" class="formCheckInput" name="GroupItemTitleName" value="Y" <%=chkiif(GroupItemTitleName="Y" or GroupItemTitleName="" or isnull(GroupItemTitleName),"checked","")%>>
                                    노출
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel" style="font-weight:500;">
                                    <input type="radio" class="formCheckInput" name="GroupItemTitleName" value="N" <%=chkiif(GroupItemTitleName="N","checked","")%>>
                                    비노출
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <strong class="nameTitle" style="min-width:60px;">가격</strong>
                            <div class="formInline">
                                <label class="formCheckLabel" style="font-weight:500;">
                                    <input type="radio" class="formCheckInput" name="GroupItemPriceView" value="Y" <%=chkiif(GroupItemPriceView="Y" or GroupItemPriceView="" or isnull(GroupItemPriceView),"checked","")%>>
                                    노출
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel" style="font-weight:500;">
                                    <input type="radio" class="formCheckInput" name="GroupItemPriceView" value="N" <%=chkiif(GroupItemPriceView="N","checked","")%>>
                                    비노출
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>컬러지정</th>
                    <td>
                        <div class="colorPicker">
                            <strong class="nameTitle" style="min-width:123px;">상품명 / 브랜드명</strong>
                            <input type="text" class="formControl" placeholder="ex) #FFFFFF" name="textColor" value="<%=textcolor%>" style="width:115px">
                            <span class="preview" style="background:<%=textcolor%>;"></span>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="colorPicker">
                                <strong class="nameTitle" style="min-width:123px;">할인율</strong>
                                <input type="text" class="formControl" placeholder="ex) #FFFFFF" name="saleColor" value="<%=saleColor%>" style="width:115px">
                                <span class="preview" style="background:<%=saleColor%>;"></span>
                            </div>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="colorPicker">
                                <strong class="nameTitle" style="min-width:123px;">판매가격</strong>
                                <input type="text" class="formControl" placeholder="ex) #FFFFFF" name="orgpriceColor" value="<%=orgpriceColor%>" style="width:115px">
                                <span class="preview" style="background:<%=orgpriceColor%>;"></span>
                            </div>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="colorPicker">
                                <strong class="nameTitle" style="min-width:123px;">최종가격</strong>
                                <input type="text" class="formControl" placeholder="ex) #FFFFFF" name="priceColor" value="<%=priceColor%>" style="width:115px">
                                <span class="preview" style="background:<%=priceColor%>;"></span>
                            </div>
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>상단 여백</th>
                    <td>
                        <div class="formInline" style="padding-left:32px;">
                            <input type="text" class="formControl formControl550" maxlength="6" placeholder="상단 여백" name="Margin" id="Margin" value="<%=chkiif(device="M",Margin,MarginPC)%>" style="width:135px;"> px
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>상단 여백 배경</th>
                    <td>
                        <div class="formInline colorPicker" style="padding-left:32px;">
                            <input type="text" class="formControl formControl550" placeholder="ex) #FFFFFF" name="MarginColor" id="MarginColor" value="<%=chkiif(device="M",MarginColor,MarginColorPC)%>" style="width:135px;">
                            <span class="preview" style="background:<%=chkiif(device="M",MarginColor,MarginColorPC)%>;"></span>
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>하단 여백</th>
                    <td>
                        <div class="formInline" style="padding-left:32px;">
                            <input type="text" class="formControl formControl550" maxlength="6" placeholder="하단 여백" name="MarginBottom" id="MarginBottom" value="<%=chkiif(device="M",MarginBottom,MarginBottomPC)%>" style="width:135px;"> px
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>하단 여백 배경</th>
                    <td>
                        <div class="formInline colorPicker" style="padding-left:32px;">
                            <input type="text" class="formControl formControl550" placeholder="ex) #FFFFFF" name="MarginBottomColor" id="MarginBottomColor" value="<%=chkiif(device="M",MarginBottomColor,MarginBottomColorPC)%>" style="width:135px;">
                            <span class="preview" style="background:<%=chkiif(device="M",MarginBottomColor,MarginBottomColorPC)%>;"></span>
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>

        <div class="tableV19B tMar15 tPad25 topLineGrey2">
            <% If isArray(ArrcGroupTemplate) Then %>
            <h3 class="fs15" style="text-align:left;padding-left:15px;">상품 정보</h3>
            <table class="tableV19A tableV19B tMar10">
                <thead>
                    <tr>
                        <th style="padding:16px 0;"></th>
                        <th style="padding:16px 0;">제품썸네일</th>
                        <th style="padding:16px 0;">썸네일수기등록</th>
                        <th style="padding:16px 0;">상품코드</th>
                        <th style="padding:16px 0;">상품명</th>
                        <th style="padding:16px 0;">삭제</th>
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
                                <img src="<%=ArrcGroupTemplate(4,ix)%>" alt="" style="width:49px;height:49px;">
                            <% else %>
                                <img src="<%= fnGetBasicImageFullURL(ArrcGroupTemplate(11,ix),ArrcGroupTemplate(2,ix)) %>" alt="" style="width:49px;height:49px;">
                            <% end if %>
                            </span>
                        </td>
                        <% if device = "M" then %>
                        <td>
                            <div class="previewThumb150W tMar10" id="spanbgimg<%=ix%>">
                                <%IF ArrcGroupTemplate(13,ix) <> "" THEN %><img src="<%=ArrcGroupTemplate(13,ix)%>" style="width:100px;height:100px;margin-bottom:10px;" alt=""><%END IF%>
                            </div>
                            <input type="hidden" name="mobileImage<%=ix%>" value="<%=ArrcGroupTemplate(13,ix) %>">
                            <button class="btnBlue1" type="button" onClick="jsSetImg('<%=eFolder%>','','mobileImage<%=ix%>','spanbgimg<%=ix%>');return false;">M 등록</button>
                            <%IF ArrcGroupTemplate(13,ix) <> "" THEN %><button class="btnGrey1" onClick="jsDelImg('mobileImage<%=ix%>','spanbgimg<%=ix%>');return false;">삭제</button><%END IF%>
                        </td>
                        <% else %>
                        <td>
                            <div class="previewThumb150W tMar10" id="spanbgimg<%=ix%>">
                                <%IF ArrcGroupTemplate(14,ix) <> "" THEN %><img src="<%=ArrcGroupTemplate(14,ix)%>" style="width:100px;height:100px;margin-bottom:10px;" alt="" onclick="jsPcSetImg('spanbgimg<%=ix%>','<%=ArrcGroupTemplate(15,ix)%>','<%=ArrcGroupTemplate(16,ix)%>');return false;"><%END IF%>
                            </div>
                            <input type="hidden" name="pcImage<%=ix%>" value="<%=ArrcGroupTemplate(14,ix) %>">
                            <button class="btnBlue1" onClick="jsSetImg('<%=eFolder%>','','pcImage<%=ix%>','spanbgimg<%=ix%>');return false;">PC 등록</button>
                            <%IF ArrcGroupTemplate(14,ix) <> "" THEN %><button class="btnGrey1" onClick="jsDelImg('pcImage<%=ix%>','spanbgimg<%=ix%>');return false;">삭제</button><%END IF%>
                            <div>
                                <input type="hidden" name="xPosition<%=ix%>" id="xPosition<%=ix%>" value="<%=ArrcGroupTemplate(15,ix)%>">
                                <input type="hidden" name="yPosition<%=ix%>" id="yPosition<%=ix%>" value="<%=ArrcGroupTemplate(16,ix)%>">
                            </div>
                        </td>
                        <% end if %>
                        <td>
                            <input type="text" name="itemid" id="itemid<%=ix%>" value="<%= ArrcGroupTemplate(2,ix) %>" class="formControl" style="width:80px">
                        </td>
                        <td>
                            <input type="text" class="formControl" placeholder="상품명" name="itemname" value="<%= ArrcGroupTemplate(3,ix) %>">
                        </td>
                        <td>
                            <button type="button" class="btn4 btnGrey1" onClick="TnTrainThemeItemBannerDel(<%=ArrcGroupTemplate(0,ix)%>);return false;">삭제</button>
                        </td>
                    </tr>
                    <% Next %>
                </tbody>
            </table>
            <% End If %>
        </div>

        <button class="btn4 btnBlock btnWhite2 tMar20 tPad20 bPad20 lt" type="button" onClick="TnTrainThemeItemBannerReg();return false;"><span class="mdi mdi-plus cBl4 fs15"></span>상품 추가</button>

        <table class="tableV19A tMar10" id="backImageTable">
            <colgroup>
				<col style="width:160px;">
				<col style="width:auto;">
			</colgroup>
            <tbody>
                <% If device = "M" Then '// Mobile, APP %>
                    <th style="padding: 32px 0 31px 20px;">배경이미지(M)</th>
                    <td>
                        <input type="hidden" name="BGImage" value="<%=BGImage%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=BGImage%>','BGImage','spanbgimg');return false;">백그라운드 이미지 등록</button>
                        <%IF BGImage <> "" THEN %><button class="btn4 btnGrey1 lMar05 deleteBtn" type="button" onClick="jsItemDelImg();return false;">삭제</button><%END IF%>
                        <div class="previewThumb150W tMar10" id="spanbgimg">
                            <%IF BGImage <> "" THEN %><img src="<%=BGImage%>" width="30%" alt=""><%END IF%>
                        </div>
                    </td>
                <% Else '// PC %>
                    <th style="padding: 32px 0 31px 20px;">배경이미지(PC)</th>
                    <td>
                        <input type="hidden" name="BGImagePC" value="<%=BGImagePC%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=BGImagePC%>','BGImagePC','spanbgimg');return false;">백그라운드 이미지 등록</button>
                        <%IF BGImagePC <> "" THEN %><button class="btn4 btnGrey1 lMar05 deleteBtn" type="button" onClick="jsItemDelImg();return false;">삭제</button><%END IF%>
                        <div class="previewThumb150W tMar10" id="spanbgimg">
                            <%IF BGImagePC <> "" THEN %><img src="<%=BGImagePC%>" width="30%" alt=""><%END IF%>
                        </div>
                    </td>
                <% End If %>
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
<form method="post" name="ibfrm">
	<input type="hidden" name="idx">
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>

<script>
    let updating_itemid; // 수정중인 상품ID
    const itemidArr = document.querySelectorAll('input[name=itemid]'); // 상품ID Input Array
    // 입력한 상품ID에 해당하는 상품정보 받아와 이미지, 상품명 변경
    const getItemInfo = function(e) {
        if( e.keyCode === 13 ) {
            e.preventDefault();
            const itemid = e.target.value;
            console.log(itemid);

            $.ajax({
                type : 'GET',
                url : '../lib/ajaxGetItemInfo.asp',
                data : { 'itemid' : itemid },
                dataType : 'json',
                success : function(data) {
                    if( data.result ) {
                        const this_tr = e.target.closest('tr');
                        this_tr.querySelector('img').src = getBasicImageFullURL(data.itemimage, data.itemid);
                        this_tr.querySelector('input[name=itemname]').value = data.itemname;
                        updating_itemid = data.itemid;
                    } else {
                        alert(data.message);
                        e.target.value = updating_itemid;
                    }
                },
                error : function(xhr) {
                    console.log(xhr.responseText);
                }
            })
        }
    }
    // 상품ID Input Array 이벤트 추가
    itemidArr.forEach(input => {
        input.addEventListener('focus', function(e) {updating_itemid = e.target.value;}); // Focus되었을 때 수정중ID에 넣어줌
        input.addEventListener('blur', function(e) {e.target.value = updating_itemid; updating_itemid = null;}); // Blur 되었을 때 수정중ID로 원복 & 수정중ID 초기화
        input.addEventListener('keydown', getItemInfo); // 엔터 입력 시 상품정보 조회 후 Set
    });
    // Get 이미지 Url
    const getBasicImageFullURL = function(basicimage,itemid) {
        return `<%=itemWebImageUrl%>/image/basic/${itemid < 100000 ? '0' : ''}${Math.floor(itemid/10000)}/${basicimage}`;
    }


    // 컬러지정 Input
    const colorInputArr = document.querySelectorAll('.colorPicker input');
    const setColorDefaultSharp = function(e) {
        if( e.target.value === '' )
            e.target.value = '#';
    }
    /**
     * 색 Input에서 Blur 되었을떄
     * 정상이면 대문자로 치환
     * 비정상이면 빈값 입력 후
     * 색 코드값을 우측 span에 보여줌
    **/
    const blurColor = function(e) {
        const thisValue = e.target.value;

        if( thisValue.startsWith('#') && thisValue.length !== 4 && thisValue.length !== 7 ) {
            e.target.value = '';
        } else {
            e.target.value = e.target.value.toUpperCase();
        }

        e.target.parentElement.querySelector('span.preview').style.background = e.target.value;
    }
    /**
     * 색 Input에서 엔터를 눌렀을 때
     * 정상이면 대문자로 치환 후 색 코드값을 우측 span에 보여줌
     * 비정상이면 '#'으로 되돌림
    **/
    const keyDownColor = function(e) {
        if( e.keyCode === 13 ) {
            e.preventDefault();
            const thisValue = e.target.value;

            if( thisValue.startsWith('#') && thisValue.length !== 4 && thisValue.length !== 7 ) {
                e.target.value = '#';
            } else {
                e.target.value = e.target.value.toUpperCase();
                e.target.parentElement.querySelector('span.preview').style.background = e.target.value;
            }
        }
    }
    colorInputArr.forEach(input => {
        input.addEventListener('focus', setColorDefaultSharp);
        input.addEventListener('blur', blurColor);
        input.addEventListener('keydown', keyDownColor);
    });
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->