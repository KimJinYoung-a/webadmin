<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_grouptemplateinfo.asp
' Discription : I형(통합형) 이벤트 기차형 템플릿 셋팅
' History : 2019.02.12 정태훈
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
Dim eCode, menuidx, GroupItemPriceView, GroupItemCheck, GroupItemType
dim menudiv, viewsort, isusing, ArrcGroupTemplate, eFolder, eregdate
Dim BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin, textColor

eCode = requestCheckVar(Request("eC"),10)
menuidx = requestCheckVar(Request("menuidx"),10)
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
if textColor="" or isnull(textColor) then textColor="2"
function fnGetBasicImageFullURL(basicimage,itemid)
    fnGetBasicImageFullURL=webImgUrl & "/image/basic/" + GetImageSubFolderByItemid(itemid) + "/"  + basicimage
end function
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){
	frm.submit();
}

function fnCheckMenudiv(objval){
    if (objval == "3"){
        document.all.TrainInfo.style.display="";
    }
    else{
        document.all.TrainInfo.style.display="none";
    }
}

function fnSwifeBannerDel(idx){
    document.ibfrm.target="ifrmProc";
    document.ibfrm.idx.value=idx;
    document.ibfrm.action="/admin/eventmanage/event/v5/template/addbanner/delmulticontents.asp";
    document.ibfrm.submit();
}

function TnPriceViewCheck(viewyn){
    if(viewyn=="Y"){
        document.all.priceview.style.display="";
    }
    else{
        document.all.priceview.style.display="none";
    }
}

function TnTrainThemeItemBannerReg(){
    
    if(!$("input:radio[name='GroupItemType']").is(":checked")){
        alert("유형을 선택해주세요.");
        return false;
    }
    if($("input:radio[name='GroupItemType']:checked").val()=="T"){
        var winpop = window.open("/admin/eventmanage/event/V5/template/addbanner/pop_train_theme_addItems.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>&GroupItemCheck="+$("input:radio[name='GroupItemCheck']:checked").val(),"winpop","width=1450,height=800,scrollbars=yes,resizable=yes");
    }else{
        var winpop = window.open("/admin/eventmanage/event/V5/template/addbanner/pop_train_theme_addbanner.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>&GroupItemCheck="+$("input:radio[name='GroupItemCheck']:checked").val(),"winpop","width=1450,height=800,scrollbars=yes,resizable=yes");
    }
    winpop.focus();
}

function TnPriceViewCheck(viewyn){
    if(viewyn=="Y"){
        document.all.priceview.style.display="";
    }
    else{
        document.all.priceview.style.display="none";
    }
}

function fnViewTempType(objval){
    var templen = $("input[name=itemid]").length;
    if (objval == "B"){
        $("#grouptemp1").show();
        $("#grouptemp2").show();
        $("#grouptemp3").hide();
        $("#grouptemp4").hide();
        $("#grouptemp6").hide();
		$("#linktype1").show();
		$("#linktype2").show();
        $('input[name="GroupItemCheck"]:radio[value="T"]').attr("checked",true);
        $('input[name="GroupItemType"]:radio[value="B"]').attr("checked",true);
        for(var i=0; i<templen; i++){
            $("#grouptemp1"+i).show();
            $("#grouptemp2"+i).show();
            $("#grouptemp3"+i).hide();
            $("#grouptemp4"+i).hide();
            $("#grouptemp6"+i).hide();
            $("#itemid"+i).attr("readonly",false);
        }
        $("#t1").hide();
        $("#t2").show();
        $("#t3").hide();
    }else if (objval == "M"){
        $("#grouptemp1").show();
        $("#grouptemp2").hide();
        $("#grouptemp3").hide();
        $("#grouptemp4").hide();
        $("#grouptemp6").show();
		$("#linktype1").show();
		$("#linktype2").show();
        $('input[name="GroupItemCheck"]:radio[value="B"]').attr("checked",true);
        $('input[name="GroupItemType"]:radio[value="B"]').attr("checked",true);
        for(var i=0; i<templen; i++){
            $("#grouptemp1"+i).show();
            $("#grouptemp6"+i).show();
            $("#grouptemp2"+i).hide();
            $("#grouptemp3"+i).hide();
            $("#grouptemp4"+i).hide();
            $("#itemid"+i).attr("readonly",false);
        }
        $("#t1").hide();
        $("#t2").hide();
        $("#t3").show();
    }else{
        $("#grouptemp3").show();
        $("#grouptemp4").show();
        $("#grouptemp1").hide();
        $("#grouptemp2").hide();
        $("#grouptemp6").hide();
		$("#linktype1").hide();
		$("#linktype2").hide();
        $('input[name="GroupItemCheck"]:radio[value="I"]').attr("checked",true);
        $('input[name="GroupItemType"]:radio[value="T"]').attr("checked",true);
        for(var i=0; i<templen; i++){
            $("#grouptemp3"+i).show();
            $("#grouptemp4"+i).show();
            $("#grouptemp1"+i).hide();
            $("#grouptemp2"+i).hide();
            $("#grouptemp6"+i).hide();
            $("#itemid"+i).attr("readonly",true);
        }
        $("#t1").show();
        $("#t2").hide();
        $("#t3").hide();
    }
}

function TnChangeThemeBannerEdit(sFolder, sImg, sName, sSpan){
    var winImg;
    winImg = window.open("/admin/eventmanage/event/v5/template/addbanner/pop_event_uploadimgV5.asp?yr=<%=Year(eregdate)%>&sF="+sFolder+"&sImg="+sImg+"&sName="+sName+"&sSpan="+sSpan,"popImg","width=370,height=150");
    winImg.focus();
}

function TnTrainThemeItemBannerDel(idx){
    document.ibfrm.target="ifrmProc";
    document.ibfrm.idx.value=idx;
    document.ibfrm.action="/admin/eventmanage/event/v5/template/addbanner/deltrainthemeitem.asp";
    document.ibfrm.submit();
}

function TnThemeBannerItemCodeEdit(idx){
    var winG 
    winG = window.open("/admin/eventmanage/event/v5/popup/pop_event_additem.asp?eC=<%=eCode%>&idx="+idx,"popG","width=1024, height=700,scrollbars=yes,resizable=yes");
    winG.focus();
}

function fnSSearchBrandPop(idx){
    var wBrandView;
    wBrandView = window.open("popBrandSearch.asp?frmName=frmEvt&idx="+idx,"winBrand","width=1400,height=800,scrollbars=yes,resizable=yes");
    wBrandView.focus();
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

function fnIconNewCheck(obj,objnum){
    if(obj.checked){
        $("#iconnew"+objnum).val("Y");
    }
    else{
        $("#iconnew"+objnum).val("N");
    }
}

function fnIconBestCheck(obj,objnum){
    if(obj.checked){
        $("#iconbest"+objnum).val("Y");
    }
    else{
        $("#iconbest"+objnum).val("N");
    }
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

//색상코드 선택
function selBackGroundColorChip(cd) {
    var i;
    frmEvt.textColor.value= cd;
    for(i=1;i<=2;i++) {
        $("#bcline"+i).removeClass("picked");
    }
    $("#bcline"+cd).addClass("picked");
}

</script>
<form name="frmEvt" method="post" style="margin:0px;" action="grouptemplate_process.asp">
<input type="hidden" name="mode" value="TU">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="menuidx" value="<%=menuidx%>">
<input type="hidden" name="textColor" value="<%=textColor%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>추천 리스트</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A" id="table">
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
                                <input type="radio" class="formCheckInput" name="GroupItemType" value="T"<% if GroupItemType="T" then response.write " checked" %> onclick="fnViewTempType('T');TnPriceViewCheck('Y');">
                                아이템 선택
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="GroupItemType" value="B"<% if GroupItemType="B" then response.write " checked" %> onclick="fnViewTempType('B')">
                                이미지업로드
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemCheck" value="I" onclick="TnPriceViewCheck('Y');fnViewTempType('T')"<% if GroupItemCheck="I" or GroupItemCheck="" then response.write " checked" %>>
                                    상품 리스트
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline" id="linktype1" style="display:<% If GroupItemCheck="I" Then Response.write "none"%>">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemCheck" value="T" onClick="TnPriceViewCheck('Y');fnViewTempType('B')"<% if GroupItemCheck="T" then response.write " checked" %>>
                                    그룹연결 리스트
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline" id="linktype2" style="display:<% If GroupItemCheck="I" Then Response.write "none"%>">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemCheck" value="B" onClick="TnPriceViewCheck('Y');fnViewTempType('M')"<% if GroupItemCheck="B" then response.write " checked" %>>
                                    브랜드연결 리스트
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1" id="priceview">
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemPriceView" value="Y"<% if GroupItemPriceView="Y" or GroupItemPriceView="" then response.write " checked" %>>
                                    가격노출
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="radio" class="formCheckInput" name="GroupItemPriceView" value="N"<% if GroupItemPriceView="N" then response.write " checked" %>>
                                    가격 미노출
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
                                    <input type="radio" class="formCheckInput" name="contentsAlign" value="1"<% if contentsAlign="1" then response.write " checked"%> onclick="fnAlignTypeChange(this.value);">
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
                <tr>
                    <th>상품명 / 가격 텍스트 색</th>
                    <td>
                        <div class="colorPicker">
                            <ul>
                                <li<% If textColor="1" Then %> class="picked"<% End If %> onClick="selBackGroundColorChip(1);" id="bcline1"><span style="background-color:#ffffff;"></span></li>
                                <li<% If textColor="2" Then %> class="picked"<% End If %> onClick="selBackGroundColorChip(2);" id="bcline2"><span style="background-color:#000000;"></span></li>
                            </ul>
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>

        <div class="tableV19BWrap tMar15 tPad25 topLineGrey2">
            <% If isArray(ArrcGroupTemplate) Then %>
            <h3 class="fs15" id="t1" style="display:<% If GroupItemCheck<>"I" Then Response.write "none"%>">상품 정보</h3>
            <h3 class="fs15" id="t2" style="display:<% If GroupItemCheck<>"T" Then Response.write "none"%>">그룹 정보</h3>
            <h3 class="fs15" id="t3" style="display:<% If GroupItemCheck<>"B" Then Response.write "none"%>">브랜드 정보</h3>
            <table class="tableV19A tableV19B tMar10">
                <thead>
                    <tr>
                        <th></th>
                        <th>이미지</th>
                        <th>상품코드</th>
                        <th>상품명&#47;카피</th>
                        <th id="grouptemp2" style="display:<% If GroupItemType="T" Then Response.write "none"%>">그룹연결코드</th>
                        <th id="grouptemp6" style="display:<% If GroupItemType="T" or GroupItemCheck<>"B" Then Response.write "none"%>">브랜드ID</th>
                        <th id="grouptemp4" style="display:<% If GroupItemType="B" Then Response.write "none"%>">NEW/BEST</th>
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
                        </td>
                        <td><span class="previewThumb50W"><% if ArrcGroupTemplate(4,ix)<>"" then %><img src="<%=ArrcGroupTemplate(4,ix)%>" alt=""><% else %><img src="<%= fnGetBasicImageFullURL(ArrcGroupTemplate(11,ix),ArrcGroupTemplate(2,ix)) %>" alt=""><% end if %></span></td>
                        <td><input type="text" name="itemid" id="itemid<%=ix%>" value="<%= ArrcGroupTemplate(2,ix) %>" class="formControl" style="width:80px"></td>
                        
                        <td id="grouptemp1<%=ix%>" style="display:<% If GroupItemType="T" Then Response.write "none"%>"><input type="text" class="formControl" placeholder="상품명" name="title" id="title<%=ArrcGroupTemplate(0,ix)%>" value="<%=ArrcGroupTemplate(1,ix) %>"></td>
                        
                        <td id="grouptemp2<%=ix%>" style="display:<% If GroupItemType="T" Then Response.write "none"%>"><input type="text" class="formControl formControl550 bgGry2 cGy1" placeholder="그룹연결설정" name="groupcode" id="groupcode<%=ArrcGroupTemplate(0,ix)%>" value="<%=ArrcGroupTemplate(5,ix) %>" onClick="TnThemeBannerGroupCodeEdit(<%=ArrcGroupTemplate(0,ix)%>);"></td>

                        <td id="grouptemp6<%=ix%>" style="display:<% If GroupItemType="T" or GroupItemCheck<>"B" Then Response.write "none"%>"><input type="text" class="formControl formControl550 bgGry2 cGy1" placeholder="브랜드연결설정"name="brandid" id="brandid<%=ArrcGroupTemplate(0,ix)%>" value="<%= ArrcGroupTemplate(10,ix) %>" onClick="fnSSearchBrandPop(<%=ArrcGroupTemplate(0,ix)%>);"></td>

                        <td id="grouptemp3<%=ix%>" style="display:<% If GroupItemType="B" Then Response.write "none"%>"><input type="text" class="formControl" placeholder="상품명" name="itemname" id="itemname<%=ArrcGroupTemplate(0,ix)%>" value="<%= ArrcGroupTemplate(3,ix) %>"></td>

                        <td id="grouptemp4<%=ix%>" style="display:<% If GroupItemType="B" Then Response.write "none"%>">
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="checkbox" name="iconnewcheck" value="Y"<% if ArrcGroupTemplate(6,ix)="Y" then response.write " checked" %> onClick="fnIconNewCheck(this,<%=ArrcGroupTemplate(0,ix)%>);">
                                    NEW
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
                            <div class="formInline">
                                <label class="formCheckLabel">
                                    <input type="checkbox" name="iconbestcheck" value="Y"<% if ArrcGroupTemplate(7,ix)="Y" then response.write " checked" %> onClick="fnIconBestCheck(this,<%=ArrcGroupTemplate(0,ix)%>);">
                                    BEST
                                    <i class="inputHelper"></i>
                                </label>
                            </div>
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