<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_rolling_image.asp
' Discription : I형(통합형) 이벤트 롤링이미지 셋팅
' History : 2019.02.07 정태훈
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
Dim cEvtCont, ix, VideoFrameReduction, device
Dim eCode, menuidx, GroupItemPriceView, GroupItemCheck, GroupItemType
dim menudiv, viewsort, isusing, ArrcMultiContentsMenu, ArrcMultiContentsSwife

eCode = requestCheckVar(Request("eC"),10)
menuidx = requestCheckVar(Request("menuidx"),10)
device = requestCheckVar(Request("device"),1)

if device="" then device="M"

IF menuidx <> "" THEN
    set cEvtCont = new ClsMultiContentsMenu
    cEvtCont.FRectEvtCode = eCode
    cEvtCont.FRectIDX = menuidx	'멀티 컨텐츠 메뉴 idx
    cEvtCont.FRectDevice = device	'멀티 컨텐츠 디바이스 구분
    ArrcMultiContentsSwife=cEvtCont.fnGetMultiContentsSwifeList
    set cEvtCont = nothing
end if

function GetMenuDivName(menudiv)
    if menudiv="1" then
        GetMenuDivName="롤링 이미지"
    elseif menudiv="2" then
        GetMenuDivName="영상"
    elseif menudiv="3" then
        GetMenuDivName="브랜드 스토리"
    elseif menudiv="4" then
        GetMenuDivName="기차형 템플릿"
    elseif menudiv="5" then
        GetMenuDivName="추가 텍스트 박스"
    end if
end function
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function fnCheckMenudiv(objval){
    if (objval == "3"){
        document.all.TrainInfo.style.display="";
    }
    else{
        document.all.TrainInfo.style.display="none";
    }
}

function fnSwifeBannerDel(idx){
    document.ibfrm.idx.value=idx;
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

function TnSwifeBannerRegCheck(d){
    if (d == "W"){
        var winpop = window.open('/admin/eventmanage/event/v5/template/slide/pop_pcweb_themeslide.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
        winpop.focus();
    }else{
        var winpop = window.open('/admin/eventmanage/event/v5/template/slide/pop_mobile_themeslide.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
        winpop.focus();
    }
}

function fnChangeDivece(div){
    location.href="?device="+div+'&menuidx=<%=menuidx%>&eC=<%=eCode%>'
}
function jsEvtSubmit(){
    document.frmEvt.submit();
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
</script>
<form name="frmEvt" method="post" style="margin:0px;" action="rollingImagesort_process.asp">
<input type="hidden" name="mode" value="SU">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="menuidx" value="<%=menuidx%>">
<input type="hidden" name="device" value="<%=device%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>이미지 슬라이드</h1>
	</div>
	<div class="popContV19">
		<div class="tabV19">
			<ul>
				<li onClick="fnChangeDivece('M');" class="<% if device="M" then %>selected<% End If %>"><a href="#">Mobile / App</a></li>
				<li onClick="fnChangeDivece('W');" class="<% if device="W" then %>selected<% End If %>"><a href="#">PC</a></li>
			</ul>
		</div>
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th>이미지 등록</th>
                    <td>
                        <button class="btn4 btnBlue1" onclick="TnSwifeBannerRegCheck('<%=device%>');return false;">슬라이드 이미지 등록</button>
                    </td>
                </tr>
			</tbody>
        </table>
	</div>
    <% If isArray(ArrcMultiContentsSwife) Then %>
    <div class="tableV19BWrap tMar15 tPad25 topLineGrey2">
        <table class="tableV19A tableV19B tMar10">
            <thead>
                <tr>
                    <th></th>
                    <th>이미지</th>
                    <th>삭제</th>
                </tr>
            <thead>
            <tbody id="subList">
                <% For ix = 0 To UBound(ArrcMultiContentsSwife,2) %>
                <tr>
                    <td>
                        <span class="mdi mdi-equal cBl4 fs20"></span><input type="hidden" name="idx" value="<%=ArrcMultiContentsSwife(0,ix)%>"><input type="hidden" name="viewidx" value="<%=ArrcMultiContentsSwife(2,ix)%>">
                    </td>
                    <td><span class="previewThumb50W"><img src="<%=ArrcMultiContentsSwife(1,ix)%>" alt=""></span></td>
                    <td>
                        <button class="btn4 btnGrey1" onClick="fnSwifeBannerDel(<%=ArrcMultiContentsSwife(0,ix)%>);return false;">삭제</button>
                    </td>
                </tr>
                <% Next %>
            </tbody>
        </table>
    </div>
    <% End If %>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
        <button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">확인</button>
	</div>
</div>
</form>
<form method="post" name="ibfrm" action="rollingImagesort_process.asp">
	<input type="hidden" name="idx">
	<input type="hidden" name="mode" value="SD">
    <input type="hidden" name="device" value="<%=device%>">
    <input type="hidden" name="evt_code" value="<%=eCode%>">
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->