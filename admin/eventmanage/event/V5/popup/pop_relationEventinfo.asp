<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_relationEventinfo.asp
' Discription : I형(통합형) 이벤트 관련 이벤트 셋팅
' History : 2019.02.27 정태훈
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
<%
Dim cEvtCont, ix
Dim eCode, menuidx, GroupItemPriceView, GroupItemCheck, GroupItemType
dim menudiv, viewsort, isusing, ArrcEvtInfo, arrevEntKind, arrevEntState, eventCount

eCode = requestCheckVar(Request("eC"),10)

IF eCode <> "" THEN
    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode
    ArrcEvtInfo=cEvtCont.fnGetRelationEvent
    set cEvtCont = nothing
end if
If isArray(ArrcEvtInfo) Then
    eventCount=UBound(ArrcEvtInfo,2)
else
    eventCount=0
end if
'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
arrevEntKind = fnSetCommonCodeArr("eventkind",False)
arrevEntState= fnSetCommonCodeArr("eventstate",False)	
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){
    var evtCnt=<%=eventCount+1%>;
    if(frm.ecode.value==""){
        frm.submit();
    }
    else{
        if(evtCnt>2){
            alert("최대 등록 수량을 초과했습니다.");
        }
        else{
            frm.submit();
        }
    }
}

function fnEventSelect(eventcode){
    var winSelectEvnt;
    winSelectEvnt = window.open('/admin/eventmanage/event/v5/popup/pop_event_select.asp?mode=relation&eC='+eventcode,'eventselect','width=900,height=600,scrollbars=yes,resizable=yes');
    winSelectEvnt.focus();
}

function fnRelateEventDelete(idx){
    document.ibfrm.idx.value=idx;
    document.ibfrm.submit();
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
<form name="frmEvt" method="post" style="margin:0px;" action="relationevent_process.asp">
<input type="hidden" name="imod" value="RI">
<input type="hidden" name="ecode" value="<%=eCode%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>추천 기획전</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A" id="table">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <!-- for dev msg 기본 작성자를 본인으로 지정, th/td 클릭시 '찾기' 팝업 노출 -->
                    <th>이벤트코드</th>
                    <td>
                        <input type="text" class="formControl formControl150" placeholder="이벤트코드" name="evt_code">
						<button class="btn4 btnBlue1 lMar05" onClick="fnEventSelect(<%=eCode%>);return false;">찾기</button>
                    </td>
                </tr>
            </tbody>
        </table>

        <div class="tableV19BWrap tMar15 tPad25 topLineGrey2">
            <!-- '상품 리스트' -->
            <% If isArray(ArrcEvtInfo) Then %>
            <table class="tableV19A tableV19B tMar10">
                <thead>
                    <tr>
                        <th></th>
                        <th>코드</th>
                        <th>종류</th>
                        <th>이벤트명</th>
                        <th>상태</th>
                        <th>삭제</th>
                    </tr>
                <thead>
                <tbody id="subList">
                    <% For ix = 0 To UBound(ArrcEvtInfo,2) %>
                    <tr>
                        <td>
                            <span class="mdi mdi-equal cBl4 fs20"></span><input type="hidden" name="idx" value="<%=ArrcEvtInfo(0,ix)%>"><input type="hidden" name="viewidx" value="<%=ArrcEvtInfo(2,ix)%>">
                        </td>
                        <td><%=ArrcEvtInfo(1,ix)%></td>
                        <td><%=fnGetCommCodeArrDesc(arrevEntKind, ArrcEvtInfo(3,ix))%></td>
                        <td><%=ArrcEvtInfo(4,ix)%></td>
                        <td><%=fnGetCommCodeArrDesc(arrevEntState, ArrcEvtInfo(5,ix))%></td>
                        <td>
                            <button class="btn4 btnGrey1" onclick="fnRelateEventDelete(<%=ArrcEvtInfo(0,ix)%>);return false;">삭제</button>
                        </td>
                    </tr>
                    <% Next %>
                </tbody>
            </table>
            <% End If %>
        </div>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">저장</button>
	</div>
</div>

</form>
<form method="post" name="ibfrm" action="relationevent_process.asp">
	<input type="hidden" name="idx">
	<input type="hidden" name="imod" value="RD">
    <input type="hidden" name="ecode" value="<%=eCode%>">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->