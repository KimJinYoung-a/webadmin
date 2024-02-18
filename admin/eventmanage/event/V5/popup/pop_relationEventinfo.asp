<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_relationEventinfo.asp
' Discription : I��(������) �̺�Ʈ ���� �̺�Ʈ ����
' History : 2019.02.27 ������
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
'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
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
            alert("�ִ� ��� ������ �ʰ��߽��ϴ�.");
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
	//�巡��
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
		<h1>��õ ��ȹ��</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A" id="table">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <!-- for dev msg �⺻ �ۼ��ڸ� �������� ����, th/td Ŭ���� 'ã��' �˾� ���� -->
                    <th>�̺�Ʈ�ڵ�</th>
                    <td>
                        <input type="text" class="formControl formControl150" placeholder="�̺�Ʈ�ڵ�" name="evt_code">
						<button class="btn4 btnBlue1 lMar05" onClick="fnEventSelect(<%=eCode%>);return false;">ã��</button>
                    </td>
                </tr>
            </tbody>
        </table>

        <div class="tableV19BWrap tMar15 tPad25 topLineGrey2">
            <!-- '��ǰ ����Ʈ' -->
            <% If isArray(ArrcEvtInfo) Then %>
            <table class="tableV19A tableV19B tMar10">
                <thead>
                    <tr>
                        <th></th>
                        <th>�ڵ�</th>
                        <th>����</th>
                        <th>�̺�Ʈ��</th>
                        <th>����</th>
                        <th>����</th>
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
                            <button class="btn4 btnGrey1" onclick="fnRelateEventDelete(<%=ArrcEvtInfo(0,ix)%>);return false;">����</button>
                        </td>
                    </tr>
                    <% Next %>
                </tbody>
            </table>
            <% End If %>
        </div>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">���</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">����</button>
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