<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_setting.asp
' Discription : I��(������) �̺�Ʈ ��ȹ�� ���� ���� â
' History : 2019.01.25 ������
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
Dim cEvtCont
Dim eCode, nocate, ebrand, etag, DispCate, maxDepth, eisort


eCode = Request("eC")
maxDepth = 2 '����ī�װ� 2depth���� �����ش�

IF eCode <> "" THEN
    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
    '�̺�Ʈ ȭ�鼳�� ���� ��������
	cEvtCont.fnGetEventDisplay
    ebrand = cEvtCont.FEBrand
    etag = db2html(cEvtCont.FETag)
	eisort = cEvtCont.FEISort
    DispCate = cEvtCont.FEDispCate
	'���� ��� �̺�Ʈ �׸� ����
	cEvtCont.fnGetEventMDThemeInfo
    nocate = cEvtCont.Fnocate
	
    set cEvtCont = nothing
else
    ebrand=""
    etag=""
end if
if eisort="" or isnull(eisort) then eisort=3
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
    if(frm.disp.value=="0" && !document.getElementById("nocate").checked){
        alert("ī�װ��� �����ϰų� ī�װ� ���þ����� üũ���ּ���.");
        return false;
    }
    if(frm.eTag.value==""){
        alert("Tag�� �Է����ּ���.");
        frm.eTag.focus();
        return false;
    }
    if(GetByteLength(frm.eTag.value) > 250){
        alert("Tag�� 250�� �̳��� �ۼ����ּ���");
        frm.eTag.focus();
        return false;
    }
	frm.submit();
}

function TnFavSearchTxt(){
    var winpop = window.open("http://61.252.133.17:5601/app/kibana#/dashboard/5c9d9970-ef60-11e6-9fb4-f3d99fd9206d?_g=(refreshInterval:(display:Off,pause:!f,value:0),time:(from:now-5h%2Fh,mode:quick,to:now))&_a=(filters:!(),options:(darkTheme:!f),panels:!((col:1,id:ca566510-ef5f-11e6-9fb4-f3d99fd9206d,panelIndex:1,row:1,size_x:3,size_y:5,type:visualization),(col:1,id:'%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4(MOB)',panelIndex:2,row:6,size_x:3,size_y:5,type:visualization),(col:1,id:'%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4(APP)',panelIndex:3,row:11,size_x:3,size_y:5,type:visualization),(col:4,id:'%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4-%EC%8B%9C%EA%B0%84%EB%8C%80%EB%B3%84(MOB)',panelIndex:4,row:6,size_x:9,size_y:5,type:visualization),(col:4,id:d06ee1e0-ef62-11e6-9fb4-f3d99fd9206d,panelIndex:5,row:1,size_x:9,size_y:5,type:visualization),(col:4,id:c7604a10-1aa2-11e7-b3b2-cb4977e75f0e,panelIndex:6,row:11,size_x:9,size_y:5,type:visualization)),query:(query_string:(analyze_wildcard:!t,query:'*')),title:'0005.%20%EC%9D%B8%EA%B8%B0%EA%B2%80%EC%83%89%EC%96%B4',uiState:(P-1:(vis:(params:(sort:(columnIndex:!n,direction:!n)))),P-2:(vis:(params:(sort:(columnIndex:!n,direction:!n)))),P-3:(vis:(params:(sort:(columnIndex:!n,direction:!n))))))",'winpop2','width=1450,height=800,scrollbars=yes,resizable=yes');
    winpop.focus();
}

function jsDelID(){ 
    document.frmEvt.ebrand.value = "";
}

function jsAddByte(obj){ 
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

        if (textLen >= 500){
            realText = realText.substr(0,i);
            obj.value = realText;
            break;
        }
    }
	$("#Tag").html(textLen);
}

//�귣�� ID �˻� �˾�â
function fnSearchBrandID(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/admin/eventmanage/event/v5/popup/popBrandSearchID.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}
</script>
<form name="frmEvt" method="post" style="margin:0px;" action="settinginfo_process.asp">
<input type="hidden" name="imod" value="SU">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>��ȹ������</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
                    <th>ī�װ�</th>
                    <td>
						<!-- #include virtual="/common/module/dispEventCateSelectBoxDepth.asp"-->
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="checkbox" class="formCheckInput" name="nocate" id="nocate" value="Y"<% If nocate="Y" Then Response.write " checked"%>>
                                ī�װ� ���� ����
                                <i class="inputHelper"></i>
                            </label>
                            <span class="mdi mdiBlue mdi-help-circle-outline cBl4"></span>
                        </div>
                    </td>
                </tr>
				<tr>
                    <th>�귣��</th>
                    <td>
                        <input type="text" class="formControl formControl150" name="ebrand" value="<%=ebrand%>" placeholder="�귣���">
                        <button class="btn4 btnBlue1 lMar05" onclick="fnSearchBrandID(this.form.name,'ebrand');return false;">�귣�� ID ã��</button>
                        <button class="btn4 btnGrey1 lMar05" onClick="jsDelID();return false;">����</button>
                    </td>
                </tr>
                <tr>
                    <th>�±�<p class="cGy2 fs13">(250�� �̳�)</p></th>
                    <td class="overHidden">
                        <textarea name="eTag" rows="10" cols="50" placeholder="�α��±׸� �Է��غ��ƿ� :)" OnKeyUp="jsAddByte(this);"><%=etag%></textarea>
                        <p class="ftLt tMar20 cGy1 fs12"><span class="cPk2 vBtm" id="Tag">50</span><span class="cPk2 vBtm">byte</span>/500byte</p>
                        <button class="ftRt btn4 btnBlue1 tMar10" onClick="TnFavSearchTxt();return false;">�ǽð� �α� �˻��� ����</button>
						<script type="text/javascript">
							jsAddByte(frmEvt.eTag);
						</script>
                    </td>
                </tr>
            </tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">���</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">����</button>
	</div>
</div>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->