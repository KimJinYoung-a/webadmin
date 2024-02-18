<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_brandstoryinfo.asp
' Discription : I��(������) �̺�Ʈ �귣�� ���丮 ���� ���� â
' History : 2019.02.12 ������
'###############################################
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
Dim eCode, menuidx, IDX, BrandName, BrandContents
Dim BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin, eFolder, eregdate

eCode = Request("eC")
menuidx = requestCheckVar(Request("menuidx"),10)

IF menuidx <> "" THEN
    set cEvtCont = new ClsMultiContentsMenu
    cEvtCont.FRectEvtCode = eCode
    cEvtCont.FRectIDX = menuidx	'��Ƽ ������ �޴� idx
    cEvtCont.fnGetMultiContentsBrandStory
    IDX = cEvtCont.Fidx
    BrandName = cEvtCont.FBrandName
    BrandContents = cEvtCont.FBrandContents
    BGImage = cEvtCont.FBGImage
	BGColorLeft = cEvtCont.FBGColorLeft
    BGColorRight = cEvtCont.FBGColorRight
	contentsAlign = cEvtCont.FcontentsAlign
	Margin = cEvtCont.FMargin
    set cEvtCont = nothing

    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
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

    if(GetByteLength(frm.eTag.value) > 250){
        alert("Tag�� 250�� �̳��� �ۼ����ּ���");
        frm.eTag.focus();
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
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		eval("document.all."+sName).value = "";
		eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<form name="frmEvt" method="post" style="margin:0px;" action="brandstoryinfo_process.asp">
<% if IDX<>"" then %>
<input type="hidden" name="imod" value="BU">
<% else %>
<input type="hidden" name="imod" value="BI">
<% end if %>
<input type="hidden" name="ebrand">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="menuidx" value="<%=menuidx%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>�귣�� ���丮</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th>�귣��</th>
                    <td>
                        <input type="text" class="formControl formControl150" placeholder="�귣���" name="brandname" id="brandname" value="<%=BrandName%>">
                        <button class="btn4 btnBlue1" onClick="fnSSearchBrandPop(0);return false;">�귣�� ID ã��</button>
                    </td>
                </tr>
                <tr>
                    <th>����</th>
                    <td>
                        <textarea name="dgncomment" rows="10" cols="50" placeholder="�귣�� ���丮�� �Է����ּ���."><%=BrandContents%></textarea>
                    </td>
                </tr>
				<tr>
					<th>��׶��� �̹���</th>
					<td>
                        <input type="hidden" name="BGImage" value="<%=BGImage%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=BGImage%>','BGImage','spanbgimg');return false;">��׶��� �̹��� ���</button>
                        <%IF BGImage <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('BGImage','spanbgimg');return false;">����</button><%END IF%>
                        <div class="previewThumb150W tMar10" id="spanbgimg">
                            <%IF BGImage <> "" THEN %>
                            <%IF BGImage <> "" THEN %><img src="<%=BGImage%>" width="30%" alt=""><%END IF%>
                            <%END IF%>
                        </div>
					</td>
				</tr>
				<tr>
                    <th>��׶��� �÷�</th>
                    <td>
                        ���� : <input type="text" class="formControl formControl150" placeholder="BackGround Color" name="BGColorLeft" id="BGColorLeft" value="<%=BGColorLeft%>">
                        ���� : <input type="text" class="formControl formControl150" placeholder="BackGround Color" name="BGColorRight" id="BGColorRight" value="<%=BGColorRight%>">
                    </td>
                </tr>
                <tr>
                    <th>����</th>
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
                    <th>��� ����</th>
                    <td>
                        <div class="formInline"><input type="text" class="formControl formControl550" maxlength="6" placeholder="��� ����" name="Margin" id="Margin" value="<%=Margin%>"> px</div>
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