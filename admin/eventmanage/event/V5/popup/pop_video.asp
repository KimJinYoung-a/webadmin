<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_video.asp
' Discription : I��(������) �̺�Ʈ ���� ����
' History : 2019.02.07 ������ 
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
Dim cEvtCont, ix, VideoFrameReduction, idx
Dim eCode, menuidx, videotype, videoFullLink, videoLink
Dim BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin, eFolder, eregdate

eCode = requestCheckVar(Request("eC"),10)
menuidx = requestCheckVar(Request("menuidx"),10)

IF menuidx <> "" THEN
    set cEvtCont = new ClsMultiContentsMenu
    cEvtCont.FRectEvtCode = eCode
    cEvtCont.FRectIDX = menuidx	'��Ƽ ������ �޴� idx
    cEvtCont.fnGetMultiContentsVideo
    idx = cEvtCont.Fidx
    videoFullLink = cEvtCont.FvideoFullLink
    videoLink = cEvtCont.FvideoLink
    videotype = cEvtCont.Fvideotype
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
	if videotype=" " or isnull(videotype) or videotype="" then videotype="1"
    VideoFrameReduction = Replace(videoFullLink,"width='720'","")
    VideoFrameReduction = Replace(VideoFrameReduction,"height='405'","")
    VideoFrameReduction = Replace(VideoFrameReduction,"height='540'","")
    VideoFrameReduction = Replace(VideoFrameReduction,"height='720'","")
else
    videoFullLink=""
end if
eFolder = eCode
%>
<script type="text/javascript" >
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){
    if(document.frmEvt.videolink.value==""){
		alert("������ URL�� �Է����ּ���.");
		return false;
	}else if (GetByteLength(frm.videolink.value) > 255){
        alert("������ URL�� 255�� ���� �Է��� �ּ���.");
        return false;
    }else{
	    frm.submit();
    }
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
<form name="frmEvt" method="post" style="margin:0px;" action="muticontentsvideo_process.asp">
<% if idx<>"" then %>
<input type="hidden" name="mode" value="VU">
<% else %>
<input type="hidden" name="mode" value="VI">
<% end if %>
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="device" value="W"/>
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
<input type="hidden" name="idx" value="<%=idx%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>����</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th>������ �ּ�</th>
                    <td>
                        <div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="videotype" value="1"<% if videotype="1" then response.write " checked" %>>
								16:9 (720*405)
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="videotype" value="2"<% if videotype="2" then response.write " checked" %>>
								4:3 (720*540)
								<i class="inputHelper"></i>
							</label>
						</div>
                        <div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="videotype" value="3"<% if videotype="3" then response.write " checked" %>>
								1:1 (720*720)
								<i class="inputHelper"></i>
							</label>
						</div>
                        <div class="tMar15 tPad15 topLine">
                            <input type="text" class="formControl" name="videolink" placeholder="��) https://youtu.be/ybbKRICeyV0" value="<%=videoLink%>">
                        </div>
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
<form method="post" name="ibfrm">
	<input type="hidden" name="idx">
	<input type="hidden" name="mode">
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->