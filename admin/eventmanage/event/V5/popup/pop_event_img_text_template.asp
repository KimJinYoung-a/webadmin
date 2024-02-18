<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_img_text_template.asp
' Discription : I��(������) �̺�Ʈ ��ȹ�� ���� �̹��� �ؽ�Ʈ ���ø� ���� â
' History : 2019.10.02 ������
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
    cEvtCont.FRectIDX = menuidx	'��Ƽ ������ �޴� idx
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
function jsEvtSubmit(frm){
    frm.submit();
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
		<h1>�̹��� & HTML</h1>
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
                    <th>�̹��� &#38; HTML</th>
                    <td>
                        <input type="hidden" name="main_mo" value="<%=ImgURL%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=ImgURL%>','main_mo','spanmain_mo');return false;" >���� �̹��� ���</button>
                        <input type="button" value="�̹�������"  onclick="jsManageEventImageNew('<%=eCode%>')" class="btn4 btnBlue1">
                        <%IF ImgURL <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('main_mo','spanmain_mo');return false;">����</button><%END IF%>
                        <div class="previewThumb150W tMar10" id="spanmain_mo">
                            <%IF ImgURL <> "" THEN %><img src="<%=ImgURL%>" alt=""><%END IF%>
                        </div>
                        <div class="tMar15 tPad15 topLineGrey1">
                            <textarea name="tHtml_mo" rows="10" cols="50" placeholder="�ҽ����"><%=BrandContents%></textarea>
                        </div>
<% if winmode="M" then %>
<div class="wrapCode">
<p class="tMar05 cGy1 fs12">*�̹����� �ڵ� ����</p>
<pre><div class="codeArea"><strong>&lt;map name="Mainmap<%=menuidx%>"&gt;&lt;/map&gt;</strong></div></pre>
<p class="tMar05 cGy1 fs12">*�� �˾� ��ũ �ڵ� ����</p>
<pre><div class="codeArea"><strong>&lt;a href="" onclick="fnAPPpopupBrowserURL('���ϸ��� ����', 'https://m.10x10.co.kr/apps/appCom/wish/web2014/offshop/point/mileagelist.asp');return false;" class="mApp"&gt;</strong></div></pre>
<pre><div class="codeArea"><strong>&lt;a href="" onclick="fnAPPpopupBrowserURL('��ȹ��','https://m.10x10.co.kr/apps/appCom/wish/web2014/event/eventmain.asp?eventid=111585');return false;" class="mApp"&gt;</strong></div></pre>
<pre><div class="codeArea"><strong>&lt;a href="" onclick="fnAPPpopupBrowserURL('������������',' https://m.10x10.co.kr/apps/appCom/wish/web2014/my10x10/userinfo/membermodify.asp');return false;" class="mApp"&gt;</strong></div></pre>
</div>
<% else %>
<div class="wrapCode">
<pre><div class="codeArea"><span class="cRd2">PC-WEB ����</span>
&lt;map name="Mainmap<%=menuidx%>"&gt;
<strong>��ǰ������ ��ũ��</strong>
&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoProduct('��ǰ��ȣ');" onfocus="this.blur();"&gt;
<strong>�̺�Ʈ�������� ��ũ��</strong>
&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventMain('�̺�Ʈ�ڵ�');" onfocus="this.blur();"&gt;
<strong>�̺�Ʈ �׷� �������� ��ũ��</strong>
&lt;area shape="rect" coords="0,0,0,0" href="#mapGroup288144" onfocus="this.blur();"&gt;
<strong>�̺�Ʈ �ڸ�Ʈ �̵�</strong>
&lt;area shape="rect" coords="0,0,0,0" href="#commentarea" onfocus="this.blur();"&gt;
<strong>�̺�Ʈ ���� �̵�</strong>
&lt;area shape="rect" coords="0,0,0,0" href="#reviewarea" onfocus="this.blur();"&gt;
<strong>�귣�������� ��ũ��</strong>
&lt;area shape="rect" coords="0,0,0,0" href="javascript:GoToBrandShop('�귣����̵�');" onfocus="this.blur();"&gt;
&lt;/map&gt;
<strong>ī�װ� ��ũ��</strong>
&lt;area shape="rect" coords="0,0,0,0" href="/shopping/category_list.asp?disp=ī�װ���ȣ" onfocus="this.blur();"&gt;
<strong>�̹��� ��� http://webimage.10x10.co.kr/event/XXX/ �� ����Ǿ����ϴ�.</strong>
<strong>*���۾� ������ �ҷ�����(�����) : #[SALEPERCENT]</strong>
</div></pre>
</div>
<% end if %>
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