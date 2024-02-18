<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : main_manager.asp
' Discription : �̹��� ���� ��ũ
' History : 2019.08.06 ������ : �ű��ۼ�
'          2019.12.09 ������ - �̺�Ʈ ������ �߰�
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/event/imageLinkCls.asp" -->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
Dim isusing, fixtype, validdate, prevDate
Dim idx, poscode, reload, gubun, edid
Dim eCode, menuidx, cEvtCont, winmode
Dim BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin, eFolder, eregdate

	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	gubun = request("gubun")

	isusing = request("isusing")
	fixtype = request("fixtype")
	validdate= request("validdate")
	prevDate = request("prevDate")

	eCode = request("eC")
    menuidx = Request("menuidx")
	winmode = Request("wm")
	if winmode="" then winmode="M"
	if idx="" then idx=0

	if reload="on" then
	    response.write "<script>opener.location.reload(); window.close();</script>"
	    dbget.close()	:	response.End
	end if

	dim oLinkContents
		set oLinkContents = new CimageLink
		oLinkContents.FRectIdx = menuidx
		oLinkContents.FRectDevice = winmode
		oLinkContents.GetOneContents

	If gubun = "" Then
		gubun = "index"
	End If

	IF menuidx <> "" THEN
		set cEvtCont = new ClsMultiContentsMenu
		cEvtCont.FRectEvtCode = eCode
		cEvtCont.FRectIDX = menuidx	'��Ƽ ������ �޴� idx
		cEvtCont.fnGetMultiContentsMenu
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
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
    window.document.domain = "10x10.co.kr";
	function SaveMainContents(frm){

		if (frm.title.value==""){
	        alert('Ÿ��Ʋ�� �Է����ּ���.');
	        frm.title.focus();
	        return;
	    }

		if (frm.Link_Image.value==""){
	        alert('�̹����� ���ε� �ϼ���.');
	        frm.Link_Image.focus();
	        return;
	    }

	    if (confirm('���� �Ͻðڽ��ϱ�?')){
	        frm.submit();
	    }
	}

	function ChangeLinktype(comp){
	    if (comp.value=="M"){
	       document.all.link_M.style.display = "";
	       document.all.link_L.style.display = "none";
	    }else{
	       document.all.link_M.style.display = "none";
	       document.all.link_L.style.display = "";
	    }
	}

	//function getOnLoad(){
	//    ChangeLinktype(frmcontents.linktype.value);
	//}

	//window.onload = getOnLoad;

	function ChangeGubun(comp){
	    location.href = "?gubun=<%=gubun%>&poscode=" + comp.value;
	    // nothing;
	}


	function ChangeGroupGubun(comp){
	    location.href = "?gubun=" + comp.value;
	    // nothing;
	}

	function cultureloadpop(){
		winLast = window.open('pop_culturelist.asp','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}

	//�����ڵ� ����
	function selColorChip(bg,cd) {
		var i;
		document.frmcontents.BGColor.value= bg;
		for(i=1;i<=11;i++) {
			document.all("cline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("cline0").bgColor='#DD3300';
		else document.all("cline"+cd).bgColor='#DD3300';
	}

	//-- jsLastEvent : ���� �̺�Ʈ �ҷ����� --//
	function jsLastEvent(num){
	  winLast = window.open('pop_event_lastlist.asp?num='+num,'pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
	}

	//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
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
<form name="frmcontents" method="post" action="imagelink_process.asp" onsubmit="return false;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
<input type="hidden" name="device" value="<%=winmode%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>�̹��� ��ũ</h1>
	</div>
	<div class="popContV19">
		<div class="tabV19">
			<ul>
				<li class="<% if winmode="M" then %>selected<% end if %>"><a href="pop_event_imagelink.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>&wm=M">Mobile / App</a></li>
				<li class="<% if winmode="W" then %>selected<% end if %>"><a href="pop_event_imagelink.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>&wm=W">PC</a></li>
			</ul>
		</div>
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th>Idx</th>
                    <td>
						<% if oLinkContents.FOneItem.Fidx<>"" then %>
						<%= oLinkContents.FOneItem.Fidx %>
						<input type="hidden" name="idx" value="<%= oLinkContents.FOneItem.Fidx %>">
						<% end if %>
                    </td>
                </tr>
                <tr>
                    <th>�̹���</th>
                    <td>
                        <input type="hidden" name="Link_Image" value="<%=oLinkContents.FOneItem.Fimage%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('evtlinkimage','<%=oLinkContents.FOneItem.FImage%>','Link_Image','newlinkimg');return false;">���� �̹��� ���</button>
                        <%IF oLinkContents.FOneItem.FImage <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('Link_Image','newlinkimg');return false;">����</button><%END IF%>
                        <div class="previewThumb150W tMar10" id="newlinkimg">
                            <%IF oLinkContents.FOneItem.FImage <> "" THEN %><img src="<%=oLinkContents.FOneItem.FImage%>" alt=""><%END IF%><br>
						<button class="btn4 btnBlue1" onclick="window.open('pop_event_imagemap.asp?menupos=<%=menupos%>&idx=<%=oLinkContents.FOneItem.Fidx%>','imagelinkedit','width=950,height=780,scrollbars=yes,resizable=yes');return false;">�̹��� �� ���</button>
                        </div>
                    </td>
                </tr>
                <tr>
                    <th>�������</th>
                    <td>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="Isusing" value="Y"<% if oLinkContents.FOneItem.FIsusing="Y" Or oLinkContents.FOneItem.FIsusing="" then %> checked<% end if %>>
                                �����
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="Isusing" value="N"<% if oLinkContents.FOneItem.FIsusing="N" then %> checked<% end if %>>
                                ������
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                    </td>
                </tr>
				<% If oLinkContents.FOneItem.Fadminid<>"" Then %>
                <tr>
                    <th>�۾���</th>
                    <td>
						�۾��� : <%=oLinkContents.FOneItem.Fadminid %><br>
						�����۾��� : <%=oLinkContents.FOneItem.Flastadminid %>
                    </td>
                </tr>
				<% End If %>
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
		<button class="btn4 btnBlue1" onClick="SaveMainContents(frmcontents);return false;">����</button>
	</div>
</div>
</form>
<%
set oLinkContents = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->