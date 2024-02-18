<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_grouptemplateinfo.asp
' Discription : I��(������) �̺�Ʈ ������ ���ø� ����
' History : 2019.02.12 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/goodNoteDiaryCls.asp"-->
<%
dim cEvtCont, idx, mode
dim title, font_color, title_image, bg_color, contents_image, contents_title, eFolder
dim contents, brand_button_color, brand_url, file_url, start_date, end_date, isusing

idx = request("idx")
eFolder = "goodnote"

if idx <> "" then
    mode="edit"
else
    mode="add"
end if

if idx <> "" then
    set cEvtCont = new GoodNoteDiaryStickerCls
    cEvtCont.FRectIDX = idx
	cEvtCont.fnGetStickerContents
    title = cEvtCont.Ftitle
    font_color = cEvtCont.Ffont_color
    title_image = cEvtCont.Ftitle_image
    bg_color = cEvtCont.Fbg_color
    contents_image = cEvtCont.Fcontents_image
    contents_title = cEvtCont.Fcontents_title
    contents = cEvtCont.Fcontents
	brand_button_color = cEvtCont.Fbrand_button_color
    brand_url = cEvtCont.Fbrand_url
	file_url = cEvtCont.Ffile_url
	start_date = cEvtCont.Fstart_date
    end_date = cEvtCont.Fend_date
    isusing = cEvtCont.Fisusing
    set cEvtCont = nothing
end if
if isusing="" then isusing="Y"
%>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<link rel="stylesheet" href="http://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
<script>
$(function(){
    $("#datepicker").datepicker({
        changeMonth: true, 
        changeYear: true,
        minDate: '-50y',
        nextText: '���� ��',
        prevText: '���� ��',
        yearRange: 'c-50:c+20',
        showButtonPanel: true, 
        currentText: '���� ��¥',
        closeText: '�ݱ�',
        dateFormat: "yy-mm-dd",
        showAnim: "slide",
        showMonthAfterYear: true, 
        dayNamesMin: ['��', 'ȭ', '��', '��', '��', '��', '��'],
        monthNamesShort: ['1��','2��','3��','4��','5��','6��','7��','8��','9��','10��','11��','12��']
    });
    $("#datepicker2").datepicker({
        changeMonth: true, 
        changeYear: true,
        minDate: '-50y',
        nextText: '���� ��',
        prevText: '���� ��',
        yearRange: 'c-50:c+20',
        showButtonPanel: true, 
        currentText: '���� ��¥',
        closeText: '�ݱ�',
        dateFormat: "yy-mm-dd",
        showAnim: "slide",
        showMonthAfterYear: true, 
        dayNamesMin: ['��', 'ȭ', '��', '��', '��', '��', '��'],
        monthNamesShort: ['1��','2��','3��','4��','5��','6��','7��','8��','9��','10��','11��','12��']
    });
});

function fnGoodNoteSubmit(frm){
    if(frm.title.value==""){
        alert("������ �Է����ּ���.");
        frm.title.focus();
        return false;
    }
    if(frm.bg_color.value==""){
        alert("������ �Է����ּ���.");
        frm.bg_color.focus();
        return false;
    }
    if(frm.contents_title.value==""){
        alert("������ ������ �Է����ּ���.");
        frm.contents_title.focus();
        return false;
    }
    if(frm.contents.value==""){
        alert("������ ������ �Է����ּ���.");
        frm.contents.focus();
        return false;
    }
    if(frm.brand_button_color.value==""){
        alert("�귣�� ��ư���� �Է����ּ���.");
        frm.brand_button_color.focus();
        return false;
    }
    if(frm.brand_url.value==""){
        alert("�귣�� URL�� �Է����ּ���.");
        frm.brand_url.focus();
        return false;
    }
    if(frm.file_url.value==""){
        alert("���� �ٿ�ε� URL�� �Է����ּ���.");
        frm.file_url.focus();
        return false;
    }
    if(frm.start_date.value==""){
        alert("�������� �������ּ���.");
        frm.start_date.focus();
        return false;
    }
    if(frm.end_date.value==""){
        alert("�������� �������ּ���.");
        frm.end_date.focus();
        return false;
    }
}

function jsSetImg(sFolder, sImg, sName, sSpan , sOpt){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan+'&sOpt='+sOpt,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<form name="frmEvt" method="post" style="margin:0px;" action="goodnote_process.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>��ƼĿ �߰�</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A" id="table">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th>����</th>
                    <td>
                        ���� : <input type="text" class="formControl formControl550" placeholder="������ �Է����ּ���." name="title" id="title" value="<%=title%>"><br>
                        ��Ʈ �÷� : <input type="text" class="formControl formControl150" placeholder="��Ʈ �÷��� �Է����ּ���." name="font_color" id="font_color" value="<%=font_color%>">
                    </td>
                </tr>
				<tr>
					<th>Ÿ��Ʋ �̹���</th>
					<td>
                        <input type="hidden" name="title_image" value="<%=title_image%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=title_image%>','title_image','span_title_image');return false;">Ÿ��Ʋ �̹��� ���</button>
                        <%IF title_image <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('title_image','span_title_image');return false;">����</button><%END IF%>
                        <div class="previewThumb150W tMar10" id="span_title_image">
                            <%IF title_image <> "" THEN %>
                            <%IF title_image <> "" THEN %><img src="<%=title_image%>" width="30%" alt=""><%END IF%>
                            <%END IF%>
                        </div>
					</td>
				</tr>
				<tr>
                    <th>��׶��� �÷�</th>
                    <td>
                        <input type="text" class="formControl formControl150" placeholder="��׶��� �÷��� �Է����ּ���." name="bg_color" id="bg_color" value="<%=bg_color%>">
                    </td>
                </tr>
                <tr>
					<th>������ �̹���</th>
					<td>
                        <input type="hidden" name="contents_image" value="<%=contents_image%>">
                        <button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=contents_image%>','contents_image','span_contents_image');return false;">Ÿ��Ʋ �̹��� ���</button>
                        <%IF contents_image <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('contents_image','span_contents_image');return false;">����</button><%END IF%>
                        <div class="previewThumb150W tMar10" id="span_contents_image">
                            <%IF contents_image <> "" THEN %>
                            <%IF contents_image <> "" THEN %><img src="<%=contents_image%>" width="30%" alt=""><%END IF%>
                            <%END IF%>
                        </div>
					</td>
				</tr>
                <tr>
                    <th>������ ����</th>
                    <td>
                        <input type="text" class="formControl formControl550" placeholder="������ ������ �Է����ּ���." name="contents_title" id="contents_title" value="<%=contents_title%>">
                    </td>
                </tr>
                <tr>
                    <th>������ ����</th>
                    <td>
                        <textarea name="contents" rows="10" cols="50" placeholder="������ �Է����ּ���."><%=contents%></textarea>
                    </td>
                </tr>
                <tr>
                    <th>�귣�� ��ư �÷�</th>
                    <td>
                        <input type="text" class="formControl formControl150" placeholder="��ư �÷��� �Է����ּ���." name="brand_button_color" id="brand_button_color" value="<%=brand_button_color%>">
                    </td>
                </tr>
                <tr>
                    <th>�귣�� URL</th>
                    <td>
                        <input type="text" class="formControl formControl550" placeholder="�귣�� URL�� �Է����ּ���." name="brand_url" id="brand_url" value="<%=brand_url%>">
                    </td>
                </tr>
                <tr>
                    <th>�ٿ�ε� ���� URL</th>
                    <td>
                        <input type="text" class="formControl formControl550" placeholder="�ٿ�ε� ���� URL�� �Է����ּ���." name="file_url" id="file_url" value="<%=file_url%>">
                    </td>
                </tr>
                <tr>
                    <th>�Ⱓ</th>
                    <td>
                        <input type="text" class="formControl formControl150" placeholder="�������� �������ּ���." name="start_date" value="<%=start_date%>" id="datepicker">
                        ~ <input type="text" class="formControl formControl150" placeholder="�������� �������ּ���." name="end_date" value="<%=end_date%>" id="datepicker2">
                    </td>
                </tr>
                <tr>
                    <th>�������</th>
                    <td>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="isusing" value="Y"<% if isusing="Y" then response.write " checked"%>>
                                ���
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="isusing" value="N"<% if isusing="N" then response.write " checked"%>>
                                �̻��
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onclick="self.close();">���</button>
		<button class="btn4 btnBlue1" onclick="fnGoodNoteSubmit(this.form);">����</button>
	</div>
</div>
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->