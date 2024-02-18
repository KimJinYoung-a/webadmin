<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_mobile_slide.asp
' Discription : ����� slide insert
' History : 2016-02-16 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim eCode , title , eFolder , topimg , btmimg , topaddimg 'floating img
Dim videoSize, videoLink '������ �߰�
Dim slideimg
Dim mode , idx , strSql , sqlStr , isarrow

	eCode = requestCheckvar(request("eC"),16)
	title = "�����̵� ��� �˾�(M)"

	eFolder = eCode

	If eCode <> "" Then 
		strSql = "SELECT idx , topimg , btmimg , topaddimg, videosize, videolink , isarrow " + vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_event_slide_template where evt_code = '"& eCode &"' and device = 'M'" 
		rsget.Open strSql,dbget,1
		IF Not rsget.Eof Then
			idx			= rsget("idx")
			topimg		= rsget("topimg")
			btmimg		= rsget("btmimg")
			topaddimg	= rsget("topaddimg")
			videosize	= rsget("videosize")
			videolink	= rsget("videolink")
			isarrow		= rsget("isarrow")
		End If
		rsget.close()
	End If 

	If idx = "" Then
		mode = "I"
	Else
		mode = "U"
	End If

%>
<!-- #include virtual="/admin/lib/popheaderslide.asp"-->
<link rel="stylesheet" type="text/css" href="http://webadmin.10x10.co.kr/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="http://webadmin.10x10.co.kr/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
.vod-wrap .vod {overflow:hidden; position:relative; width:100%; height:100%; padding-bottom:100%; /padding-bottom:70.4%;/}
.vod-wrap .vod iframe {position:absolute; top:0; left:0; bottom:0; width:100%; height:100%;}
.shape-rtgl .vod {padding-bottom:56.25%;}
</style>
<script type="text/javascript" src="http://m.10x10.co.kr/lib/js/jquery.swiper-3.1.2.min.js"></script>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ���縦 �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.action="pop_slide_proc.asp";
		document.frmList.submit();
	}
}

//'������ ����
function slideimgDel(v){
	if (confirm("��ǰ�� �����˴ϴ� ���� �Ͻðڽ��ϱ�?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.submit();
	}
}

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<script type="text/javascript">
$(function(){
	dfslide(); //���� �����̵� �ε�
	
	//�巡��
	$( "#subList").sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});

});

function dfslide(){
	var str = $.ajax({
		type: "GET",
		url: "pop_mobile_slide_ajax.asp",
		data: "eC=<%=eCode%>",
		dataType: "text",
		async: false
		}).responseText;
	if (str != ""){
		$('#preview_ajax').append(str);
	}
}

//��ũ������
function showDrop(){
	$(".selectLink ul").show();
}

//�����Է�
function populateTextBox(v){
	var val = v;
	$("#mlinkurl").val(val);
	$("#linkurl").val(val);
	$(".selectLink ul").css("display","none");
}

function linkcopy(){
	var val = $("#mlinkurl").val();
	$("#linkurl").attr("value",val);
	$(".selectLink ul").css("display","none");
}

function closel(){
	$(".selectLink ul").css("display","none");
}

function simgsubmit(){
	//�����̵��̹����Է�
	var frm = document.slideimgfrm;
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}
function mimgsubmit(){
	//��Ÿ ���� �÷���
	var frm = document.slidefrm;
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function videosizeins(v){
	var val = v;
	$("#videosize").val(v);
}

function chkisarrow(v){
	var val = v;
	$("#isarrow").val(v);
}

function videolinkins(v){
	var val = v;
	$("#videolink").val(v);
}


$(function(){
    var currentPosition = parseInt($("#preview_ajax").css("top"));
    $(window).scroll(function() {  
         var position = $(window).scrollTop(); // ���� ��ũ�ѹ��� ��ġ���� ��ȯ�մϴ�.
		 if (position > 0){
			$("#preview_ajax").stop().animate({"top":position+"px"},500);
		 }else{
			$("#preview_ajax").stop().animate({"top":position+currentPosition+"px"},500);  
		 }
         
    });
});  
</script>
</head>
<body>
<div class="slideRegister adminMob">
	<h1>�����̵� �� ������ ��� (MOBILE)</h1>
	<%'�����̵� ȭ�� �ҷ�����%>
	<div class="preview <% If topaddimg <> "" Then %>txtFix<% End If %>" id="preview_ajax"></div>
	<%'�����̵� ȭ�� �ҷ�����%>
	<div class="register">
		<h2>������ ���</h2>
		<dl>
			<dt>- ���<span>(��������)</span></dt>
			<dd><input type="button" value="�̹��� �ҷ�����" name="mtopimg" onClick="jsSetImg('<%=eFolder%>','','topimg','spantopimg')"/><%IF topimg <> "" THEN %><div><br/><%=topimg%>&nbsp;<a href="javascript:jsDelImg('topaddimg','spantopaddimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a></div><%END IF%></dd>
		</dl>
		<dl>
			<dt>- �����̵� ȭ��ǥ ��� ����</dt>
			<dd>
				<input type="radio" value="1" name="isarrowyn" id="arrowYes" <%=chkiif(isarrow="" or isarrow = 1 , "checked" , "")%> onclick="chkisarrow(this.value);"/> <label for="arrowYes">�����</label>&nbsp;&nbsp;&nbsp;<input type="radio" value="0" name="isarrowyn" id="arrowNo" <%=chkiif(isarrow = 0, "checked" , "")%> onclick="chkisarrow(this.value);"/> <label for="arrowNo">������</label>
			</dd>
		</dl>
		<dl>
			<dt>- �����̵� <span>(�ʼ�/3~10������ ���)</span></dt>
			<dd>
				<p class="floatImg">�÷��� �̹���(width:750px, png�� ���) :<input class="button" type="button" value="�̹��� �ҷ�����" name="mtopimg" onclick="jsSetImg('<%=eFolder%>','','topaddimg','spantopaddimg')"/></p>
				<div class="insertImg">
					<h3>�����̵� �̹��� ���<br/><span style="color:#c80a0a;line-height:2;">�� ���ϴ� �̹��� �켱 ����� �����̵� �̹����� ������ּ���. �Ǵ� �����̵� �켱 ����� ���ϴ� �̹����� ��� ���ּ���. ��<br/>�� ��Ϲ�ư Ŭ���� ���� Viewer�� �ݿ� �˴ϴ�.��</span></h3>
					<table class="tbType1 listTb tMar10">
						<colgroup>
							<col width="7%" /><col /><col width="42%" /><col width="7%" /><col width="12%" />
						</colgroup>
						<tbody>
						<tr>
							<td></td>
							<td>
								<div id="spanslideimg"></div>
								<input class="button" type="button" value="�̹��� �ҷ�����" name="mslideimg" onClick="jsSetImg('<%=eFolder%>','','slideimg','spanslideimg');"/>
							</td>
							<td>
								<div class="selectLink">
									<input type="text" value="��ũ�� �Է�(����)" onclick="showDrop();" id="mlinkurl" onkeyup="linkcopy();" />
									<ul style="display:none;">
										<li onclick="populateTextBox('');">���þ���</li>
										<li onclick="populateTextBox('#group�׷��ڵ�');">#group�׷��ڵ�</li>
										<li onclick="populateTextBox('/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�');">/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�</li>
										<li onclick="populateTextBox('/category/category_itemprd.asp?itemid=��ǰ�ڵ�');">/category/category_itemprd.asp?itemid=��ǰ�ڵ� (O)</li>
										<li onclick="populateTextBox('/category/category_list.asp?disp=ī�װ�');">/category/category_list.asp?disp=ī�װ�</li>
										<li onclick="populateTextBox('/street/street_brand.asp?makerid=�귣����̵�');">/street/street_brand.asp?makerid=�귣����̵�</li>
									</ul>
								</div>
							</td>
							<td><input type="text" value="0" /></td>
							<td>
								<input type="button" class="btn" value="���" onclick="simgsubmit();">
							</td>
						</tr>
						</tbody>
					</table>
				</div>
				<form name="frmList" method="POST" action="" style="margin:0;">
				<input type="hidden" name="mode" value="SU"/>
				<input type="hidden" name="device" value="M"/>
				<input type="hidden" name="eventid" value="<%=eCode%>"/>
				<div class="tMar20">
					<p>
						<input type="button" class="btn" value="��ü ����" onclick="chkAllItem();">
						<input type="button" class="btn" value="���� ����" onClick="saveList();" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
					</p>
					<table class="tbType1 listTb tMar10">
						<colgroup>
							<col width="7%" /><col /><col width="42%" /><col width="7%" /><col width="12%" />
						</colgroup>
						<thead>
						<tr>
							<th><input type="checkbox" onclick="chkAllItem();"/></th>
							<th>�̹���</th>
							<th>��ũ(����)</th>
							<th>����</th>
							<th>��뿩��</th>
						</tr>
						</thead>
						<tbody id="subList">
						<% 
							If eCode <> "" Then 

							sqlStr = "SELECT idx , slideimg , linkurl , sorting , isusing " + vbcrlf
							sqlStr = sqlStr & " from db_event.[dbo].[tbl_event_slide_addimage] where evt_code = '"& eCode &"' and device = 'M' " 
							sqlStr = sqlStr & " order by sorting asc , idx asc " 
							rsget.Open sqlStr,dbget,1
							if Not(rsget.EOF or rsget.BOF) Then
								Do Until rsget.eof
						%>
						<tr class="<%=chkIIF(rsget("isusing")="N","bgGry1","")%>">
							<td><input type="checkbox" name="chkIdx" value="<%=rsget("idx")%>" /></td>
							<td><img src="<%=rsget("slideimg")%>" style="width:100px;" /></td>
							<td class="lt"><input type="text" style="width:400px;" name="linkurl<%=rsget("idx")%>" value="<%=rsget("linkurl")%>"/></td>
							<td><input type="text" value="<%=rsget("sorting")%>" name="sort<%=rsget("idx")%>"/></td>
							<td>
								<span><input type="radio" <%=chkIIF(rsget("isusing")="Y","checked","")%> name="use<%=rsget("idx")%>" value="Y"/> Y</span>
								<span class="lMar10"><input type="radio" <%=chkIIF(rsget("isusing")="N","checked","")%> name="use<%=rsget("idx")%>" value="N"/> N</span>
								<br/><input type="button" class="btn" value="����" onclick="slideimgDel(<%=rsget("idx")%>);">
							</td>
						</tr>
						<% 
								rsget.movenext
								Loop
							End If
							rsget.close

							End If
						%>
						</tbody>
					</table>
					<p style="text-align:right;margin-top:5px;">
						<input type="button" class="btn" value="���� ����" onClick="saveList();" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
					</p>
				</div>
				</form>
			</dd>
		</dl>
		<dl>
			<dt>- �ϴ�<span>(��������)</span></dt>
			<dd><input class="button" type="button" value="�̹��� �ҷ�����" name="mbtmimg" onClick="jsSetImg('<%=eFolder%>','','btmimg','spanbtmimg')" /><%IF btmimg <> "" THEN %><div><br/><%=btmimg%>&nbsp;<a href="javascript:jsDelImg('btmimg','spanbtmimg');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a></div><%END IF%></dd>
		</dl>
		<dl>
			<dt>- ������<span>(��������)</span></dt>
			<dd>
				<span style="color:#c80a0a;line-height:2;">�� ��޿�/��Ʃ�� ������ ��ũ�� �����մϴ�.��</span>
				<table class="tbType1 listTb tMar10">
					<colgroup>
						<col width="30%" /><col /><col width="42%" />
					</colgroup>
					<thead>
					<tr>
						<th>������</th>
						<th>��ũ</th>
					</tr>
					</thead>
					<tbody>
					<tr>
						<td><input type="radio" name="videosizechk" id="videosizechk" value="N" <%=chkIIF(videosize="N","checked","")%> onclick="videosizeins(this.value);">������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="videosizechk" value="W" <%=chkIIF(videosize="W","checked","")%> onclick="videosizeins(this.value);">���̵�</td>
						<td class="lt"><input type="text" style="width:500px;" name="videolinktxt" id="videolinktxt" value="<%=videolink%>" onkeyup="videolinkins(this.value);"/></td>
					</tr>
					</tbody>
				</table>
				<!--p style="text-align:right;margin-top:5px;">
					<input type="button" class="btn" value="������ ����" onClick="saveList();" title="�������� �����մϴ�.">
				</p-->			
			</dd>
		</dl>		
		<div class="btnArea">
			<input type="image" src="http://webadmin.10x10.co.kr/images/icon_save.gif" alt="����" onclick="mimgsubmit();"/>
			<a href=""><img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" alt="���" /></a>
		</div>
	</div>
</div>
<form name="frmdel" method="POST" action="pop_slide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SD"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="chkIdx" />
</form>
<form name="slideimgfrm" method="post" action="pop_slide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SI"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="slideimg" value=""/>
<input type="hidden" name="linkurl" id="linkurl" value=""/>
</form>
<form name="slidefrm" method="post" action="pop_slide_proc.asp" style="margin:0px;">
<input type="hidden" name="idx" value="<%=idx%>"/>
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="<%=mode%>"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="topimg" value="<%=topimg%>"/>
<input type="hidden" name="btmimg" value="<%=btmimg%>"/>
<input type="hidden" name="topaddimg" value="<%=topaddimg%>"/>
<input type="hidden" name="videosize" id="videosize" value="<%=videosize%>"/>
<input type="hidden" name="videolink" id="videolink" value="<%=videolink%>"/>
<input type="hidden" name="isarrow" id="isarrow" value="<%=isarrow%>"/>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->