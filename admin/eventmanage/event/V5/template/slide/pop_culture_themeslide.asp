<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_culture_themeslide.asp
' Discription : ����� slide insert
' History : 2019-08-28 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim eCode , title , eFolder , topimg , btmimg , topaddimg 'floating img
Dim slideimg, menuidx
Dim mode , idx , strSql , sqlStr, smode, saveafter
smode = requestCheckvar(request("smode"),16)
saveafter = requestCheckvar(request("saveafter"),2)
eCode = requestCheckvar(request("eC"),16)
menuidx = requestCheckvar(request("menuidx"),10)

If saveafter <>"" Then
Response.write "<script>opener.document.location.reload();</script>"
End If

If smode="SU" Then
Response.write "<script>self.close();</script>"
Response.end
End If

%>
<!-- #include virtual="/admin/lib/popheaderslide.asp"-->
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
</style>
<script type="text/javascript" src="/admin/common/lib/js/jquery.swiper-3.1.2.min.js"></script>
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
	chkAllItem();
	var chk=0;
	$("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ���縦 �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.action="pop_themeslide_proc.asp";
		document.frmList.submit();
	}
}

//'������ ����
function slideimgDel(v){
	if (confirm("�����˴ϴ� ���� �Ͻðڽ��ϱ�?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.submit();
	}
}

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/event/v5/lib/pop_event_multi_uploadimgV5.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>&yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan+'&wid=7500&hei=5280','popImg','width=370,height=350');
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

function dfslide(){
	var str = $.ajax({
		type: "GET",
		url: "pop_mobile_themeslide_ajax.asp",
		data: "eC=<%=eCode%>&menuidx=<%=menuidx%>",
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

function fnUploadTypeSelect(type){
	if(type=="I"){
		$("#utypeI").show();
		$("#utypeV").hide();
	}
	else{
		$("#utypeI").hide();
		$("#utypeV").show();
	}
}

function fnuploadvideo(){
	if(document.frmvideo.videolink.value==""){
		alert("������ URL�� �Է����ּ���.");
		return false;
	}
	else{
		document.frmvideo.submit();
	}
}

function fnDeleteImg() {
	var chk=0;
	var idxarr;
	$("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")){
			chk++;
			if(chk==1){
				idxarr = $(this).val();
			}
			else{
				idxarr = idxarr + "," + $(this).val();
			}
		}
	});
	if(chk==0) {
		alert("���� �̹����� �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� �̹����� �����Ͻðڽ��ϱ�?")) {
		document.frmchkdel.chkIdx.value=idxarr;
		document.frmchkdel.submit();
	}
}
</script>
</head>
<body>
<div class="slideRegister adminMob">
	<h1>�����̵� ���</h1>
	<%'�����̵� ȭ�� �ҷ�����%>
	<div class="preview <% If topaddimg <> "" Then %>txtFix<% End If %>" id="preview_ajax"></div>
	<%'�����̵� ȭ�� �ҷ�����%>
	<div class="register">
		<h2>������ ���</h2>
		<dl>
			<dt>- �����̵� <span>(�ʼ�/3~10������ ���)</span></dt>
			<dd>
				<div class="insertImg">
					<h3>�����̵� �̹��� ���<br/><span style="color:#c80a0a;line-height:2;">�� ���ϴ� �̹��� �켱 ����� �����̵� �̹����� ������ּ���. �Ǵ� �����̵� �켱 ����� ���ϴ� �̹����� ��� ���ּ���. ��<br/>�� ��Ϲ�ư Ŭ���� ���� Viewer�� �ݿ� �˴ϴ�.��</span></h3>
					<table class="tbType1 listTb tMar10" id="utypeI">
						<tbody>
						<tr>
							<td>
								<div id="spanslideimg"></div>
								<input class="button" type="button" value="�̹��� �ҷ�����" name="mslideimg" onClick="jsSetImg('<%=eFolder%>','','slideimg','spanslideimg');"/>
							</td>
						</tr>
						</tbody>
					</table>
				</div>
				<form name="frmList" method="POST" action="" style="margin:0;">
				<input type="hidden" name="mode" value="SU"/>
				<input type="hidden" name="device" value="M"/>
				<input type="hidden" name="eventid" value="<%=eCode%>"/>
				<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
				<div class="tMar20">
					<p>
						<input type="button" class="btn" value="���� ����" onclick="fnDeleteImg();">
					</p>
					<table class="tbType1 listTb tMar10">
						<colgroup>
							<col width="8%" /><col /><col width="42%" /><col width="10%" /><col width="40%" />
						</colgroup>
						<thead>
						<tr>
							<th><input type="checkbox" onclick="chkAllItem();"/></th>
							<th>�̹���</th>
							<th>����</th>
							<th>��뿩��</th>
						</tr>
						</thead>
						<tbody id="subList">
						<% 
							If eCode <> "" Then 
								sqlStr = "SELECT idx , imgurl, viewidx , videoLink, isusing" + vbcrlf
								sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_multi_contents] where menuidx = '"& menuidx &"' and device ='M' " 
								sqlStr = sqlStr & " order by viewidx asc, idx asc"
								rsget.Open sqlStr,dbget,1
								if Not(rsget.EOF or rsget.BOF) Then
									Do Until rsget.eof
						%>
						<tr class="<%=chkIIF(rsget("isusing")="N","bgGry1","")%>">
							<td><input type="checkbox" name="chkIdx" id="chkIdx" value="<%=rsget("idx")%>" /></td>
							<% if rsget("videoLink")<>"" then %>
							<td><%=rsget("videoLink")%></td>
							<% else %>
							<td><img src="<%=rsget("imgurl")%>" style="width:100px;" /></td>
							<% end if %>
							<td><input type="text" value="<%=rsget("viewidx")%>" name="viewidx<%=rsget("idx")%>"/></td>
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
				</div>
				</form>
			</dd>
		</dl>
		<div class="btnArea">
			<input type="image" src="http://scm.10x10.co.kr/images/icon_save.gif" alt="����" onclick="saveList();"/>
			<a href="" onclick="self.close();"><img src="http://scm.10x10.co.kr/images/icon_cancel.gif" alt="���" /></a>
		</div>
	</div>
</div>
<form name="frmdel" method="POST" action="pop_themeslide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SD"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="chkIdx" />
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
</form>
<form name="frmchkdel" method="POST" action="pop_themeslide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SAD"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="chkIdx" />
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
</form>
<form name="slideimgfrm" method="post" action="pop_themeslide_proc.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SI"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="slideimg" value=""/>
<input type="hidden" name="linkurl" id="linkurl" value=""/>
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->