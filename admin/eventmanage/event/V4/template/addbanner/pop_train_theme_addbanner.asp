<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_train_theme_addbanner.asp
' Discription : H�� �߰� �׸� ������ ��� ���
' History : 2018.08.14 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim eCode , title , eFolder , topimg , btmimg , topaddimg 'floating img
Dim slideimg, GroupItemCheck
Dim mode , idx , strSql , sqlStr, smode, saveafter
smode = requestCheckvar(request("smode"),16)
saveafter = requestCheckvar(request("saveafter"),2)
GroupItemCheck = requestCheckvar(request("GroupItemCheck"),1)
If saveafter <>"" Then
Response.write "<script>opener.location.reload();</script>"
End If

If smode="SU" Then
Response.write "<script>self.close();</script>"
Response.end
End If

eCode = requestCheckvar(request("eC"),16)

title = "�����̵� ��� �˾�(M)"

eFolder = eCode

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
		document.frmList.action="dotrainthemebanner.asp";
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
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV3.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan+'&wid=1000&hei=1000','popImg','width=370,height=150');
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
		url: "pop_mobile_themeslide_ajax.asp",
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
	<h1>���ø� �̹��� ���</h1>
	<div class="register">
		<h2>������ ���</h2>
		<dl>
			<dt>- �����̵� <span>(�ʼ�/3~10������ ���)</span></dt>
			<dd>
				<div class="insertImg">
					<h3>�����̵� �̹��� ���<br/><span style="color:#c80a0a;line-height:2;">�� ���ϴ� �̹��� �켱 ����� �����̵� �̹����� ������ּ���. �Ǵ� �����̵� �켱 ����� ���ϴ� �̹����� ��� ���ּ���. ��<br/>�� ��Ϲ�ư Ŭ���� ���� Viewer�� �ݿ� �˴ϴ�.��<br/>�� �̹��� ������ : 600 * 600��</span></h3>
					<table class="tbType1 listTb tMar10">
						<colgroup>
							<col width="50%" /><col /><col width="50%" /><col />
						</colgroup>
						<tbody>
						<tr>
							<td>
								<div id="spanslideimg"></div>
								<input class="button" type="button" value="�̹��� �ҷ�����" name="mslideimg" onClick="jsSetImg('<%=eFolder%>','','slideimg','spanslideimg');"/>
							</td>
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

							sqlStr = "SELECT g.idx , g.title, g.imgurl , g.groupcode, g.viewidx, i.basicimage" + vbcrlf
							sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_manual_group] g "
							sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item] i on i.itemid=g.itemid"
							sqlStr = sqlStr & " where g.evt_code = '"& eCode &"'"
							sqlStr = sqlStr & " and g.grouptype='B'"
							sqlStr = sqlStr & " order by g.viewidx asc , g.idx asc "
							rsget.Open sqlStr,dbget,1
							if Not(rsget.EOF or rsget.BOF) Then
								Do Until rsget.eof
						%>
						<tr>
							<td><input type="checkbox" name="chkIdx" id="chkIdx" value="<%=rsget("idx")%>" /></td>
							<td><img src="<% if rsget("imgurl")<>"" then %><%=rsget("imgurl")%><% else %><%=rsget("basicimage")%><% end if %>" style="width:100px;" /></td>
							<td><input type="text" value="<%=rsget("viewidx")%>" name="sort<%=rsget("idx")%>"/></td>
							<td><input type="button" class="btn" value="����" onclick="slideimgDel(<%=rsget("idx")%>);"></td>
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
			<input type="image" src="http://webadmin.10x10.co.kr/images/icon_save.gif" alt="����" onclick="saveList();"/>
			<a href="" onclick="self.close();"><img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" alt="���" /></a>
		</div>
	</div>
</div>
<form name="frmdel" method="POST" action="dotrainthemebanner.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SD"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="chkIdx" />
</form>
<form name="slideimgfrm" method="post" action="dotrainthemebanner.asp" style="margin:0px;">
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SI"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="slideimg" value=""/>
<input type="hidden" name="linkurl" id="linkurl" value=""/>
<input type="hidden" name="GroupItemCheck" value="<%=GroupItemCheck%>"/>
</form>
<form name="slidefrm" method="post" action="dotrainthemebanner.asp" style="margin:0px;">
<input type="hidden" name="idx" value="<%=idx%>"/>
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="<%=mode%>"/>
<input type="hidden" name="device" value="M"/>
<input type="hidden" name="topimg" value="<%=topimg%>"/>
<input type="hidden" name="btmimg" value="<%=btmimg%>"/>
<input type="hidden" name="topaddimg" value="<%=topaddimg%>"/>
<input type="hidden" name="GroupItemCheck" value="<%=GroupItemCheck%>"/>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->