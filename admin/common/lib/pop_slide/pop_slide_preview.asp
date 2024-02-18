<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : �����̵� ���� - preview �̸�����
' History : 2019-02-20 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/common/lib/pop_slide/classes/slidemanageCls.asp"-->
<%
Dim title
Dim menu , mastercode , detailcode , prevDate , device
Dim oSlideManage
dim i 

menu = request("menu")
mastercode = request("mastercode")
detailcode = request("detailcode")
prevDate = request("prevDate")
device = request("device")

if prevDate = "" then prevDate = date()
if device = "" then device = "P"

set oSlideManage = new SlideListCls
    oSlideManage.FPageSize = 10
	oSlideManage.FCurrPage = 1
	oSlideManage.FrectMasterCode = mastercode
	oSlideManage.FrectDetailCode = detailcode
    oSlideManage.FRectSelDate    = prevDate
    oSlideManage.FRectMenu       = menu
    oSlideManage.FRectDevice     = device
    oSlideManage.FRectOrderby    = "sort"
	oSlideManage.getSlideList()

	title = "�����̵� �̸����� �˾�"& chkiif(device="P","(PC)","(M/A)")

%>
<!-- #include virtual="/admin/lib/popheaderslide.asp"-->
<link rel="stylesheet" type="text/css" href="http://webadmin.10x10.co.kr/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="http://webadmin.10x10.co.kr/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
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
		document.frmList.action="pop_slide_manage_proc.asp";
		document.frmList.submit();
	}
}

//'������ ����
function slideimgDel(v){
	if (confirm("�ش� �������� ���� �Ͻðڽ��ϱ�?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.submit();
	}
}

// ���� �˾�
function fnAddPopSlideManage(idx,m,d,device){
    var popwin = window.open('/admin/common/lib/pop_slide/pop_slide_manage_insert.asp?idx='+idx+'&menu=<%=menu%>&mastercode='+m+'&detailcode='+d+'&device='+device,'mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// ajaxlist
function dfslide(){
	var str = $.ajax({
		type: "GET",
		url: "ajax_pop_slide_preview.asp",
		data: "menu=<%=menu%>&mastercode=<%=mastercode%>&detailcode=<%=detailcode%>&prevDate=<%=prevDate%>&device=<%=device%>",
		dataType: "text",
		async: false
		}).responseText;
	if (str != ""){
		$('#preview_ajax').append(str);
	}
}

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
				if(i>$(this).val()) i = 0 //i = $(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});

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
	<h1><%=title%></h1>
	<%'�����̵� ȭ�� �ҷ�����%>
	<div class="preview" id="preview_ajax" style="padding-top:5%;"></div>
	<%'�����̵� ȭ�� �ҷ�����%>
	<div class="register">
		<h2 style="text-align:left;font-size:12px;">
            <form name="frmsearch" method="get" action="" style="margin:0px;">
            <input type="hidden" name="menu" value="<%=menu%>"/>
            <input type="hidden" name="mastercode" value="<%=mastercode%>" />
            <input type="hidden" name="detailcode" value="<%=detailcode%>"/>
            �������� :  <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" style="vertical-align:middle;"/>
            <script language="javascript">
                var CAL_Start = new Calendar({
                    inputField : "prevDate", trigger    : "prevDate_trigger",
                    onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
                });
            </script>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            ä�� : 
            <input type="radio" name="device" value="P" <% if device = "P" then response.write "checked"%>/> PC
            <input type="radio" name="device" value="M" <% if device = "M" then response.write "checked"%>/> M/A
            <span style="float:right;"><input type="submit" class="button_s" value="�˻�"></span>
            </form>
        </h2>
		<dl>
			<dd>
				<form name="frmList" method="POST" action="" style="margin:0;">
				<input type="hidden" name="mode" value="sort"/>
				<input type="hidden" name="device" value="M"/>
                <input type="hidden" name="backurl" value="<%=Request.ServerVariables("HTTP_URL")%>"/>
				<div class="tMar10">
                    <span style="color:#ff0000">
                        �� �⺻ �������ڴ� <%=date()%> �Դϴ�. ��<br/><br/>
                        �� �⺻ ä���� PC �Դϴ�. ��<br/><br/>
                        �� ���콺 �巡�׷� ������ �����ϰ� �������� ���� �Ͽ� ���������� �����ּ���. ��
                    </span>
					<p style="text-align:right;">
						<input type="button" class="btn" value="���� ����" onClick="saveList();" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
					</p>
					<table class="tbType1 listTb tMar10">
						<colgroup>
							<col width="7%" /><col /><col width="10%" /><col width="25%" /><col width="12%" />
						</colgroup>
						<thead>
						<tr>
							<th><input type="checkbox" onclick="chkAllItem();"/></th>
                            <th>idx</th>
							<th>�̹���</th>
							<th>Ÿ��Ʋ(�������� �ؽ�Ʈ)</th>
							<th>����</th>
							<th>��뿩��</th>
						</tr>
						</thead>
						<tbody id="subList">
						<%
                            for i=0 to oSlideManage.FResultCount - 1
                        %>
						<tr class="<%=chkiif((oSlideManage.FItemList(i).IsEndDateExpired) or (oSlideManage.FItemList(i).FIsusing="0"),"bgGry1","")%>">
							<td><input type="checkbox" name="chkIdx" value="<%=oSlideManage.FItemList(i).Fidx%>" /></td>
                            <td><a href="javascript:fnAddPopSlideManage('<%=oSlideManage.FItemList(i).Fidx%>','','','');"><%= oSlideManage.FItemList(i).Fidx%></a></td>
							<td>
                                <% if oSlideManage.FItemList(i).Fisvideo = 1 then %>
                                ������
                                <% else %>
                                    <% if oSlideManage.FItemList(i).Fimageurl = "" then %>
                                    �̹��� �̵��
                                    <% else %>
                                    <img src="<%= oSlideManage.FItemList(i).Fimageurl %>" width="75"/>
                                    <% end if %>
                                <% end if %>
                            </td>
							<td><%= oSlideManage.FItemList(i).Ftitlename%></td>
							<td><input type="text" value="<%= oSlideManage.FItemList(i).Fsorting%>" name="sort<%=oSlideManage.FItemList(i).Fidx%>"/></td>
							<td>
								<span><input type="radio" <%= chkiif(oSlideManage.FItemList(i).Fisusing="1","checked","") %> name="use<%=oSlideManage.FItemList(i).Fidx%>" value="1"/> Y</span>
								<span class="lMar10"><input type="radio" <%= chkiif(oSlideManage.FItemList(i).Fisusing="0","checked","") %> name="use<%=oSlideManage.FItemList(i).Fidx%>" value="0"/> N</span>
								<br/><input type="button" class="btn" value="����" onclick="slideimgDel('<%=oSlideManage.FItemList(i).Fidx%>');">
							</td>
						</tr>
						<% 
                            next 
                        %>
						</tbody>
					</table>
				</div>
				</form>
			</dd>
		</dl>
		<div class="btnArea">
			<input type="image" src="http://webadmin.10x10.co.kr/images/icon_save.gif" alt="����" onclick="mimgsubmit();"/>
			<a href=""><img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" alt="���" /></a>
		</div>
	</div>
</div>
<form name="frmdel" method="POST" action="pop_slide_manage_proc.asp" style="margin:0px;">
<input type="hidden" name="mode" value="idel"/>
<input type="hidden" name="chkIdx" />
<input type="hidden" name="backurl" value="<%=Request.ServerVariables("HTTP_URL")%>"/>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->