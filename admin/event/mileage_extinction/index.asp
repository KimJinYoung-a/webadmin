<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<!-- #include virtual="/lib/classes/event/EventMileageCls.asp" -->
<%
	dim i, currentPath, taskList, page
	currentPath = request.ServerVariables("PATH_INFO")

	page 				= request("page")
	if page="" then page=1

	'�˻� �Ķ����
	dim optionsdt, optionedt, optionkeyword, keyword

	optionsdt = request("optionsdt")
	optionedt = request("optionedt")
	optionkeyword = request("optionkeyword")
	keyword = request("keyword")

    set taskList = new MileageExtinctionCls
	taskList.FPageSize			= 20
	taskList.FCurrPage			= page
	
	taskList.FUsdt = optionsdt
	taskList.FUedt = optionedt
	taskList.FUkeyword =keyword
	taskList.FUoption = optionkeyword	

    taskList.getTaskList()
%>
<style type="text/css">

</style>
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<script language='javascript'>
$(function(){
	$("li a").click(function(e){
		e.stopPropagation();
	});
	$("li span").click(function(e){
		e.stopPropagation();
	});
    $('#datepicker').datepicker( {
        changeMonth: true,
        changeYear: true,
        showButtonPanel: true,
        dateFormat: 'yymm'
    });
    $('#startDate').datepicker( {
        changeMonth: true,
        changeYear: true,
        showButtonPanel: true,
        dateFormat: 'yy-mm-dd',
    });
    $('#endDate').datepicker( {
        changeMonth: true,
        changeYear: true,
        showButtonPanel: true,
        dateFormat: 'yy-mm-dd',
    });
})
function jsmodify(v){
	location.href = "addtenquizcontent.asp?idx="+v;
}
function quizTest(chasu){
	alert("�������� ����Դϴ�.");

}
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;
	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "���ã�⿡�� �����Ͻðڽ��ϱ�?";
	} else {
		msg = "���ã�⿡ �߰��Ͻðڽ��ϱ�?";
	}
	ret = confirm(msg);
	if (ret) {
		frm.submit();
	}
}
function jsOpen(sPURL,sTG){
	if (sTG =="M" ){
		var winView = window.open(sPURL,"popView","width=400, height=600,scrollbars=yes,resizable=yes");
	}
}
function popQuestionEdit(id = ""){
	var popwin = window.open("popedit.asp?id=" + id, "popup", "width=800,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>

<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2>[ON]�̺�Ʈ���� &gt; <strong><a href="">�̺�Ʈ���ϸ��� �Ҹ� ����</a></strong></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="1836">
			</form>
			<a href="javascript:fnMenuFavoriteAct('addonefavorite')">���ã��</a> l
		</div>
	</div>
	<div class="tab" style="margin:0 0 0 -1px;">
		<ul>
			<li class="col11 <%=chkIIF(currentPath = "/admin/event/mileage_extinction/index.asp","selected","")%> "><a href="index.asp">�Ҹ��۾�����Ʈ</a></li>
			<li class="col11 <%=chkIIF(currentPath = "/admin/event/mileage_extinction/tasklog","selected","")%>"><a href="tasklog.asp">�۾�����</a></li>
		</ul>
	</div>

	<!-- ��� �˻��� ���� -->
	<form name="frm" method="post" style="margin:0px;" action="">
	<input type="hidden" name="page" value="<%=page%>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<!-- search -->
	<div class="searchWrap" style="border-top:none;">
		<div class="search">
			<ul>				
				<li>
					<p class="formTit">������ :</p>
					<input type="text" name="optionSdt" class="formTxt" id="startDate" style="width:85px" value="<%=optionsdt%>"/>
				</li>
				<li>
					<p class="formTit">������ :</p>
					<input type="text" name="optionEdt" class="formTxt" id="endDate" style="width:85px" value="<%=optionedt%>"/>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="schWord">Ű���� �˻� :</label>
					<select class="formSlt" name="OptionKeyword" id="keyword" title="Ű���� �˻�">
						<option value="" <%=chkIIF(optionkeyword = "" , "selected", "")%>>=����=</option>
						<option value="jukyo" <%=chkIIF(optionkeyword = "jukyo" , "selected", "")%>>����</option>
						<option value="jukyocd" <%=chkIIF(optionkeyword = "jukyocd", "selected", "")%>>�����ڵ�</option>
					</select>
					<input type="text" name="Keyword" class="formTxt" id="schWord" style="width:400px" placeholder="Ű���带 �Է��Ͽ� �˻��ϼ���./" value="<%=keyword%>"/>					
					<button type="button" onclick="window.location.href=''">�ʱ�ȭ</button>
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="�˻�" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt">
					<input type="button" class="btnRegist btn bold fs12" value="�Ҹ� �۾� ���" onclick="popQuestionEdit()"/>
				</div>
			</div>
			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">�� ��ϼ� : <strong><%=taskList.FtotalCount%></strong></p>
				</div>
				<div class="tbListWrap">
					<ul class="thDataList">
						<li>
							<p style="width:140px">����(����Ʈ�� ����)</p>
							<p style="width:90px">�̺�Ʈ�ڵ�(�����ڵ�)</p>
							<p style="width:55px">üũ �� ��</p>
							<p style="width:65px">�̺�Ʈ������</p>
							<p style="width:65px">�̺�Ʈ������</p>
							<p style="width:65px">�Ҹ��۾�������</p>
							<p style="width:70px">���� �������</p>
							<p style="width:70px">���� ��������</p>
							<p style="width:65px">�۾��ϷῩ��</p>
						</li>
					</ul>
					<!-- ����Ʈ -->
					<ul class="tbDataList">
<%
    dim lastDays, statusTxt, daysInfo, txtColor
	for i=0 to taskList.FResultCount-1
		lastDays = datediff("d", date(),dateadd("d", taskList.FItemList(i).task_chkDays + 1, taskList.FItemList(i).task_enddate))
'����, ������, ���࿹��
'�� ����, ���� ����, ���� ����, ""
		if lastDays = 0 then
			statusTxt = "������"
			daysInfo = "���� ����"
			txtColor = "blue"
		elseif lastDays < 0 then
			statusTxt = "����"
			daysInfo = 	""
			txtColor = "red"
		elseif lastDays <= taskList.FItemList(i).task_chkDays then
			statusTxt = "������"
			daysInfo = 	lastDays & "�� ����"
			txtColor = "blue"
		elseif lastDays > taskList.FItemList(i).task_chkDays then
			statusTxt = "���� ����"
			daysInfo = 	lastDays - taskList.FItemList(i).task_chkDays & "�� �� ����"
			txtColor = "mediumspringgreen"
		end if

		IF taskList.FItemList(i).task_taskStatus = "1" THEN
			statusTxt = "����"
			daysInfo = 	""
			txtColor = "red"
		END IF
%>
						<li style="cursor:pointer;" onclick="popQuestionEdit(<%=taskList.FItemList(i).task_id%>)" onmouseover="this.style.backgroundColor='#a3d0f5'" onmouseout="this.style.backgroundColor=''">
							<p style="width:110px"><%=taskList.FItemList(i).task_jukyo%></p>
							<p style="width:90px"><%=taskList.FItemList(i).task_jukyocd%></p>
							<p style="width:55px"><%=taskList.FItemList(i).task_chkDays%></p>
							<p style="width:60px"><%=LEFT(taskList.FItemList(i).task_startdate, 10)%></p>
							<p style="width:65px"><%=LEFT(taskList.FItemList(i).task_enddate, 10)%></p>
							<p style="width:65px"><%=LEFT(dateadd("d", 1, taskList.FItemList(i).task_enddate), 10)%></p>
							<p style="width:70px"><%=taskList.FItemList(i).task_regUser%><br /><%=taskList.FItemList(i).task_regdate%></p>
							<p style="width:70px"><%=taskList.FItemList(i).task_updateUser%><br /><%=taskList.FItemList(i).task_lastupdate%></p>
							<p style="width:65px">
                                <strong class="fs14" style="color:<%=txtColor%>"><%=statusTxt%></strong>
                                <br/>
                                <%=daysInfo%>
                            </p>
						</li>
<% Next %>
					</ul>
					<div class="ct tPad20 cBk1">
						<% if taskList.HasPreScroll then %>
							<span class="list_link"><a href="?page=<%= taskList.StartScrollPage-1 %>">[pre]</a></span>
						<% else %>
						[pre]
						<% end if %>
						<% for i = 0 + taskList.StartScrollPage to taskList.StartScrollPage + taskList.FScrollCount - 1 %>
							<% if (i > taskList.FTotalpage) then Exit for %>
							<% if CStr(i) = CStr(taskList.FCurrPage) then %>
							<span class="page_link"><font color="red"><b><%= i %></b></font></span>
							<% else %>
							<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
							<% end if %>
						<% next %>
						<% if taskList.HasNextScroll then %>
							<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
						<% else %>
						[next]
						<% end if %>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
