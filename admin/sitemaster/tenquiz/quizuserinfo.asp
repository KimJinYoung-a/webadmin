<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/TenQuizCls.asp" -->
<%
	dim page
	dim i
	dim tenQuizList
	dim optionMonthGroup
	dim optionUserType
	dim optionKeyword	 
	dim KeyWord
	dim currentPath
    dim userid

	currentPath = request.ServerVariables("PATH_INFO")	 

	page 				= request("page")
	optionMonthGroup	= request("optionMonthGroup")
	optionKeyword	 	= request("optionKeyword")
	KeyWord				= request("KeyWord")
    userid              = request("userid")

	if page="" then page=1	

	set tenQuizList = new TenQuiz
	tenQuizList.FPageSize			= 20
	tenQuizList.FCurrPage			= page
	
	if optionMonthGroup <> "" then
		tenQuizList.FmonthGroupOption	= optionMonthGroup
	end if	    	
	if optionUserType <> "0" then
		tenQuizList.FQuizUserOption	= optionUserType
	end if		
	
	if KeyWord <> "" then
        if optionKeyword = "chasu" then
            tenQuizList.FChasuOption		= KeyWord
        else
            tenQuizList.FUserIdOption		= KeyWord
        end if
	end if		    

	tenQuizList.GetUserInfoList()	    
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
	location.href = "addtenquizcontent.asp?menupos=<%=menupos%>&idx="+v;
}
function quizTest(chasu){
	alert("�������� ����Դϴ�.");	
	// location.href = "addtenquizcontent.asp?menupos=<%=menupos%>&idx="+chasu;
}
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function userQuizInfo(userid, chasu){
	var popwin = window.open("/admin/sitemaster/tenquiz/userquizinfo.asp?userid="+userid+"&chasu="+chasu, "mileagediv", "width=800,height=850,scrollbars=yes,resizable=yes");
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
</script>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>

<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2>[ON]����Ʈ���� &gt; <strong><a href="">������</a></strong></h2></div>
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
			<li class="col11 <%=chkIIF(currentPath = "/admin/sitemaster/tenquiz/index.asp","selected","")%> "><a href="index.asp">�����Ʈ</a></li>
			<li class="col11 <%=chkIIF(currentPath = "/admin/sitemaster/tenquiz/quizuserinfo.asp","selected","")%>"><a href="quizuserinfo.asp">���������</a></li>
		</ul>
	</div>

	<!-- ��� �˻��� ���� -->
	<form name="frm" method="post" style="margin:0px;" action="/admin/sitemaster/tenquiz/quizuserinfo.asp">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<!-- search -->
	<div class="searchWrap" style="border-top:none;">
		<div class="search">
			<ul>
				<li>
					<label class="formTit">� �� :</label>
					<input type="text" name="optionMonthGroup" class="formTxt" id="datepicker" style="width:55px" maxlength=6 value="<%=optionMonthGroup%>"/>
				</li>
				<li>
					<p class="formTit">��� :</p>
					<select class="formSlt" id="open" title="�ɼ� ����" name="optionUserType">
						<option value=0 <%=chkIIF(optionUserType=0,"selected","")%>>==����==</option>
						<option value=1 <%=chkIIF(optionUserType=1,"selected","")%>>������</option>
						<option value=2 <%=chkIIF(optionUserType=2,"selected","")%>>������</option>
					</select>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="schWord">Ű���� �˻� :</label>
					<select class="formSlt" name="OptionKeyword" id="keyword" title="Ű���� �˻�">
						<option value="chasu" <%=chkIIF(optionKeyword="chasu","selected","")%>>����</option>
						<option value="userid" <%=chkIIF(optionKeyword="userid","selected","")%>>������id</option>
					</select>
					<input type="text" name="Keyword" class="formTxt" id="schWord" style="width:400px" placeholder="Ű���带 �Է��Ͽ� �˻��ϼ���." value="<%=Keyword%>"/>
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
					
				</div>
			</div>
			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">�� ������ : <strong><%=tenQuizList.FtotalCount%></strong></p>
				</div>
				<div class="tbListWrap">
					<ul class="thDataList">
						<li>
                            <p style="width:110px">����</p>
							<p style="width:110px">���</p>
							<p style="width:110px">���̵�</p>
							<p style="width:110px">����</p>							
							<p style="width:110px">���䰳��</p>
							<p style="width:110px">�ֱٱ�������</p>
							<p style="width:110px">��������Ƚ��</p>
							<p style="width:110px">���ϸ����ֱٻ��</p>
						</li>
					</ul>
					<!-- ����Ʈ -->
					<ul class="tbDataList">       
<% for i=0 to tenQuizList.FResultCount-1 %>					
						<li style="cursor:pointer;" onclick="userQuizInfo('<%=tenQuizList.FItemList(i).FUuserId%>','<%=tenQuizList.FItemList(i).FUchasu%>')" onmouseover="this.style.backgroundColor='#D8D8D8'" onmouseout="this.style.backgroundColor=''">
							<p style="width:110px"><%=tenQuizList.FItemList(i).FUchasu%></p>
							<p style="width:110px"><%=tenQuizList.FItemList(i).FUuserLevel%></p>
							<p style="width:110px"><%=tenQuizList.FItemList(i).FUuserId%></p>
                            <p style="width:110px"><%=tenQuizList.FItemList(i).FUage%></p>                            				
                            <p style="width:110px"><%=tenQuizList.FItemList(i).FUuserScore%></p>                            				
                            <p style="width:110px"><%=chkIIF(tenQuizList.FItemList(i).FUbuyDate<>"", tenQuizList.FItemList(i).FUbuyDate, "����")%></p>                            				
                            <p style="width:110px"><%=tenQuizList.FItemList(i).FUquizCnt%></p>                            				
                            <p style="width:110px"><%=chkIIF(tenQuizList.FItemList(i).FUrecentMileageLog<>"", tenQuizList.FItemList(i).FUrecentMileageLog, "����")%></p>    
						</li>						
<% Next %>						
					</ul>
					<div class="ct tPad20 cBk1">
						<% if tenQuizList.HasPreScroll then %>
							<span class="list_link"><a href="?page=<%= tenQuizList.StartScrollPage-1 %>">[pre]</a></span>
						<% else %>
						[pre]
						<% end if %>
						<% for i = 0 + tenQuizList.StartScrollPage to tenQuizList.StartScrollPage + tenQuizList.FScrollCount - 1 %>
							<% if (i > tenQuizList.FTotalpage) then Exit for %>
							<% if CStr(i) = CStr(tenQuizList.FCurrPage) then %>
							<span class="page_link"><font color="red"><b><%= i %></b></font></span>
							<% else %>
							<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
							<% end if %>
						<% next %>
						<% if tenQuizList.HasNextScroll then %>
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

