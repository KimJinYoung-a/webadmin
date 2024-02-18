<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateMainCls.asp"-->
<%
	Dim cMain, i, vCateCode, vPage, vStartDate, vArr, vWWW
	Dim vWorkComment, vRegUserID
	vCateCode = Request("catecode")
	vPage = Request("page")
	'vPage = NullFillWith(Request("page"),"1")
	vStartDate = Request("startdate")
	
	Dim vMultiImg1, vMultiLink1, vMultiWorker, vMultiImg2, vMultiLink2, vMultiImg3, vMultiLink3, vItemID1, vItemImg1, vItem1Worker, vItemID2, vItemImg2, vItem2Worker, vItemID3, vItemImg3, vItem3Worker
	Dim vItemID4, vItemImg4, vItem4Worker, vItemID5, vItemImg5, vItem5Worker, vItemID6, vItemImg6, vItem6Worker, vItemID7, vItemImg7, vItem7Worker, vItemID8, vItemImg8, vItem8Worker
	Dim vItemID9, vItemImg9, vItem9Worker, vItemID10, vItemImg10, vItem10Worker, vItemID11, vItemImg11, vItem11Worker, vItemID12, vItemImg12, vItem12Worker
	Dim vEventID1, vEventImg1, vEvent1Worker, vEventID2, vEventImg2, vEvent2Worker, vEventID3, vEventImg3, vEvent3Worker, vEventID4, vEventImg4, vEvent4Worker, vBookImg, vBookLink, vBookWorker
	Dim vEventHtml1, vEventHtml2, vEventHtml3, vEventHtml4, vRecipeID, vRecipeImg, vRecipeWorker, vIdx

	If vStartDate <> "" Then
%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/mainpage/templete_include.asp"-->
<%
	End If
%>

<html>
<head>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link rel="stylesheet" href="/css/scm.css" type="text/css">

<script language="JavaScript" src="/js/calendar.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
document.domain = "10x10.co.kr";

function jsItemReg(t,i){
<% If vStartDate = "" Then %>
alert("새 페이지를 저장 또는 \n리스트에서 반영일을 선택하셔야 합니다.");
return;
<% End If %>
	var itempop1 = window.open("pop_item.asp?startdate=<%=vStartDate%>&page=<%=vPage%>&catecode=<%=vCateCode%>&type="+t+"&itemid="+i+"","itempop1","width=400,height=200, scrollbars=yes, resizable=yes");
	itempop1.focus();
}

function jsBannerReg(t){
<% If vStartDate = "" Then %>
alert("새 페이지를 저장 또는 \n리스트에서 반영일을 선택하셔야 합니다.");
return;
<% End If %>
	var bannerpop1 = window.open("pop_banner.asp?startdate=<%=vStartDate%>&page=<%=vPage%>&catecode=<%=vCateCode%>&type="+t+"","bannerpop1","width=600,height=350, scrollbars=yes, resizable=yes");
	bannerpop1.focus();
}

function jsEventReg(t,e){
<% If vStartDate = "" Then %>
alert("새 페이지를 저장 또는 \n리스트에서 반영일을 선택하셔야 합니다.");
return;
<% End If %>
	var eventpop1 = window.open("pop_event.asp?startdate=<%=vStartDate%>&page=<%=vPage%>&catecode=<%=vCateCode%>&type="+t+"&eventid="+e+"","eventpop1","width=400,height=200, scrollbars=yes, resizable=yes");
	eventpop1.focus();
}

function jsRecipeReg(t,r){
<% If vStartDate = "" Then %>
alert("새 페이지를 저장 또는 \n리스트에서 반영일을 선택하셔야 합니다.");
return;
<% End If %>
	var recipepop1 = window.open("pop_recipe.asp?startdate=<%=vStartDate%>&page=<%=vPage%>&catecode=<%=vCateCode%>&type="+t+"&recipeid="+r+"","recipepop1","width=400,height=200, scrollbars=yes, resizable=yes");
	recipepop1.focus();
}

function calendarOpenAA(objTarget){
<% If vStartDate <> "" Then %>
alert("반영일을 변경하면 해당 반영일의\n모든 페이지 반영일자가 변경됩니다.\n\n※ 반영일 변경은 신중히 해주세요.");
<% End If %>
    if (typeof calPopup == "function"){
        var compname = 'document.' + objTarget.form.name + '.' + objTarget.name;
        calPopup(objTarget,'calendarPopup',20+80,0, compname,'');
    }else{
        var fName = objTarget.form.name;
        var sName = objTarget.name;
    	var winCal = window.open('/lib/common_cal.asp?in_domain=o&FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
    	winCal.focus();
    }
}

function jsSaveCateMain(){
	if($("input[name=page]").val() == ""){
		alert("페이지를 선택해주세요.");
		return;
	}
	if($("input[name=startdate]").val() == ""){
		alert("반영일을 등록해주세요.");
		return;
	}
<% '<!-- #include virtual="/admin/CategoryMaster/displaycate/mainpage/javascript_form_check.asp"--> %>
	
	frmMain.action = "templete_proc.asp"
	frmMain.submit();
}

function jsUpdateCateMain(){
	frmMain.action = "templete_update.asp"
	frmMain.submit();
}

function jsDeleteCateMain(){
	if(confirm("삭제를 하시면 아래 모든 항목이 삭제 됩니다.\n정말 삭제하시겠습니까?") == true) {
		frmMain.mode.value = "delete";
		frmMain.action = "templete_update.asp"
		frmMain.submit();
	}
}

function jsRealServerReg(){
<%
	IF application("Svr_Info") = "Dev" THEN
		vWWW = "http://2013www.10x10.co.kr"
	Else
		vWWW = "http://www1.10x10.co.kr"
	End IF
%>
	var realpop1 = window.open("<%=vWWW%>/chtml/dispcate/category_main.asp?startdate=<%=vStartDate%>&page=<%=vPage%>&disp=<%=vCateCode%>","realpop1","width=1200,height=930, scrollbars=yes, resizable=yes");
	realpop1.focus();
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2; width:170px;"></div>
<form name="frmMain" method="post" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="page" value="<%=vPage%>">
<input type="hidden" name="startdate_before" value="<%=vStartDate%>">
<table bgcolor="#F4F4F4" width="100%" class="a">
<tr>
	<td style="padding:30px 0 10px 0;" width="25%">
		<% If vIdx <> "" Then Response.Write "<b>idx : " & vIdx & "&nbsp;</b>" End If %>
		반영일 : <input type="text" name="startdate" size="10" maxlength="10" style="border:1px solid black;" readonly value="<%=vStartDate%>">
		<a href="javascript:calendarOpenAA(frmMain.startdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>
	<td style="padding:30px 0 10px 0;" width="75%">
		<% If vStartDate <> "" Then %>
			<input type="button" value=" 반영일및작업코멘트수정 " style="border:1px solid black;" onClick="jsUpdateCateMain()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" value=" 미리보기 " style="border:1px solid black;" onClick="jsRealServerReg()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" value=" 새 페이지 생성 " style="border:1px solid black;" onClick="location.href='<%=CurrURL()%>?catecode=<%=vCateCode%>&page=<%=vPage%>';">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<% Else %>
		<input type="button" value=" 새 페이지 저장 " style="border:1px solid black;" onClick="jsSaveCateMain()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<% End If %>
		<% If vRegUserID = session("ssBctId") Then %>
			<input type="button" value=" 삭 제 " style="border:1px solid black;" onClick="jsDeleteCateMain()">
		<% End If %>
	</td>
</tr>
<tr>
	<td style="padding:0px 0 20px 0;" colspan="2">
		작업코멘트<br><textarea name="workcomment" cols="130" rows="10"><%=vWorkComment%></textarea>
	</td>
</tr>
</table>

<table bgcolor="#FFFFFF" cellpadding="7" cellspacing="7" border="0" class="a">
<tr>
	<td align="center" valign="middle">
	<table bgcolor="#FFFFFF" width="926px" style="padding:10px 0 0 20px;" cellpadding="3" cellspacing="0" border="0" class="a">
	<tr>
		<td><% Call printCategoryHistory(vCateCode) %></td>
		<td>반영일 : <b><%=vStartDate%></b></td>
		<td align="right">
			<%=CHKIIF(vPage="1","<b><font size=3 color=blue><u>01</u></font></b>","01")%>&nbsp;&nbsp;
			<%=CHKIIF(vPage="2","<b><font size=3 color=blue><u>02</u></font></b>","02")%>&nbsp;&nbsp;
			<%=CHKIIF(vPage="3","<b><font size=3 color=blue><u>03</u></font></b>","03")%>&nbsp;&nbsp;
			<%=CHKIIF(vPage="4","<b><font size=3 color=blue><u>04</u></font></b>","04")%>&nbsp;&nbsp;
			<%=CHKIIF(vPage="5","<b><font size=3 color=blue><u>05</u></font></b>","05")%>
		</td>
	</tr>
	</table>
	<table bgcolor="#FFFFFF" width="908px" cellpadding="3" cellspacing="0" class="a">
	<tr>
		<td>
			<table width="454px" border="1" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td colspan="2" align="center">
					<table id="multi" background="<%=vMultiImg1%>" cellpadding="0" cellspacing="0" width="444px" height="444px" height="100%" class="a">
					<tr>
						<td align="center">이미지1 만 보여집니다.<br>
						<% If vPage = "1" Then %><input type="button" value=" multi 등 록 " onClick="jsBannerReg('multi');"><br><br><% End If %>
						<span id="multiworker" style="background-color:#FFFFFF;"><%=vMultiWorker%></span></td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width="232px" align="center">
					<table id="item3" background="<%=vItemImg3%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
					<tr>
						<td align="center"><input type="button" value=" item3 등 록 " onClick="jsItemReg('item3','<%=vItemID3%>');"><br><br><span id="item3worker" style="background-color:#FFFFFF;"><%=vItem3Worker%></span></td>
					</tr>
					</table>
				</td>
				<td width="232px" align="center">
					<table id="item4" background="<%=vItemImg4%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
					<tr>
						<td align="center"><input type="button" value=" item4 등 록 " onClick="jsItemReg('item4','<%=vItemID4%>');"><br><br><span id="item4worker" style="background-color:#FFFFFF;"><%=vItem4Worker%></span></td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
		</td>
		<td valign="top" align="center">
			<table width="454px" border="1" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align="center" valign="top">
					<table width="227px" cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="item1" background="<%=vItemImg1%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
							<tr>
								<td align="center"><input type="button" value=" item1 등 록 " onClick="jsItemReg('item1','<%=vItemID1%>');"><br><br><span id="item1worker" style="background-color:#FFFFFF;"><%=vItem1Worker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
				<td align="center" valign="top">
					<table width="227px" cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="item2" background="<%=vItemImg2%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
							<tr>
								<td align="center"><input type="button" value=" item2 등 록 " onClick="jsItemReg('item2','<%=vItemID2%>');"><br><br><span id="item2worker" style="background-color:#FFFFFF;"><%=vItem2Worker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width="227px" rowspan="2" align="center" valign="top">
					<table cellpadding="0" cellspacing="0" class="a" align="center" height="444px">
					<tr>
						<td align="center" valign="top">
							<table id="event1" background="<%=vEventImg1%>" cellpadding="0" cellspacing="0" class="a">
							<tr>
								<td align="center" width="200px" height="200px">
									<input type="button" value=" event1 등 록 " onClick="jsEventReg('event1','<%=vEventID1%>');"><br><br><span id="event1worker" style="background-color:#FFFFFF;"><%=vEvent1Worker%></span></td>
							</tr>
							</table>
							<table cellpadding="0" cellspacing="0" class="a">
							<tr>
								<td width="160px"><span id="event1description"><%=vEventHtml1%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
				<td width="227px" rowspan="2" align="center" valign="top">
					<table cellpadding="0" cellspacing="0" class="a" align="center" height="444px">
					<tr>
						<td align="center" valign="top">
							<table id="event2" background="<%=vEventImg2%>" cellpadding="0" cellspacing="0" class="a">
							<tr>
								<td align="center" width="200px" height="200px">
									<input type="button" value=" event2 등 록 " onClick="jsEventReg('event2','<%=vEventID2%>');"><br><br><span id="event2worker" style="background-color:#FFFFFF;"><%=vEvent2Worker%></span></td>
							</tr>
							</table>
							<table cellpadding="0" cellspacing="0" class="a">
							<tr>
								<td width="160px"><span id="event2description"><%=vEventHtml2%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<table border="1" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align="center" valign="top">
					<table width="224px" cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="item5" background="<%=vItemImg5%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
							<tr>
								<td align="center"><input type="button" value=" item5 등 록 " onClick="jsItemReg('item5','<%=vItemID5%>');"><br><br><span id="item5worker" style="background-color:#FFFFFF;"><%=vItem5Worker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
				<td align="center" valign="top">
					<table width="224px" cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="item6" background="<%=vItemImg6%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
							<tr>
								<td align="center"><input type="button" value=" item6 등 록 " onClick="jsItemReg('item6','<%=vItemID6%>');"><br><br><span id="item6worker" style="background-color:#FFFFFF;"><%=vItem6Worker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
				<td width="465px">
					<table id="book" background="<%=vBookImg%>" height="212px" width="444px" cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
						<% If vPage = "1" Then %><input type="button" value=" book 등 록 " onClick="jsBannerReg('book');"><br><br><% End If %>
						<span id="bookworker" style="background-color:#FFFFFF;"><%=vBookWorker%></span></td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<table border="1" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td width="224px" rowspan="2" align="center" valign="top">
					<table cellpadding="0" cellspacing="0" class="a" align="center" height="444px">
					<tr>
						<td align="center" valign="top">
							<table id="event3" background="<%=vEventImg3%>" cellpadding="0" cellspacing="0" class="a">
							<tr>
								<td align="center" width="200px" height="200px">
									<input type="button" value=" event3 등 록 " onClick="jsEventReg('event3','<%=vEventID3%>');"><br><br><span id="event3worker" style="background-color:#FFFFFF;"><%=vEvent3Worker%></span></td>
							</tr>
							</table>
							<table cellpadding="0" cellspacing="0" class="a">
							<tr>
								<td width="160px"><span id="event3description"><%=vEventHtml3%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
				<td width="224px" rowspan="2" align="center" valign="top">
					<table cellpadding="0" cellspacing="0" class="a" align="center" height="444px">
					<tr>
						<td align="center" valign="top">
							<table id="event4" background="<%=vEventImg4%>" cellpadding="0" cellspacing="0" class="a">
							<tr>
								<td align="center" width="200px" height="200px">
									<input type="button" value=" event4 등 록 " onClick="jsEventReg('event4','<%=vEventID4%>');"><br><br><span id="event4worker" style="background-color:#FFFFFF;"><%=vEvent4Worker%></span></td>
							</tr>
							</table>
							<table cellpadding="0" cellspacing="0" class="a">
							<tr>
								<td width="160px"><span id="event4description"><%=vEventHtml4%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
				<td align="center" valign="top" align="absmiddle">
					<table width="231px" cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="item7" background="<%=vItemImg7%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
							<tr>
								<td align="center"><input type="button" value=" item7 등 록 " onClick="jsItemReg('item7','<%=vItemID7%>');"><br><br><span id="item7worker" style="background-color:#FFFFFF;"><%=vItem7Worker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
				<td align="center" valign="top" align="absmiddle">
					<table width="231px" cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="item8" background="<%=vItemImg8%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
							<tr>
								<td align="center"><input type="button" value=" item8 등 록 " onClick="jsItemReg('item8','<%=vItemID8%>');"><br><br><span id="item8worker" style="background-color:#FFFFFF;"><%=vItem8Worker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td align="center" align="absmiddle">
					<table cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="item9" background="<%=vItemImg9%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
							<tr>
								<td align="center"><input type="button" value=" item9 등 록 " onClick="jsItemReg('item9','<%=vItemID9%>');"><br><br><span id="item9worker" style="background-color:#FFFFFF;"><%=vItem9Worker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
				<td align="center" align="absmiddle">
					<table cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="item10" background="<%=vItemImg10%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
							<tr>
								<td align="center"><input type="button" value=" item10 등 록 " onClick="jsItemReg('item10','<%=vItemID10%>');"><br><br><span id="item10worker" style="background-color:#FFFFFF;"><%=vItem10Worker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<table border="1" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align="center" align="absmiddle">
					<table width="224px" cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="item11" background="<%=vItemImg11%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
							<tr>
								<td align="center"><input type="button" value=" item11 등 록 " onClick="jsItemReg('item11','<%=vItemID11%>');"><br><br><span id="item11worker" style="background-color:#FFFFFF;"><%=vItem11Worker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
				<td align="center" align="absmiddle" rowspan="2">
					<table width="690px" cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="recipe" background="<%=vRecipeImg%>" height="444px" width="676px" align="right" class="a">
							<tr>
								<td align="center">
								<% If vPage = "1" Then %><input type="button" value=" recipe 등 록 " onClick="jsBannerReg('recipe');"><br><br><% End If %>
								<span id="recipeworker" style="background-color:#FFFFFF;"><%=vRecipeWorker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td align="center" align="absmiddle">
					<table cellpadding="0" cellspacing="0" class="a" align="center">
					<tr>
						<td align="center">
							<table id="item12" background="<%=vItemImg12%>" cellpadding="0" cellspacing="0" width="200px" height="200px" class="a">
							<tr>
								<td align="center"><input type="button" value=" item12 등 록 " onClick="jsItemReg('item12','<%=vItemID12%>');"><br><br><span id="item12worker" style="background-color:#FFFFFF;"><%=vItem12Worker%></span></td>
							</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	</table>
	</td>
</tr>
</table>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->