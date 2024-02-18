<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 1:1 ���
' History : 2009.04.17 �̻� ����
'			2016.03.25 �ѿ�� ����(���Ǻо� ��� DBȭ ��Ŵ)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%
dim i, j, sDate ,eDate ,blnDate, page, ckReplyDate, replyDate1, replyDate2, searchDiv, searchText, tmpqadivname
dim itemqanotinclude, research, finishyn, userid, orderserial, qadiv , writeid, chargeid, replyqadiv, userlevel, evalPoint
dim isusing, sitename
dim currstate, userGubun
	qadiv               = request("qadiv")
	itemqanotinclude    = request("itemqanotinclude")
	research            = request("research")
	userid              = request("userid")
	orderserial         = request("orderserial")
	writeid             = request("writeid")
	chargeid            = request("chargeid")
	replyqadiv          = request("replyqadiv")
	userlevel			= request("userlevel")
	evalPoint			= request("evalPoint")
	searchDiv			= request("searchDiv")
	searchText			= Trim(request("searchText"))
	sDate = request("sdt")
	eDate = request("edt")
	blnDate = request("edc")
	'if (itemqanotinclude="") and (research="") then itemqanotinclude="on"
	isusing				= request("isusing")
	sitename			= request("sitename")

	if (sitename = "") and not(C_ADMIN_AUTH) then
		sitename = "10x10"
	end if

	currstate			= request("currstate")
	userGubun			= request("userGubun")

if (research = "") then
	isusing = "Y"
	sitename = "10x10"
end if

if ((userid <> "") or (orderserial <> "")) then
    qadiv = ""
    itemqanotinclude = ""
end if

page	= req("page",1)

ckReplyDate	= req("ckReplyDate",req("ckReplyDateDefault",""))
replyDate1	= req("replyDate1",LEFT(CStr(dateAdd("d",-7,now())),10))
replyDate2	= req("replyDate2",LEFT(CStr(now()),10))

if (blnDate="") and (research="") then
    blnDate = "on"
    sDate   = LEFT(CStr(dateAdd("m",-3,now())),10)
    eDate   = LEFT(CStr(now()),10)
end if

finishYN = req("finishYN","")

dim boardqna
set boardqna = New CMyQNA
	boardqna.FPageSize = 50
	boardqna.FCurrPage = page
	boardqna.RectQadiv = qadiv
	boardqna.FSearchUserID = userid
	boardqna.FSearchOrderSerial = orderserial
	boardqna.FSearchWriteId = writeId
	boardqna.FSearchChargeId = chargeid
	boardqna.FRectReplyQADiv = replyqadiv
	boardqna.FSearchUserLevel = userlevel
	boardqna.FSearchDiv = searchDiv
	boardqna.FSearchText = searchText
	boardqna.FRectEvalPoint = evalPoint
	boardqna.FRectIsUsing = isusing

	IF blnDate="on" Then
		boardqna.FSearchStartDate = sDate
		boardqna.FSearchEndDate =eDate
	End IF

	IF ckReplyDate="on" Then
		boardqna.FreplyDate1 = replyDate1
		boardqna.FreplyDate2 =replyDate2
	End IF

	boardqna.FRectItemNotInclude = itemqanotinclude
	boardqna.FRectSiteName = sitename
	boardqna.FRectCurrState = currstate
	boardqna.FRectUserGubun = userGubun

	''boardqna.list finishYN

	''old ver
	'if (finishyn = "N") then
	    boardqna.SearchNew = finishyn
	'end if

	boardqna.fqnalist

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">

$(function() {
  	$("#finishyn_A").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="","font-weight:bold;color:red;","color:black;")%>");
  	$("#finishyn_N").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="N","font-weight:bold;color:red;","color:black;")%>");
  	$("#finishyn_VV").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="V","font-weight:bold;color:red;","color:black;")%>");
	$("#finishyn_VE").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="E","font-weight:bold;color:red;","color:black;")%>");
	$("#finishyn_VD").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="D","font-weight:bold;color:red;","color:black;")%>");
  	$("#finishyn_V").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="V","font-weight:bold;color:red;","color:black;")%>");
	$("#finishyn_E").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="E","font-weight:bold;color:red;","color:black;")%>");
	$("#finishyn_D").button().children().attr("style","font-size:12px;<%=CHKIIF(finishyn="D","font-weight:bold;color:red;","color:black;")%>");
});

function CloseWindow(){
	window.close();
}

function  TnSearch(frm){
	if (frm.rectuserid.length<1){
		alert('�˻�� �Է��ϼ���.');
		return;
	}
	frm.method="get";
	frm.submit();
}

function NextPage(ipage){
	document.frmSrc.page.value= ipage;
	document.frmSrc.submit();
}

function SubmitSearch() {
    document.qnaform.submit();
}

function SubmitSearchUserId(userid) {
    document.qnaform.userid.value = userid;
    document.qnaform.orderserial.value = "";
    document.qnaform.submit();
}


function jsPopCal(fName,sName){
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function EnableDate(obj){
	var f = document.qnaform;
	if (obj.checked)
	{
		f.sdt.readOnly=false;
		f.edt.readOnly=false;
		f.sdt.className="text";
		f.edt.className="text";
	}
	else
	{
		f.sdt.readOnly=true;
		f.edt.readOnly=true;
		f.sdt.className="text_ro";
		f.edt.className="text_ro";
	}
}

function replyEnableDate(obj){
	var f = document.qnaform;
	if (obj.checked)
	{
		f.replyDate1.readOnly=false;
		f.replyDate2.readOnly=false;
		f.replyDate1.className="text";
		f.replyDate2.className="text";
	}
	else
	{
		f.replyDate1.readOnly=true;
		f.replyDate2.readOnly=true;
		f.replyDate1.className="text_ro";
		f.replyDate2.className="text_ro";
	}
}

function jsFinishYNButton(a) {
    document.qnaform.userid.value = "";
    document.qnaform.orderserial.value = "";
    document.qnaform.finishyn.value = a;
    document.qnaform.submit();
}

function jsDelQna(id) {
	if (confirm("�����Ͻðڽ��ϱ�?")) {
		document.delform.id.value = id;
		document.delform.submit();
	}
}

document.title = "1:1 ��㸮��Ʈ";

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="400" style="padding:5; border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>"  background="/images/menubar_1px.gif">
					<font color="#333333"><b>[CS]������>>[1:1���]�Խ��ǰ���</b></font>
				</td>

				<td align="right" style="border-bottom:1px solid <%= adminColor("tablebg") %>" bgcolor="#F4F4F4">
				</td>

			</tr>
		</table>
	</td>
</tr>

<!--	���� ������ ���ϴ�.	-->
<tr bgcolor="#FFFFFF">
	<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
		- �ֱ� 100�Ǹ� �˻��˴ϴ�.<br>
		- ���̵� �Ǵ� �ֹ���ȣ�� �˻��� ���� �亯����/���������� ������� ��� ǥ�õ˴ϴ�.
	</td>
</tr>
</table>

<!-- �˻� ���� -->
<form method="get" name="qnaform" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�� ��<br>�� ��</td>
	<td height="31" align="left">&nbsp;
		�����̵� : <input type="text" class="text" name="userid" value="<%= userid %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;&nbsp;
  		�ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;&nbsp;/&nbsp;&nbsp;
  		����ھ��̵� : <input type="text" class="text" name="chargeid" value="<%= chargeid %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;&nbsp;
  		�亯�ھ��̵� : <input type="text" class="text" name="writeid" value="<%= writeid %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
		&nbsp;&nbsp;
		���θ� :
	    <% call drawSelectBoxXSiteOrderInputPartnerCS("sitename", sitename) %>
		&nbsp;&nbsp;
		���� :
        <select class="select" name="currstate">
        	<option value="" <%=CHKIIF(currstate="","selected","")%>>����</option>
        	<option value="B001" <%=CHKIIF(currstate="B001","selected","")%>>�亯���� ��ü</option>
			<option value="B006" <%=CHKIIF(currstate="B006","selected","")%>>��ü �亯�Ϸ�</option>
			<option value="B007" <%=CHKIIF(currstate="B007","selected","")%>>�亯�Ϸ�</option>
			<option value="B008" <%=CHKIIF(currstate="B008","selected","")%>>��������(����)</option>
			<option value="B009" <%=CHKIIF(currstate="B009","selected","")%>>���ۿϷ�(����)</option>
        </select>
  	</td>
	<td rowspan="4" width="80" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�� ��" style="width:60px;height:70px;" onClick="SubmitSearch()">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td height="31" align="left">&nbsp;
  		�������� :
  		<% drawSelectBoxqadiv "qadiv", qadiv, "", "Y", "N", "Y" %>
        &nbsp;&nbsp;
  		�������� :
		<select class="select" name="replyqadiv">
			<option value="">��ü</option>
			<option value="">======</option>
            <option value="01" <% if replyqadiv = "01" then response.write "selected" %>>�ܼ�����</option>
			<option value="">======</option>
            <option value="all" <% if replyqadiv = "all" then response.write "selected" %>>���Ҹ� ��ü</option>
			<option value="02"  <% if replyqadiv = "02" then response.write "selected" %>>��ü�Ҹ�</option>
            <option value="03"  <% if replyqadiv = "03" then response.write "selected" %>>���(CJ)�Ҹ�</option>
            <option value="10"  <% if replyqadiv = "10" then response.write "selected" %>>�ý��۰�����û</option>
            <option value="99"  <% if replyqadiv = "99" then response.write "selected" %>>��Ÿ�Ҹ�</option>
        </select>
        &nbsp;&nbsp;
        ȸ����� : <% DrawselectboxUserLevel "userlevel", userlevel, "" %>
        &nbsp;&nbsp;
        �������� :
        <select class="select" name="evalPoint">
        	<option value="" <%=CHKIIF(evalPoint="","selected","")%>>��ü</option>
        	<option value="5" <%=CHKIIF(evalPoint="5","selected","")%>>5��</option>
			<option value="4" <%=CHKIIF(evalPoint="4","selected","")%>>4��</option>
			<option value="3" <%=CHKIIF(evalPoint="3","selected","")%>>3��</option>
			<option value="2" <%=CHKIIF(evalPoint="2","selected","")%>>2��</option>
			<option value="1" <%=CHKIIF(evalPoint="1","selected","")%>>1��</option>
			<option>--------</option>
			<option value="3DN" <%=CHKIIF(evalPoint="3DN","selected","")%>>3�� ���� ��ü</option>
        </select>
		&nbsp;&nbsp;
		�˻����� :
        <select class="select" name="searchDiv">
        	<option value="" <%=CHKIIF(searchDiv="","selected","")%>>����</option>
        	<option value="title" <%=CHKIIF(searchDiv="title","selected","")%>>����</option>
			<option value="contents" <%=CHKIIF(searchDiv="contents","selected","")%>>����</option>
			<option value="makerid" <%=CHKIIF(searchDiv="makerid","selected","")%>>�귣��</option>
			<option value="username" <%=CHKIIF(searchDiv="username","selected","")%>>����</option>
        </select>
		<input type="text" class="text" name="searchText" value="<%= searchText %>" size="12" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
		&nbsp;&nbsp;
		�ۼ��� :
        <select class="select" name="userGubun">
			<option></option>
			<option value="C" <%= CHKIIF(userGubun="C", "selected", "") %>>��</option>
			<option value="M" <%= CHKIIF(userGubun="M", "selected", "") %>>����</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td height="31" align="left">&nbsp;
        <input type="checkbox" name="edc" <%IF blnDate="on" then response.write "checked" %> onclick="EnableDate(this);">
        ���ۼ��� : <input type="text" size="10" name="sdt" value="<%= sDate %>" onClick="jsPopCal('qnaform','sdt');" <% IF blnDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        ~<input type="text" size="10" name="edt" value="<%= eDate %>" onClick="jsPopCal('qnaform','edt');" <% IF blnDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        &nbsp;&nbsp;
        <input type="checkbox" name="ckReplyDate" <%IF ckReplyDate="on" then response.write "checked" %> onclick="replyEnableDate(this);">
        �亯�� : <input type="text" size="10" name="replyDate1" value="<%= replyDate1 %>" onClick="jsPopCal('qnaform','replyDate1');" <% IF ckReplyDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        ~<input type="text" size="10" name="replyDate2" value="<%= replyDate2 %>" onClick="jsPopCal('qnaform','replyDate2');" <% IF ckReplyDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
		&nbsp;&nbsp;
		<input type="checkbox" name="isusing" value="Y" <%= CHKIIF(isusing="Y", "checked", "") %> > �������� ǥ�þ���
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td height="31" align="left">&nbsp;
		<input type="hidden" name="finishyn" value="<%=finishyn%>">
		<button type="button" id="finishyn_A" onClick="jsFinishYNButton('');">��ü���</button>
		&nbsp;
		<button type="button" id="finishyn_N" onClick="jsFinishYNButton('N');">��ó�����</button>
		&nbsp;
		<button type="button" id="finishyn_VV" onClick="jsFinishYNButton('VV');">VVIP ��ó����ü</button>
		<button type="button" id="finishyn_VE" onClick="jsFinishYNButton('VE');">VVIP �Ϲݻ��</button>
		<button type="button" id="finishyn_VD" onClick="jsFinishYNButton('VD');">VVIP ��۹���</button>
		&nbsp;
		<button type="button" id="finishyn_V" onClick="jsFinishYNButton('V');">VIP ��ó����ü</button>
		<button type="button" id="finishyn_E" onClick="jsFinishYNButton('E');">VIP �Ϲݻ��</button>
		<button type="button" id="finishyn_D" onClick="jsFinishYNButton('D');">VIP ��۹���</button>
	</td>
</tr>
</table>
</form>

<br>
* <font color="blue">[M]</font> : ����� ����Ʈ���� �ۼ��� �����Դϴ�.<br>
* <font color="orange">[A]</font> : �ۿ��� �ۼ��� �����Դϴ�.
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15" style="padding:3 0 3 5">�˻���� : <b><%=boardqna.ResultCount%></b> / <%=boardqna.TotalCount%></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="70" height="25">����</td>
    <td width="135">����(���̵�)</td>
    <td width="70">����Ʈ</td>
	<td width="70">�ֹ���ȣ</td>
    <td width="120">�귣��</td>
    <td width="50">���ǻ�ǰ</td>
	<td width="100">����</td>
    <td>����</td>
    <td width="30">÷��</td>
	<td width="80">�ۼ���</td>
    <td width="60">�����</td>
    <td width="100">��ü�亯</td>
	<td width="100">�亯����</td>
	<td width="30">����</td>
    <td width="30">����</td>
</tr>
<% if boardqna.ResultCount>0 then %>
	<%
	for i = 0 to boardqna.ResultCount - 1

	if isarray(split(boardqna.results(i).fqadivname,"!@#")) then
		if ubound(split(boardqna.results(i).fqadivname,"!@#")) > 0 then
			tmpqadivname =  split(boardqna.results(i).fqadivname,"!@#")(1)
		end if
	end if
	%>
	<% if (boardqna.results(i).dispyn = "N") then %>
		<tr align="center" bgcolor="#EEEEEE">
	<% else %>
		<tr align="center" bgcolor="#FFFFFF">
	<% end if %>

		<td align="center" height="25">
			<% if (boardqna.results(i).Fsitename = "10x10") or (boardqna.results(i).Fsitename = "") then %>
				<font color="<%= getUserLevelColorByDate(boardqna.results(i).fUserLevel,Left(boardqna.results(i).regdate,10)) %>">
				<strong><%= getUserLevelStrByDate(boardqna.results(i).fUserLevel,Left(boardqna.results(i).regdate,10)) %></strong></font>
			<% end if %>
		</td>
	    <td align="left">
			<% if boardqna.results(i).Fuserlevel="7" then %>
				<font color="blue"><%= boardqna.results(i).username %></font>
		    	<!--<a href="javascript:SubmitSearchUserId('<%'= boardqna.results(i).userid %>');">-->
		    	<font color="blue">(<%= printUserId(boardqna.results(i).userid, 2, "*") %>)</font>
		    	<!--</a>-->
				</font>
			<% else %>
		    	<%= boardqna.results(i).username %>
		    	<!--<a href="javascript:SubmitSearchUserId('<%'= boardqna.results(i).userid %>');">-->
		    	(<%= printUserId(boardqna.results(i).userid, 2, "*") %>)
		    	<!--</a>-->
			<% end if %>
	    </td>
	    <td><%= boardqna.results(i).Fsitename %></td>
		<td><%= boardqna.results(i).orderserial %></td>
	    <td>
	    	<%= boardqna.results(i).Fmakerid %>
	    	<% if (boardqna.results(i).IsUpchebeasong) then %>
	    		<font color=red>(����)</font>
	    	<% end if %>
	    </td>
	    <td><%= boardqna.results(i).itemid %></td>
		<td>
			<a href="cscenter_qna_board_reply.asp?id=<%= boardqna.results(i).id %>">
			<% if boardqna.results(i).qadiv="26" then %>
				<font color="blue"><%= tmpqadivname %></font>
			<% else %>
				<%= tmpqadivname %>
			<% end if %>
			</a>
		</td>
	    <td align="left">
			<% if Not IsNull(boardqna.results(i).FExtSiteName) then %>
				<% if (boardqna.results(i).FExtSiteName = "mobile") then %>
					<font color="blue">[M]</font>
				<% elseif (boardqna.results(i).FExtSiteName = "app") then %>
					<font color="orange">[A]</font>
				<% end if %>
			<% end if %>
			<% if (boardqna.results(i).FEvalPoint > 0) then %>
				<% for j = 1 to boardqna.results(i).FEvalPoint %><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/star_red.gif"><% next %>
			<% end if %>
			<a href="cscenter_qna_board_reply.asp?id=<%= boardqna.results(i).id %>">
				<%= db2html(boardqna.results(i).title) %>
				<% if (boardqna.results(i).title = "") then %>(�������)<% end if %>
			</a>
			<!--
			<a href="cscenter_qna_board_reply_new.asp?id=<%= boardqna.results(i).id %>">
				[������]
			</a>
			-->
		</td>
		<td><%= CHKIIF(boardqna.results(i).FattachFile <> "", "Y", "") %></td>
	    <td align="center">
	    	<%
			' �̹����̻�� ����. ���� ǥ�� ���� ���� �׳� ��¥�� �ð� �ܼ��ϰ� ǥ���϶� �Ͻ�.	' 2019.05.16 �ѿ��
			'if (Left(boardqna.results(i).regdate, 10) < Left(now, 10)) then
			%>
			<% if boardqna.results(i).regdate<>"" and not(isnull(boardqna.results(i).regdate)) then %>
				<%= Left(boardqna.results(i).regdate,10) %>
				<br><%= mid(boardqna.results(i).regdate,11,18) %>
	    	<% end if %>
			<% 'else %>
	      	<!--���� <%'= Right(FormatDate(boardqna.results(i).regdate, "0000.00.00 00:00:00"), 8) %>-->
	    	<% 'end if %>
	    </td>
	    <td>
	    	<% if (boardqna.results(i).chargeid<>"") then %><%= boardqna.results(i).chargeid %><% end if %>
	    </td>
	    <td>
			<% if  boardqna.results(i).FtargetMakerID<>"" and Not IsNull(boardqna.results(i).FtargetMakerID) then %>

				<% if  boardqna.results(i).Fupchereplydate<>"" and Not IsNull(boardqna.results(i).Fupchereplydate) then %>
					<% if boardqna.results(i).replyDate<>"" and not(isnull(boardqna.results(i).replyDate)) then %>
					�亯�Ϸ�<br />
					<% else %>
					<b>�亯�Ϸ�</b><br />
					<% end if %>
					<%= Left(boardqna.results(i).Fupchereplydate,10) %><br />
					<%= mid(boardqna.results(i).Fupchereplydate,11,18) %>
				<% elseif  boardqna.results(i).Fupcheviewdate<>"" and Not IsNull(boardqna.results(i).Fupcheviewdate) then %>
					��üȮ����<br />
					<%= Left(boardqna.results(i).Fupcheviewdate,10) %><br />
					<%= mid(boardqna.results(i).Fupcheviewdate,11,18) %>
				<%  else %>
					<%= boardqna.results(i).FtargetMakerID %><br />
				<% end if %>
			<% end if %>
	    </td>
	    <td>
	    	<% if boardqna.results(i).replyuser<>"" and not(isnull(boardqna.results(i).replyuser)) then %>
				�Ϸ�(<%= boardqna.results(i).replyuser %>)

				<% if boardqna.results(i).replyDate<>"" and not(isnull(boardqna.results(i).replyDate)) then %>
					<br>
					<%= Left(boardqna.results(i).replyDate,10) %>
					<br><%= mid(boardqna.results(i).replyDate,11,18) %>
				<% end if %>
			<% end if %>
	    </td>
	    <td>
	    	<%= boardqna.results(i).FsendYN %>
	    </td>
	    <td>
	    	<% if (boardqna.results(i).dispyn="N") then %>
			<font color="red">����</font>
			<% elseif (boardqna.results(i).Fsitename <> "10x10") and (boardqna.results(i).Fsitename <> "") and (boardqna.results(i).replyuser = "" or isnull(boardqna.results(i).replyuser)) then %>
				<%
					If (session("ssAdminPOsn") = "4") OR (session("ssAdminPOsn") = "5") Then
				 %>
					<a href="javascript:jsDelQna(<%= boardqna.results(i).id %>)">X</a>
				<%
					Else
						response.write "���Ѿ���"
					End If
				%>
			<% end if %>
	    </td>
	</tr>
	<% next %>

    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% sbDisplayPaging "page="&page, boardqna.FTotalCount, boardqna.FPageSize, 10%>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">�˻������ �����ϴ�.</td>
	</tr>
<% end if %>
</table>

<form method="post" name="delform" action="cscenter_qna_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="">
<input type="hidden" name="mode" value="del">
</form>

<%
Set boardqna = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
