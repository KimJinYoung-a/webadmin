<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �¶��� 1:1 �Խ��� ���� ����
' Hieditor : 2010.01.03 �ѿ�� �¶��� �̵� ����/����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%
dim itemqanotinclude, research, finishyn, userid, orderserial, qadiv , writeid ,i, j ,sDate ,eDate ,blnDate
Dim ckReplyDate, replyDate1, replyDate2 ,page ,boardqna ,ocheck ,shopflg
	qadiv               = request("qadiv")
	itemqanotinclude    = request("itemqanotinclude")
	research            = request("research")
	userid              = request("userid")
	orderserial         = request("orderserial")
	shopflg             = request("shopflg")
	writeid             = request("writeid")
	page	= req("page",1)
	finishYN = req("finishYN","")
	'if (itemqanotinclude="") and (research="") then itemqanotinclude="on"
	
	qadiv = "16"	'�������ι��Ǹ� ����
	
	sDate = request("sdt")
	eDate = request("edt")
	blnDate = request("edc")

	ckReplyDate	= req("ckReplyDate",req("ckReplyDateDefault",""))
	replyDate1	= req("replyDate1",LEFT(CStr(dateAdd("d",-7,now())),10))
	replyDate2	= req("replyDate2",LEFT(CStr(now()),10))
	
	if (blnDate="") and (research="") then 
	    blnDate = "on"
	    sDate   = LEFT(CStr(dateAdd("m",-3,now())),10)
	    eDate   = LEFT(CStr(now()),10)
	end if

'//����üũ
set ocheck = new CMyQNA_list
	ocheck.frectssBctId = session("ssBctId")
	ocheck.fmembercheck()
	
set boardqna = New CMyQNA
	boardqna.FPageSize = 50
	boardqna.FCurrPage = page
	boardqna.RectQadiv = qadiv
	boardqna.FSearchUserID = userid
	boardqna.FSearchOrderSerial = orderserial
	boardqna.FSearchWriteId = writeId
	
	IF blnDate="on" Then
		boardqna.FSearchStartDate = sDate
		boardqna.FSearchEndDate =eDate
	End IF
	
	IF ckReplyDate="on" Then
		boardqna.FreplyDate1 = replyDate1
		boardqna.FreplyDate2 =replyDate2
	End IF
	
	boardqna.FRectItemNotInclude = itemqanotinclude
	
	''boardqna.list finishYN
	
	''old ver
	if (finishyn = "N") then
	    boardqna.SearchNew = "Y"
	end if
	
	if ocheck.foneitem.getmemberdisp = false then	
	    shopflg = "Y" ''������ ���常 ���� - ������ ���� �� �� ����.// 2012/06/18 eastone
		if ocheck.FOneItem.getmemberofficedisp = true then
			boardqna.frectshopid = ocheck.FOneItem.fbigo
		else
			boardqna.frectshopid = ocheck.FOneItem.fssBctId
		end if
	end if
	
	boardqna.frectshopflg = shopflg
	boardqna.list
%>

<script language='javascript'>

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


function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);
	
	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function EnableDate(obj)
{
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

function replyEnableDate(obj)
{
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

document.title = "1:1 ��㸮��Ʈ";

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form method="get" name="qnaform">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		���̵� : <input type="text" class="text" name="userid" value="<%= userid %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;
  		�ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;
  		�亯���̵� : <input type="text" class="text" name="writeid" value="<%= writeid %>" size="15" onKeyPress="if (event.keyCode == 13) document.qnaform.submit();">
  		&nbsp;
  		�������� :
		<select class="select" name="qadiv">
            <option value="">��ü</option>
            <option value="00" <% if qadiv = "00" then response.write "selected" %>>��۹���</option>
            <option value="01" <% if qadiv = "01" then response.write "selected" %>>�ֹ�����</option>
            <option value="02" <% if qadiv = "02" then response.write "selected" %>>��ǰ����</option>
            <option value="03" <% if qadiv = "03" then response.write "selected" %>>�����</option>
            <option value="04" <% if qadiv = "04" then response.write "selected" %>>��ҹ���</option>
            <option value="05" <% if qadiv = "05" then response.write "selected" %>>ȯ�ҹ���</option>
            <option value="06" <% if qadiv = "06" then response.write "selected" %>>��ȯ����</option>
            <option value="07" <% if qadiv = "07" then response.write "selected" %>>AS����</option>
            <option value="08" <% if qadiv = "08" then response.write "selected" %>>�̺�Ʈ����</option>
            <option value="09" <% if qadiv = "09" then response.write "selected" %>>������������</option>
            <option value="10" <% if qadiv = "10" then response.write "selected" %>>�ý��۹���</option>
            <option value="11" <% if qadiv = "11" then response.write "selected" %>>ȸ����������</option>
            <option value="12" <% if qadiv = "12" then response.write "selected" %>>ȸ����������</option>
            <option value="13" <% if qadiv = "13" then response.write "selected" %>>��÷����</option>
            <option value="14" <% if qadiv = "14" then response.write "selected" %>>��ǰ����</option>
            <option value="15" <% if qadiv = "15" then response.write "selected" %>>�Աݹ���</option>
            <option value="16" <% if qadiv = "16" then response.write "selected" %>>�������ι���</option>
            <option value="17" <% if qadiv = "17" then response.write "selected" %>>����/���ϸ�������</option>
            <option value="18" <% if qadiv = "18" then response.write "selected" %>>�����������</option>
            <option value="20" <% if qadiv = "20" then response.write "selected" %>>��Ÿ����</option>
            <option value="21" <% if qadiv = "21" then response.write "selected" %>>���̶�ҹ���</option>
            <option value="23" <% if qadiv = "23" then response.write "selected" %>>����ǰ����</option>
            <option value="24" <% if qadiv = "24" then response.write "selected" %>>POINT1010����</option>
        </select>
        <br>
        <input type="checkbox" name="edc" <%IF blnDate="on" then response.write "checked" %> onclick="EnableDate(this);">
        ���ۼ��� : <input type="text" size="10" name="sdt" value="<%= sDate %>" onClick="jsPopCal('qnaform','sdt');" <% IF blnDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        ~<input type="text" size="10" name="edt" value="<%= eDate %>" onClick="jsPopCal('qnaform','edt');" <% IF blnDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        <input type="checkbox" name="ckReplyDate" <%IF ckReplyDate="on" then response.write "checked" %> onclick="replyEnableDate(this);">
        �亯�� : <input type="text" size="10" name="replyDate1" value="<%= replyDate1 %>" onClick="jsPopCal('qnaform','replyDate1');" <% IF ckReplyDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        ~<input type="text" size="10" name="replyDate2" value="<%= replyDate2 %>" onClick="jsPopCal('qnaform','replyDate2');" <% IF ckReplyDate="on" Then%>class="text" <%Else%>readonly class="text_ro"<%END IF%> style="cursor:hand;">
        ��������:
        <select name="shopflg">
        	<option value="" <% if shopflg = "" then response.write " selected" %>>����</option>
        	<option value="Y" <% if shopflg = "Y" then response.write " selected" %>>���������Ϸ�</option>
        	<option value="N" <% if shopflg = "N" then response.write " selected" %>>���������</option>
        </select>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitSearch()">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="radio" name="finishyn" value="" <% if finishyn = "" then response.write "checked" %>> ��ü
    	<input type="radio" name="finishyn" value="N" <% if finishyn = "N" then response.write "checked" %>> ��ó��
	</td>
</tr>
</form>
</table>

<br>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="20" style="padding:3 0 3 5">�˻���� : <b><%=boardqna.ResultCount%></b> / <%=boardqna.TotalCount%></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
    <td>����(���̵�)</td>
    <td>�ֹ���ȣ</td>
    <td>���ǻ�ǰ</td>
    <td>����</td>
    <td>����</td>
    <td>�ۼ���</td>
    <td>�亯����</td>
    <td>�亯��</td>
    <td>���</td>
</tr>
<% for i = 0 to (boardqna.ResultCount - 1) %>

<% if (boardqna.results(i).dispyn = "N") then %>
<tr align="center" bgcolor="#EEEEEE">
<% else %>
<tr align="center" bgcolor="#FFFFFF">
<% end if %>
	<td><b><%= getUserLevelStrByDate(boardqna.results(i).fUserLevel, Left(boardqna.results(i).regdate, 10)) %></b></td>
    <td>
    	<%= boardqna.results(i).username %>
    	<!--<a href="javascript:SubmitSearchUserId('<%'= boardqna.results(i).userid %>');">-->
    	(<%= printUserId(boardqna.results(i).userid, 2, "*") %>)
    	<!--</a>-->
    	</td>
    <td><%= boardqna.results(i).orderserial %></td>
    <td><%= boardqna.results(i).itemid %></td>
    <td align="left"><%= db2html(boardqna.results(i).title) %></td>
    <td>
    	<%= boardqna.code2name(boardqna.results(i).qadiv) %>
    	<% if boardqna.results(i).fshopid = "" or isnull(boardqna.results(i).fshopid) then %>
    		(���������ȵ�)
    	<% else %>
    		(<%= boardqna.results(i).fshopid %>)
    	<% end if %>
    </td>
    <td align="center">
    	<% if (Left(boardqna.results(i).regdate, 10) < Left(now, 10)) then %>
      	<%= Left(boardqna.results(i).regdate,10) %>
    	<% else %>
      	���� <%= Right(FormatDate(boardqna.results(i).regdate, "0000.00.00 00:00:00"), 8) %>
    	<% end if %>
    </td>
    <td>
    	<% if (boardqna.results(i).replyuser<>"") then %>�亯�Ϸ�<% end if %>
    </td>
    <td>
    	<% if (boardqna.results(i).replyuser<>"") then %><%= boardqna.results(i).replyuser %><% end if %>
    </td>

    <td>
    	<input type="button" value="����" class="button" onclick="location.href='online_cscenter_qna_reply.asp?id=<%= boardqna.results(i).id %>';">
    	<% if (boardqna.results(i).dispyn="N") then %><font color="red">����</font><% end if %>
    </td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=20>	
		<div align="center">
			<% sbDisplayPaging "page="&page, boardqna.FTotalCount, boardqna.FPageSize, 10%>
		</div>
	</td>
</tr>
</table>

<%
Set boardqna = Nothing
set ocheck = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
