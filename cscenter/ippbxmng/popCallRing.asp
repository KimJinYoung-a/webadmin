<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs �޸� -CALL �߰�
' History : 2007.10.26 �ѿ�� ����
'           2009-01-07 ������ ����,
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%

1 �ʾ���.!!!!!!

dim i, userid, orderserial,id
dim mode, sqlStr
dim ippbxuser, intel, phoneNumber, phoneNumberOut, qadiv
dim isEditMode


ippbxuser   = requestCheckVar(request("ippbxuser"),32)
intel       = requestCheckVar(request("intel"),32)
phoneNumber      = requestCheckVar(request("phoneNumber"),32)
phoneNumberOut  = requestCheckVar(request("phoneNumberOut"),32)

if (phoneNumber<>"") then phoneNumber=ParsingPhoneNumber(phoneNumber)

userid          = RequestCheckVar(request("userid"),32)
orderserial     = RequestCheckVar(request("orderserial"),11)
id              = RequestCheckVar(request("id"),9)




dim ocsmemo
set ocsmemo = New CCSMemo

if (id <> "") then
	ocsmemo.FRectId = id
	ocsmemo.FRectUserID = userid
	ocsmemo.FRectOrderserial = orderserial
	ocsmemo.GetCSMemoDetail

	userid = ocsmemo.FOneItem.FUserID
	orderserial = ocsmemo.FOneItem.Forderserial
	phoneNumber = ocsmemo.FOneItem.FphoneNumber

	isEditMode = true
else
	ocsmemo.GetCSMemoBlankDetail
	''mayBe Inbound
	if (phoneNumber<>"") then ocsmemo.FOneItem.FmmGubun = "1"
	isEditMode = false
end if




'=============================================================================
%>
<script language='javascript'>
var NowDoing = false;
<% if (phoneNumber<>"") or (orderserial<>"") or (userid<>"") then %>
    NowDoing = true;
<% end if %>
function setDoingState(){
    document.all.doingdispinfo.innerHTML = (NowDoing)?"<strong><font color=red>[ó����]</font></strong>":"[�����]";
}

function setGubunState(){
    var comp = frm.mmGubun;

    if (comp.value == "0") {
        //�Ϲݸ޸�
        frm.phoneNumber.disabled = true;
        frm.phoneNumber.style.background = "#DDDDDD";

        frm.phoneNumberOut.disabled = true;
        frm.phoneNumberOut.style.background = "#DDDDDD";

    }else if(comp.value=="1"){
        //�ιٿ��
        frm.phoneNumber.disabled = false;
        frm.phoneNumber.style.background = "#FFFFFF"; //className="text";

        frm.phoneNumberOut.disabled = true;
        frm.phoneNumberOut.style.background = "#DDDDDD"; //className="text_ro";
    }else if(comp.value=="2"){
        //�ƿ��ٿ��
        frm.phoneNumber.disabled = true;
        frm.phoneNumber.style.background = "#DDDDDD";

        frm.phoneNumberOut.disabled = false;
        frm.phoneNumberOut.style.background = "#FFFFFF";

    }else if(comp.value=="3"){
        //��ü��ȭ
        frm.phoneNumber.disabled = true;
        frm.phoneNumber.style.background = "#DDDDDD";

        frm.phoneNumberOut.disabled = false;
        frm.phoneNumberOut.style.background = "#FFFFFF";
    }
}

function checkDoing(){
    if (!NowDoing){
        NowDoing=true;
        setDoingState();
    }
}

function reInput(){
    document.location.href = '/cscenter/ippbxmng/popCallRing.asp';
}

function Clip2Paste(){
    var clipTxt = window.clipboardData.getData("Text");

    if (clipTxt.length<1){ return; }

    //indexOf
    var posSpliter = clipTxt.indexOf("|");
    var iorderserial ="";
    var iuserid ="";
    if (posSpliter>0){
        iorderserial = clipTxt.substring(0,posSpliter);
        iuserid      = clipTxt.substring(posSpliter+1,255);

        frm.orderserial.value = iorderserial;
        frm.userid.value = iuserid;
    }


}


function SearchOrderByPhoneNo(comp){
    var iphoneNum = comp.value;
    if (iphoneNum.length<1){
        alert('��ȭ��ȣ�� �ְ� �˻��ϼ���.');
        if (comp.enabled) { comp.focus(); };
        return;
    }

    var isNewWin = false;
    var window_width = 1280;
    var window_height = 1024;

    try{
        opener.Hndlw;
    }catch(e){
        //
        alert('â�� ������ �� �ٽ� �õ��� �ּ���.');
        return;
    }

    if (opener.Hndlw==null){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//â�� ������� üũ
	try{
	    opener.Hndlw.focus();
	}catch(e){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//this.focus();
	//�� ��ũ�γ����� ����.. setTimeout();
	//opener.Hndlw.listFrame.SearchByPhoneNumber(iphoneNum); //��ũ�� ����..

	if (isNewWin){
	    setTimeout("opener.Hndlw.listFrame.SearchByPhoneNumber('" + iphoneNum + "')",1000);
	}else{
	    opener.Hndlw.listFrame.SearchByPhoneNumber(iphoneNum);
	}

}

function SearchOrderByOrderSerial(comp){
    var iOrderserial = comp.value;
    if (iOrderserial.length<1){
        alert('�ֹ���ȣ�� �ְ� �˻��ϼ���.');
        if (comp.enabled) { comp.focus(); };
        return;
    }

    var isNewWin = false;
    var window_width = 1280;
    var window_height = 1024;

    try{
        opener.Hndlw;
    }catch(e){
        //
        alert('â�� ������ �� �ٽ� �õ��� �ּ���.');
        return;
    }

    if (opener.Hndlw==null){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//â�� ������� üũ
	try{
	    opener.Hndlw.focus();
	}catch(e){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//this.focus();
	//�� ��ũ�γ����� ����.. setTimeout();
	//opener.Hndlw.listFrame.SearchByOrderserial(iOrderserial); //��ũ�� ����..

	if (isNewWin){
	    setTimeout("opener.Hndlw.listFrame.SearchByOrderserial('" + iOrderserial + "')",1000);
	}else{
	    opener.Hndlw.listFrame.SearchByOrderserial(iOrderserial);
	}

}

function SearchOrderByUserID(comp){
    var iUserid = comp.value;
    if (iUserid.length<1){
        alert('���̵� �ְ� �˻��ϼ���.');
        if (comp.enabled) { comp.focus(); };
        return;
    }

    var isNewWin = false;
    var window_width = 1280;
    var window_height = 1024;

    try{
        opener.Hndlw;
    }catch(e){
        //
        alert('â�� ������ �� �ٽ� �õ��� �ּ���.');
        return;
    }

    if (opener.Hndlw==null){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//â�� ������� üũ
	try{
	    opener.Hndlw.focus();
	}catch(e){
	    opener.Hndlw = window.open("/cscenter/ordermaster/ordermaster.asp","PopOrderMaster","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	    opener.Hndlw.focus();
	    isNewWin=true;
	}

	//this.focus();
	//�� ��ũ�γ����� ����.. setTimeout();
	//opener.Hndlw.listFrame.SearchByUserID(iUserid); //��ũ�� ����..

	if (isNewWin){
	    setTimeout("opener.Hndlw.listFrame.SearchByUserID('" + iUserid + "')",1000);
	}else{
	    opener.Hndlw.listFrame.SearchByUserID(iUserid);
	}

}

function fnClick2Call(comp){
    var iphoneNum = comp.value;
    if (iphoneNum.length<1){
        alert('��ȭ��ȣ�� �Է��ϼ���.');
        if (!comp.disabled) { comp.focus(); };
        return;
    }

    //��� �߰� ��ƾ �ʿ� js��


    if (!opener){
        alert('Err1 - â�� �ٽ� �����ּ���..');
        return;
    }

    //���� ����.. �θ�â�� ������� �� ����..
    try{
        opener.name;
    }catch(e){
        alert('â�� ������ �� �ٽ� �õ��� �ּ���.');
        return;
    }

    opener.click2call(iphoneNum);
}

function iMemoList(comp){
    var iphoneNum    = "";
    var iuserid      = "";
    var iorderserial = "";

    if ((comp.name=="phoneNumber")||(comp.name=="phoneNumberOut")){
        iphoneNum = comp.value;
        if (iphoneNum.length<1){
            alert('��ȭ��ȣ�� �Է��ϼ���.');
            if (!comp.disabled) { comp.focus(); };
            return;
        }
    }else if (comp.name=="userid"){
        iuserid = comp.value;
        if (iuserid.length<1){
            alert('���̵� �Է��ϼ���.');
            comp.focus();
            return;
        }
    }else if (comp.name=="orderserial"){
        iorderserial = comp.value;
        if (iorderserial.length<1){
            alert('�ֹ���ȣ��  �Է��ϼ���.');
            comp.focus();
            return;
        }
    }

    document.all.i_history_memo.src = "/cscenter/ippbxmng/iframeHistory.asp?userid=" + iuserid + "&orderserial=" + iorderserial + "&phoneNumer=" + iphoneNum;
}

function GotoHistoryMemoMidify(id,userid,orderserial)
{
    frm.action="/cscenter/history/history_memo_write.asp?id=" + id + "&backwindow=" + "opener" + "&userid=" + userid + "&orderserial=" + orderserial
    frm.submit();
}

function SubmitSave()
{
    if ((document.frm.orderserial.value.length<1)&&(document.frm.userid.value.length<1)&&(document.frm.phoneNumber.value.length<1)) {
	    alert("��ȭ��ȣ, �ֹ���ȣ, ���̵� �� �ϳ��� �Է� �Ǿ�� �մϴ�.");
		return;
	}

	if (document.frm.contents_jupsu.value == "") {
		alert("�޸𳻿��� �Է��ϼ���.");
		document.frm.contents_jupsu.focus();
		return;
	}

	if (document.frm.qadiv.value.length<1){
	    alert("���� ������ ���� �ϼ���.");
		document.frm.qadiv.focus();
		return;
	}

	if(document.frm.id.value == "") {
    	document.frm.mode.value = "write";
    	document.frm.submit();
	}else{
    	document.frm.mode.value = "modify";
    	document.frm.submit();
	}
}

function SubmitFinish(){
	if (document.frm.contents_jupsu.value == "") {
			alert("�޸𳻿��� �Է��ϼ���.");
			return;
	}

    if (confirm("�Ϸ�ó���ϰڽ��ϱ�?") == true) {
            document.frm.mode.value = "finish";
            document.frm.submit();
    }
}

function SubmitDelete()
{
    if (confirm("�����ϰڽ��ϱ�?") == true) {
            document.frm.mode.value = "delete";
            document.frm.submit();
    }
}


</script>
<body topmargin=10 leftmargin=10 marginwidth=0 marginheight=0>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS�޸� - CALL <DIV id=pindispinfo style='display:inline;border:solid 0 gray;font-size:9pt;height:20px;text-align:center'>[ ]</div></b>
        	<input type="button" class="button" value="�ű��Է�" onclick="javascript:reInput();">
            <DIV id=doingdispinfo style='display:inline;border:solid 0 gray;font-size:9pt;height:20px;text-align:center'></div>
        </td>
        <td align="right">

            <input type="button" class="button" value="<%= chkIIF(isEditMode,"����","����") %>" onclick="javascript:SubmitSave();">
	       	<input type="button" class="button" value="�Ϸ�" <%= chkIIF((Not isEditMode) or (ocsmemo.FOneItem.Fdivcd<>"2"),"disabled","") %> onclick="javascript:SubmitFinish();">
	        <input type="button" class="button" value="����" <%= chkIIF(isEditMode,"","disabled") %> onclick="javascript:SubmitDelete();">
	        <input type="button" class="button" value="�ݱ�" onclick="javascript:window.close();">
	    </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" method="post" action="popCallRing_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="id" value="<%= ocsmemo.FOneItem.Fid %>">
	<tr>
    	<td width="50" bgcolor="<%= adminColor("tabletop") %>">����</td>
    	<td bgcolor="#FFFFFF">
	        <select name="mmGubun" onChange="setGubunState(this);">
	            <option value="0" <% if ocsmemo.FOneItem.FmmGubun = "0" then %>selected<% end if %>>�Ϲݸ޸�</option>
	            <option value="1" <% if ocsmemo.FOneItem.FmmGubun = "1" then %>selected<% end if %>>�ιٿ����ȭ</option>
	            <option value="2" <% if ocsmemo.FOneItem.FmmGubun = "2" then %>selected<% end if %>>�ƿ��ٿ����ȭ</option>
	            <option value="3" <% if ocsmemo.FOneItem.FmmGubun = "3" then %>selected<% end if %>>��ü��ȭ</option>
	            <!--
	            <option value="4" <% if ocsmemo.FOneItem.FmmGubun = "4" then %>selected<% end if %>>SMS</option>
	            <option value="5" <% if ocsmemo.FOneItem.FmmGubun = "5" then %>selected<% end if %>>EMAIL</option>
	            -->
	        </select>
        </td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">In��ȭ</td>
    	<td bgcolor="#FFFFFF">
        	<table width="400" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="phoneNumber" class="text" value="<%= phoneNumber %>" size="20" onKeyDown="checkDoing();" onKeyPress="if (event.keyCode == 13) SearchOrderByPhoneNo(frm.phoneNumber);"></td>
        	    <td width="100" align="center"><a href="javascript:SearchOrderByPhoneNo(frm.phoneNumber);">�ֹ��˻�</a></td>
        	    <td width="100" align="center"><!-- a href="javascript:fnClick2Call(frm.phoneNumber);">��ȭ�ɱ�</a --></td>
        	    <td width="100" align="center"><a href="javascript:iMemoList(frm.phoneNumber);">���ø޸�</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">Out��ȭ</td>
    	<td bgcolor="#FFFFFF">
        	<table width="400" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="phoneNumberOut" class="text" value="<%= phoneNumberOut %>" size="20" onKeyDown="checkDoing();" onKeyPress="if (event.keyCode == 13) SearchOrderByPhoneNo(frm.phoneNumber);"></td>
        	    <td width="100" align="center"><a href="javascript:SearchOrderByPhoneNo(frm.phoneNumberOut);">�ֹ��˻�</a></td>
        	    <td width="100" align="center"><a href="javascript:fnClick2Call(frm.phoneNumberOut);">��ȭ�ɱ�</a></td>
        	    <td width="100" align="center"><a href="javascript:iMemoList(frm.phoneNumberOut);">���ø޸�</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
    	<td bgcolor="#FFFFFF">
    	    <table width="400" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="orderserial" class="text" value="<%= orderserial %>" size="20" onKeyDown="checkDoing();" ></td>
        	    <td width="100" align="center"><a href="javascript:SearchOrderByOrderSerial(frm.orderserial)">�ֹ��˻�</a></td>
        	    <td width="100" align="center"><a href="javascript:Clip2Paste()">�ٿ��ֱ�</a></td>
        	    <td width="100" align="center"><a href="javascript:iMemoList(frm.orderserial);">���ø޸�</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">��ID</td>
    	<td bgcolor="#FFFFFF">
    	    <table width="400" cellpadding="0" cellspacing="0" border="0" >
        	<tr>
        	    <td width="100"><input type="text" name="userid" class="text" value="<%= userid %>" size="20" onKeyDown="checkDoing();"></td>
        	    <td width="100" align="center"><a href="javascript:SearchOrderByUserID(frm.userid)">�ֹ��˻�</a></td>
        	    <td width="100" align="center"></td>
        	    <td width="100" align="center"><a href="javascript:iMemoList(frm.userid);">���ø޸�</a></td>
        	</tr>
        	</table>
    	</td>
    </tr>
    <% if id = "" then %>
    <% else %>
	    <tr>
	    	<td bgcolor="<%= adminColor("tabletop") %>">����<br>��</td>
	    	<td bgcolor="#FFFFFF">
	    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.fregdate %>" size="26" readonly>&nbsp;
	    		�����ID : <%= ocsmemo.FOneItem.Fwriteuser %>
	    	</td>
	    </tr>
	<% end if %>
	<% if ucase(ocsmemo.FOneItem.Ffinishyn) <> "Y" then %>
    <% else %>
	    <tr>
	    	<td bgcolor="<%= adminColor("tabletop") %>">�Ϸ���</td>
	    	<td bgcolor="#FFFFFF">
	    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.Ffinishdate %>" size="26" readonly>&nbsp;
	    		�����ID : <%= ocsmemo.FOneItem.Ffinishuser %>
	    	</td>
	    </tr>
	<% end if %>
	<tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
    	<td bgcolor="#FFFFFF">

    	    <% if ocsmemo.FOneItem.Fdivcd="2" then %>
    	    <input type=hidden name="divcd" value="2">
    	    <input type="checkbox" name="dummi" checked disabled >ó����û
    	    <% else %>
    	    <input type="checkbox" name="divcd" value="2" >ó����û
    	    <% end if %>

	        <!-- ���� : -->
	        &nbsp;&nbsp;
  			<select class="select" name="qadiv">
                <option value="">��ü</option>
                <option value="00" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="00","selected","") %> >��۹���</option>
                <option value="01" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="01","selected","") %> >�ֹ�����</option>
                <option value="02" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="02","selected","") %> >��ǰ����</option>
                <option value="03" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="03","selected","") %> >�����</option>
                <option value="04" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="04","selected","") %> >��ҹ���</option>
                <option value="05" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="05","selected","") %> >ȯ�ҹ���</option>
                <option value="06" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="06","selected","") %> >��ȯ����</option>
                <option value="07" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="07","selected","") %> >AS����</option>
                <option value="08" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="08","selected","") %> >�̺�Ʈ����</option>
                <option value="09" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="09","selected","") %> >������������</option>
                <option value="10" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="10","selected","") %> >�ý��۹���</option>
                <option value="11" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="11","selected","") %> >ȸ����������</option>
                <option value="12" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="12","selected","") %> >ȸ����������</option>
                <option value="13" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="13","selected","") %> >��÷����</option>
                <option value="14" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="14","selected","") %> >��ǰ����</option>
                <option value="15" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="15","selected","") %> >�Աݹ���</option>
                <option value="16" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="16","selected","") %> >�������ι���</option>
                <option value="17" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="17","selected","") %> >����/���ϸ�������</option>
                <option value="18" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="18","selected","") %> >�����������</option>
                <option value="20" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="20","selected","") %> >��Ÿ����</option>
            </select>
	    </td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">�޸�<br>����</td>
    	<td bgcolor="#FFFFFF"><textarea name="contents_jupsu" class="textarea" cols="68" rows="10" onKeyPress="checkDoing();"><%= db2html(ocsmemo.FOneItem.Fcontents_jupsu) %></textarea></td>
    </tr>

</table>

<p>

���� �޸�
<br>
<iframe id="i_history_memo" name="i_history_memo" src="/cscenter/ippbxmng/iframeHistory.asp?userid=<%= userid %>&orderserial=<%= orderserial %>&phoneNumer=<%= phoneNumber %>" width="480" height="300" scrolling="auto" frameborder="1"></iframe>


<script language='javascript'>
function getOnLoad(){
    alert('��� ���� �����Դϴ�. ������ ���� ���');
    setDoingState();
    setGubunState();
    document.all.pindispinfo.innerHTML = "[" + window.name.substr(18,9) + "]";
}

window.onload = getOnLoad;
</script>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->