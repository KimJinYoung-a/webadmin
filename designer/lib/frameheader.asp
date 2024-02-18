<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<%
dim btcid, grpid
btcid= session("ssBctID")
grpid= session("ssGroupid")
if (btcid="") then response.End

dim noFinCtrExists, isNewContractTypeExists
dim NoConfirmPreContractID : NoConfirmPreContractID=-1

noFinCtrExists = isNotFinishNewContractExists(btcid, grpid, isNewContractTypeExists)

if (Not noFinCtrExists) and (Not isNewContractTypeExists) then
    NoConfirmPreContractID = getLastPrecontractID(btcid)
end if
%>

<html>
<head>
<title>[10x10] Business Comunication</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="JavaScript" src="/js/common.js"></script>
<script language='javascript'>
function WindowMinSize(){
	parent.document.all('menuset').cols = "20,*";
	document.all.WINSIZE[0].style.display = "none";
	document.all.WINSIZE[1].style.display = "";
}

function WindowMaxSize(){
	parent.document.all('menuset').cols = "180,*";
	document.all.WINSIZE[0].style.display = "";
	document.all.WINSIZE[1].style.display = "none";
}

function pop_editcompany(){
	var popwin = window.open('<%=getSCMSSLURL%>/designer/company/editcompany3.asp?menupos=53' ,'op1','width=640,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function pop_10x10_person(){
	var popwin = window.open('<%=getSCMSSLURL%>/common/pop_10x10_person.asp','op2','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function pop_10x10_map(){
	var popwin = window.open('/common/pop_10x10_map.asp','op3','width=650,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}


function ShiftBrand(comp){
	// refere ������ �������� ����
	/*
	var targetFrm = top.contents;


	var o  = targetFrm.document.createElement("form");
    var oi1 = targetFrm.document.createElement("input");

	oi1.type = "hidden";
    oi1.name = "shiftid";
    oi1.value = comp.value;

    o.appendChild(oi1);
    targetFrm.document.body.appendChild(o);

    o.method = "get";
    o.action = "/designer/lib/shiftbrand.asp";



    o.submit();
    */

    top.contents.location.href="/designer/lib/shiftbrand.asp?shiftid="+comp.value;

	//focusing out
	document.location.reload();
	top.menu.document.location.reload();


}

function pop_contract(ContractID){
    if (ContractID<1){
        top.contents.location.href='/designer/company/contract/ctrListBrand.asp?menupos=1623';
    }else{
        var popwin = window.open('/designer/company/popContract.asp?ContractID=' + ContractID,'popContract','width=650,height=800,scrollbars=yes,resizable=yes')
        popwin.focus();
    }
}

<% if NoConfirmPreContractID>0 then %>
    pop_contract('<%= NoConfirmPreContractID %>');
<% end if %>


// �������� �˾�(2008.08.31; ������)
function pop_survey(srvSn) {
    alert("�ٹ����� ��Ʈ�� �е鲲 ���������� �����մϴ�.\n�ڼ��� ������ �˾�â���� Ȯ�����ּ���.\n\n��Ȯ�� �� �˾��� �ȶ߽ø� �˾������� ������ �ּ���.");
    var popSrv = window.open('/designer/board/upche_survey.asp?sn=' + srvSn,'popSurvey','width=705,height=700,scrollbars=yes');
    popSrv.focus();
}
<%
	'/// ���� �������� �������� �������� Ȯ��(��ü�� 1��)
	sqlStr = "exec db_board.dbo.sp_Ten_check_UpcheSurvey '" & grpid & "'"
	rsget.CursorLocation = adUseClient 
	rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly

	if Not(rsget.EOF) then
		if rsget("pollCnt")=0 then
			Response.Write "pop_survey('" & rsget("srv_sn") & "');"
		end if
	end if

	rsget.Close


	'####### ��ǰ����������� �˾� 20121107
	If Now() < #11/19/2012 23:59:59# Then
%>
	    var popSrv20121107 = window.open('http://webadmin.10x10.co.kr/designer/etc/notice_20121107.html','popSrv20121107','width=499,height=530,scrollbars=no');
	    popSrv20121107.focus();
<%
	End If
%>
</script>

</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="0">

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">

	<tr height="35" valign="top">
        <td valign="bottom">
        	&nbsp;
        	<img id="ilogiimage" src="/images/admin_logo_10x10.jpg" width="90" height="25" align="absbottom">
        	<b>10x10 Business Communication Tool</b>
        	<% if (application("Svr_Info")="Dev") then %>
        		<b><font color="red">[This is Test Server...]</font></b>
        	<% end if %>

        </td>
        <td align="right" valign="bottom">
<%
dim sqlStr,i
dim Resultcount


sqlStr = " select top 50 p.id, c.socname, c.socname_kor,c.userdiv" + VbCrlf
sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p" + VbCrlf
sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on p.id=c.userid" + VbCrlf
sqlStr = sqlStr + " where p.groupid='" & grpid & "'" + VbCrlf
sqlStr = sqlStr + " and p.userdiv='9999'" + VbCrlf
sqlStr = sqlStr + " and p.isusing='Y'" + VbCrlf
sqlStr = sqlStr + " and c.userdiv<15"  ''����ó��.. => �������� (14)
''sqlStr = sqlStr + " and c.isusing='Y'" + VbCrlf

rsget.Open sqlStr,dbget,1

	if not rsget.Eof then
		Resultcount = rsget.RecordCount
%>
        	<select class="select" name="brandshift" onChange="ShiftBrand(this)">
        	<% for i=0 to Resultcount - 1 %>
        	<option value="<%= rsget("id") %>" <% if (LCase(rsget("id"))=LCase(session("ssBctId"))) then response.write "selected" %> ><%= rsget("id") %> (<%= db2html(rsget("socname_kor")) %> <%= CHKIIF(rsget("userdiv")="14","-���ΰŽ�","") %>)
        	<% rsget.MoveNext %>
        	<% next %>
        	</select>
        	&nbsp;
<%
	end if
rsget.Close
%>
        	<a href="javascript:pop_editcompany('<%= menupos %>');" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">��ü �� �귣����������</a>
	        |
	        <a href="javascript:pop_contract('<%=NoConfirmPreContractID%>');" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">��ü��༭ �ٿ�ε�</a>
	        |
	        <a href="javascript:pop_10x10_person();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">��Ʈ�� �����</a>
	        |
	        <a href="javascript:pop_10x10_map();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">�ٹ����� �൵</a>
    		|
        	<a href="#" onclick="printbarcode_on_off_multi(); return false;" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'" >���ڵ����</a>
	        &nbsp;
        </td>
	</tr>
	<tr height="5" valign="top">
        <td colspan="10"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#CCCCCC">
    <tr height="20"  valign="top">
        <td width="175" align="right" valign="middle">
			<div id=WINSIZE style="display:">â Ȯ���ϱ�
				<input type="button" class="button" value="��" onClick="javascript:WindowMinSize()">
			</div>
			<div id=WINSIZE style="display:none">â ����ϱ�
				<input type=button class="button" value="��" onClick="javascript:WindowMaxSize()">
			</div>
		</td>
        <td align="right" valign="middle">
	        <b><%=session("ssBctID")%>(<%=session("ssBctCname")%>)</b> ���� �α��� �ϼ̽��ϴ�.
	    	&nbsp;
	    	<a href="/login/dologout.asp" target="_top"><img src="/images/icon_logout.gif" width="64" height="17" border="0" align="absbottom"></a>
	    	&nbsp;
		        &nbsp;
        </td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
