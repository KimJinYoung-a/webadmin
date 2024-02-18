<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<!-- #include virtual="/admin/upchejungsan/upchejungsan_function.asp"-->
<%

Dim ipFileNo : ipFileNo=requestCheckVar(request("ipFileNo"),10)
Dim targetGbn : targetGbn=requestCheckVar(request("targetGbn"),32)
Dim frmName : frmName=requestCheckVar(request("frmName"),32)
Dim ipFileState : ipFileState=requestCheckVar(request("ipFileState"),10)
Dim DetailVewTp : DetailVewTp=requestCheckVar(request("DetailVewTp"),10)
Dim dvType  : dvType=requestCheckVar(request("dvType"),10)

Dim intLoop
Dim arrList

Dim sqlStr, ipFileName

' sqlStr = "select top 100 M.ipFileNo,M.ipFileName,M.ipFileRegdate,M.ipFileState,M.ipFileGbn"
' sqlStr = sqlStr & " ,(select count(*) as CNT from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D where D.ipFileNo=M.ipFileNo) as CNT"
' sqlStr = sqlStr & " ,(select count(*) as CNT from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D where D.ipFileNo=M.ipFileNo and D.ipFileDetailState=7 ) as ipkumCNT"
' sqlStr = sqlStr & " ,IsNULL((select Sum(ub_totalsuplycash+ me_totalsuplycash+wi_totalsuplycash+sh_totalsuplycash+et_totalsuplycash+dlv_totalsuplycash) "
' sqlStr = sqlStr & "     from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail P"
' sqlStr = sqlStr & " 	Join db_jungsan.dbo.tbl_designer_jungsan_master J"
' sqlStr = sqlStr & " 	on P.ipFileNo=M.ipFileNo"
' sqlStr = sqlStr & " 	and P.targetGbn='ON'"
' sqlStr = sqlStr & " 	and P.targetIdx=J.id),0) as onJSum"
' sqlStr = sqlStr & " ,IsNULL((select Sum(tot_jungsanprice) "
' sqlStr = sqlStr & "     from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail P"
' sqlStr = sqlStr & " 	Join db_jungsan.dbo.tbl_off_jungsan_master J"
' sqlStr = sqlStr & " 	on P.ipFileNo=M.ipFileNo"
' sqlStr = sqlStr & " 	and P.targetGbn='OF'"
' sqlStr = sqlStr & " 	and P.targetIdx=J.idx),0) as offJSum"
' sqlStr = sqlStr & " ,(select count(*) from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D where M.ipFileNo=D.ipFileNo and erpLinkType is Not NULL) as sendedCNT"
' sqlStr = sqlStr & " ,M.jgubun,M.reqDate"
' sqlStr = sqlStr & " From db_jungsan.dbo.tbl_jungsan_ipkumFile_Master M"
' sqlStr = sqlStr & " where 1=1"
' if (ipFileNo<>"") then
'    sqlStr = sqlStr & " and M.ipFileNo="&ipFileNo
' end if
' if (ipFileState<>"") then
'    if (ipFileState="-1") then
'         sqlStr = sqlStr & " and M.ipFileState<1"
'    else
'         sqlStr = sqlStr & " and M.ipFileState="&ipFileState
'    end if
' end if
' if (targetGbn<>"") then
'    sqlStr = sqlStr & " and M.ipfileGbn='"&targetGbn&"'"
' end if

' sqlStr = sqlStr & " order by M.ipFileNo desc"

sqlStr ="exec [db_jungsan].[dbo].[usp_Ten_jungsanFixedMasterListTop] "&CHKIIF(ipFileNo="","NULL",ipFileNo)&","&CHKIIF(ipFileState="","NULL",ipFileState)&",'"&targetGbn&"'"
'response.write sqlStr & "<br>"
'response.end
rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
IF Not rsget.Eof THEN
    arrList = rsget.getRows
ENd IF
rsget.Close

Dim arrDetailList
Dim MipFileGbn, isWonChonFile, MipFileState
MipFileState = 0
if (ipFileNo<>"") then

	if isarray(arrList) then
	    MipFileGbn = arrList(4,0)
	    MipFileState = arrList(3,0)
	end if

    IF (dvType="smry") then
        arrDetailList = fnGetJFixIpkumListSum(ipFileNo)
    ELSE
        arrDetailList = fnGetJFixIpkumList(ipFileNo)
    END IF
end if

isWonChonFile = (MipFileGbn="WN")

Dim ttlCnt:ttlCnt=0
Dim ttlSum:ttlSum=0
Dim ttlIpCnt:ttlIpCnt=0
Dim ttlSndCnt:ttlSndCnt=0
dim thismonth
thismonth = Left(CStr(DateSerial(year(now()),month(now())-1,1)),7)

dim ipsum, isGrp, isSubGrp, wn_ipsum

function getJGetStateColor(jstate)
    if IsNULL(jstate) then
        getJGetStateColor="#FF0000"
        exit function
    end if

    jstate=CStr(jstate)
    if jstate="0" then
		getJGetStateColor = "#000000"
	elseif jstate="1" then
	    getJGetStateColor = "#448888"
	elseif jstate="2" then
	    getJGetStateColor = "#0000FF"
	elseif jstate="3" then
		getJGetStateColor = "#0000FF"
	elseif jstate="7" then
		getJGetStateColor = "#FF0000"
	else

	end if
end function

function getJGetStateName(jstate)
    if IsNULL(jstate) then
        getJGetStateName="������"
        exit function
    end if

    jstate = CStr(jstate)
    if jstate="0" then
		getJGetStateName = "������"
	elseif jstate="1" then
	    getJGetStateName = "��üȮ�δ��"
	elseif jstate="2" then
	    getJGetStateName = "��üȮ�οϷ�"
	elseif jstate="3" then
		getJGetStateName = "����Ȯ��"
	elseif jstate="7" then
		getJGetStateName = "�ԱݿϷ�"
	else
        getJGetStateName = jstate
	end if
end function

function fnGetIpFileGbnColor(igbn)
    fnGetIpFileGbnColor = "#000000"

    if IsNULL(igbn) then Exit function

    SELECT CASE igbn
        CASE "WN" : fnGetIpFileGbnColor = "#22CCCC"
        CASE "ON" : fnGetIpFileGbnColor = "#2222CC"
        CASE "OF" : fnGetIpFileGbnColor = "#CC2222"
    END SELECT

end function

function fnGetIpkumStateName(iState)
    if IsNULL(iState) then
        fnGetIpkumStateName="������"
        exit function
    end if

    iState = CStr(iState)
    if iState="0" then
		fnGetIpkumStateName = "���"
	elseif iState="1" then

	elseif iState="2" then

	elseif iState="3" then
        fnGetIpkumStateName = "<font color=blue>��������</font>"
	elseif iState="7" then
		fnGetIpkumStateName = "<font color=red>�ԱݿϷ�</font>"
    elseif iState="8" then
		fnGetIpkumStateName = "�Ϸ�"
	else
        fnGetIpkumStateName = iState
	end if
end function

public function getJGubunName(ijgubun)
    if isNULL(ijgubun) then Exit function

    if (ijgubun="MM") then
        getJGubunName = "����"
    elseif (ijgubun="CC") then
        getJGubunName = "<font color=blue>������</font>"
    else
        getJGubunName = ijgubun
    end if
end function

%>


<script language='javascript'>
var firstSel = 0;
var secondSel = 0;
var thirdSel = 0;

var firstGroup = '';
var secondGroup= '';
var thirdGroup= '';

var firstAcct = '';
var secondAcct= '';
var thirdAcct= '';

var isthird =0;

function jsGroupSelect(idx,bank,account){
    var popwin = window.open("/admin/upchejungsan/pop_Group_select.asp?ipFileNo=<%=ipFileNo%>&idx="+idx+"&bank="+bank+"&bankaccount="+account,"PopGroupSelect","width=800 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function jsDelGroup(firstSel,grpidx){
	 if (confirm('�׷��� �����Ͻðڽ��ϱ�?')){
        var frm = document.frmSbmit;
        frm.mode.value="delGroup";
        frm.firstSel.value=firstSel;
 				frm.grpidx.value = grpidx
        frm.submit();
    }
}

function PopJungsanUpload(){
	var popwin = window.open("/admin/upchejungsan/pop_jungsan_upload.asp","PopJungsanUpload","width=800 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function delbankingup2(iidx,ipFileDIdx){
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		var popwin = window.open("/admin/upchejungsan/dobankingupflag.asp?mode=delflagWF&id=" + iidx + '&ipFileDIdx='+ ipFileDIdx,"delipkumfinish","width=100 height=100");
		popwin.focus();
	}
}

function DelIcheMaster(ipFileNo){
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		var popwin = window.open("/admin/upchejungsan/dobankingupflag.asp?mode=delmast&ipFileNo=" + ipFileNo,"DelIcheMaster","width=100 height=100");
		popwin.focus();
	}
}

function jsIpkumFinish(ipFileNo){
    if (frmList.ipkumregdate.value.length<1){
		alert('�Ա����� �Է��ϼ���.');
		frmList.ipkumregdate.focus();
		calendarOpen(frmList.ipkumregdate);
		return;
	}else{
	    frmSbmit.ipkumregdate.value=frmList.ipkumregdate.value;
	}


    if (confirm(frmSbmit.ipkumregdate.value + '�� �Ա�Ȯ�� ���� �Ͻðڽ��ϱ�?')){
        frmSbmit.mode.value="ipkumfinishWF";
        frmSbmit.submit();
    }
}

function upFilexLDown(ipFileNo,xltype){

    ifraXL.location.href="jungsan_file_xls.asp?ipFileNo="+ipFileNo+"&xltype="+xltype;
}

//
function makePayreqList(ipFileNo){
    if (confirm('������û Data�� �ۼ��Ͻðڽ��ϱ�?')){
        frmSbmit.mode.value="makeItemBuyingErpData";
        frmSbmit.ipFileNo.value = ipFileNo;
        frmSbmit.submit();
    }
}

function popSendERP(ipFileNo){
    var ipopURI = "popSendERP.asp?ipFileNo="+ipFileNo;
    var popWin = window.open(ipopURI,'popSendERP','width=800,height=700,scrollbars=yes,resizable=yes');
    popWin.focus();
}

function popSendIcheFileERP(ipFileNo){
    if (confirm('��ü ������ ERP�� ���� �Ͻðڽ��ϱ�?')){
        document.frmErp.LTp.value="AF";
        document.frmErp.ipFileNo.value=ipFileNo;
        document.frmErp.submit();
    }
}

function jsPopErpReceiveOrCustMap(igroupid){
    alert('���� ���� ��� '+igroupid);
    // ToDo
    // http://webadmin.10x10.co.kr/admin/linkedERP/cust/popGetCust.asp?opnType=&rdoCgbn=0&selCT=&selSTp=1&sSTx=G05241 ���� ��ϼ���
    // or
    // db_scm_link.dbo.[sp_BA_CUST_Update]
    // or
    // ��ü ���� â����  ERP�����ڵ� ���� (���� �ٸ� ����ڷ� ��ϵ� �����)
}
function deleteFileNo(iFileno){
    if (confirm('Ȯ�� Ŭ���� FileNo�� ���� ���������� �����˴ϴ�\n\n�����Ͻðڽ��ϱ�?')){
		document.frmSbmit.target = "ifraXL";
		document.frmSbmit.ipFileNo.value = iFileno;
		document.frmSbmit.mode.value = "deleteFileNo";
		document.frmSbmit.submit();
	}
}
</script>


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
            ���ϻ��� :
            <select name="ipFileState">
            <option value="" >��ü
            <option value="0" <%= CHKIIF(ipFileState="0","selected","") %> >�ۼ���
            <option value="3" <%= CHKIIF(ipFileState="3","selected","") %> >ERP����
            <option value="7" <%= CHKIIF(ipFileState="7","selected","") %> >�ԱݿϷ�
            </select>
            &nbsp;
            ���Ϲ�ȣ : <input type="text" name="ipFileNo" value="<%= ipFileNo %>" size="4" maxlength="7">
            &nbsp;
            <input type="button" value="��ü����" class="button" onClick="document.frm.ipFileNo.value='';document.frm.submit();">
            &nbsp;
            ���� ���� :
            <select name="targetGbn">
            <option value="" >��ü
            <option value="ON" <%= CHKIIF(targetGbn="ON","selected","") %> >�¶���
            <option value="OF" <%= CHKIIF(targetGbn="OF","selected","") %> >��������
            </select>
            <!--
        	<input type="button" value="������ε�����" onclick="PopJungsanUpload();">
        	-->
        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- ����Ʈ -->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" >���� ���</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">FileNo</td>
		<td width="40">���ϱ���</td>
		<td width="50">���걸��</td>
		<td width="70">������</td>
		<td width="120">���ϸ�</td>
		<td width="120">�ۼ���</td>
      	<td width="100">����</td>
      	<td width="90">�������</td>
      	<td width="70">�ѰǼ�</td>
      	<td width="70">�ԱݿϷ�<br>�Ǽ�</td>
      	<td width="70">erp����<br>�Ǽ�</td>
      	<td >��ü���� ����</td>
		<td >��꼭 ����</td>
		<td width="40">���</td>
	</tr>
	<% IF isArray(arrList) THEN %>
	<%  For intLoop = 0 To UBound(arrList,2) %>
	    <%

         ttlSum = ttlSum + arrList(7,intLoop)+arrList(8,intLoop)
         ttlCnt = ttlCnt + arrList(5,intLoop)
         ttlIpCnt = ttlIpCnt + arrList(6,intLoop)
         ttlSndCnt = ttlSndCnt + arrList(9,intLoop)
	    %>
	    <tr align="center" bgcolor="#FFFFFF">
		<td><a href="?ipFileNo=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>"><%=arrList(0,intLoop)%></a></td>
		<td><a href="?ipFileNo=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>"><font color="<%= fnGetIpFileGbnColor(arrList(4,intLoop)) %>"><%=arrList(4,intLoop)%></font></a></td>
		<td><%=getJGubunName(arrList(10,intLoop))%></td>
		<td><%=arrList(11,intLoop)%></td>
		<td><a href="?ipFileNo=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>"><%=arrList(1,intLoop)%></a></td>
		<td><%=arrList(2,intLoop)%></td>
		<td><%=fnGetIpkumStateName(arrList(3,intLoop))%></td>
		<td align="right"><%=FormatNumber(arrList(7,intLoop)+arrList(8,intLoop),0)%></td>
		<td><%=arrList(5,intLoop)%></td>
		<td><%=arrList(6,intLoop)%></td>
		<td><%= FormatNumber(arrList(9,intLoop),0) %></td>
		<td>
		    <% if  arrList(3,intLoop)<3 then %>
	    <% if arrList(4,intLoop)="OF" or arrList(4,intLoop)="ON" then %>
		    ���� ���� <img src="/images/icon_arrow_link.gif" onclick="popSendIcheFileERP('<%= arrList(0,intLoop) %>');" style="cursor:pointer">
		    <% end if %>
		    <% end if %>

		</td>
		<td>
		   <!--
		   <% if arrList(3,intLoop)>=3 then %>
		   <% if arrList(4,intLoop)="OF" or arrList(4,intLoop)="ON" then %>
		   ERP���� <img src="/images/icon_arrow_link.gif" onclick="popSendERP('<%= arrList(0,intLoop) %>');" style="cursor:pointer">
		   <% end if %>
		   <% end if %>

		   <% if arrList(3,intLoop)>=0 and arrList(5,intLoop)=0 then %>
           <a href="javascript:DelIcheMaster('<%=arrList(0,intLoop)%>')"><img src="/images/i_delete.gif" width="8" height="9" border="0"></a>
           <% end if %>
           -->
		</td>
		<td>
		<% If arrList(3,intLoop) = 0 Then %>
			<input type="button" class="button" value="����" onclick="deleteFileNo('<%=arrList(0,intLoop)%>');" />
		<% End If %>
		</td>
    <%  next %>
        <% if UBound(arrList,2)>0 then %>
        <tr align="center" bgcolor="#EEEEEE">
        <td>�հ�</td>
        <td colspan="6"></td>
        <td align="right"><%= FormatNumber(ttlSum,0) %></td>
        <td><%= FormatNumber(ttlCnt,0) %></td>
        <td><%= FormatNumber(ttlIpCnt,0) %></td>
        <td><%= FormatNumber(ttlSndCnt,0) %></td>
        <td></td>
        <td></td>
		<td></td>
        </tr>
        <% End IF  %>
    <% End IF  %>

</table>
<br><br>
<% IF isArray(arrDetailList) THEN %>
<form name="frmList">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="16" >���� �� <%= UBound(arrDetailList,2)+1 %>

    	<input type="radio" name="dvType" value="" <%= CHKIIF(dvType="","checked","") %> onClick="location.href='?ipFileNo=<%=ipFileNo%>&dvType=&menupos=<%=menupos%>'">����Ʈ
    	<input type="radio" name="dvType" value="smry" <%= CHKIIF(dvType="smry","checked","") %> onClick="location.href='?ipFileNo=<%=ipFileNo%>&dvType=smry&menupos=<%=menupos%>'">�հ�(���ε�����)
    	&nbsp;&nbsp;&nbsp;&nbsp;
    	��ü���� <img src="/images/iexcel.gif" onclick="upFilexLDown('<%= ipFileNo %>',1);" style="cursor:pointer">
    	&nbsp;&nbsp;&nbsp;&nbsp;
    	��������ȸ <img src="/images/iexcel.gif" onclick="upFilexLDown('<%= ipFileNo %>',2);" style="cursor:pointer">
    	&nbsp;&nbsp;&nbsp;&nbsp;
    	<% if isarray(arrList) then %>
	    	<% if arrList(3,0)<=3 then %>

	    	�Ա��� : <input type=text name=ipkumregdate value="" size=10 maxlength=10 readonly >
			�Ա�Ȯ�� ����<img src="/images/icon_arrow_link.gif" onclick="jsIpkumFinish(<%= ipFileNo %>);" style="cursor:pointer">
			<% end if %>
		<% end if %>
    	</td>
    </tr>
    <% IF (dvType="smry") then %>
    <!-- ����ó(�ŷ�ó�ڵ�), �Ա�����, �Աݰ���, ��ü�ݾ�, ��������μ⳻��(�ŷ�ó�� ��ũ���� (5))==�����ָ�x,�Ա������μ⳻��((��)�ٹ�����)-->
        <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    		<td width="80">�ŷ�ó��</td>
    		<td width="80">�Ա�����</td>
    		<td width="120">�Աݰ���</td>
    		<% if (isWonChonFile) then %>
    		<td width="100">(��õ)��ü�ݾ�</td>
    		<% else %>
    		<td width="100">��ü�ݾ�</td>
    		<% end if %>
          	<td width="150">��������μ⳻��</td>
    		<td width="100">�Ա������μ⳻��</td>
    		<td width="100">������</td>
    	</tr>
    	<%  For intLoop = 0 To UBound(arrDetailList,2) %>
    	<%
    	ipsum = ipsum + arrDetailList(3,intLoop)
    	wn_ipsum = wn_ipsum + GetHoldingJungSanSum(arrDetailList(3,intLoop))
    	%>
    	<tr align="center" bgcolor="#FFFFFF">
    		<td><%= arrDetailList(5,intLoop) %></td>
    		<td><%= arrDetailList(1,intLoop) %></td>
    		<td><%= arrDetailList(2,intLoop) %></td>
    		<% if (isWonChonFile) then %>
    		<td><%= GetHoldingJungSanSum(arrDetailList(3,intLoop)) %></td>
    		<% else %>
    		<td><%= arrDetailList(3,intLoop) %> (<%= CLNG(arrDetailList(3,intLoop)) %>)</td>
    		<% end if %>
    		<td><%= arrDetailList(5,intLoop) %></td>
    		<td>(��)�ٹ�����</td>
    		<td><%= arrDetailList(6,intLoop) %></td>
    	</tr>
    	<%  next %>
    	<tr bgcolor="#FFFFFF">
    		<td colspan="3"></td>
    		<% if (isWonChonFile) then %>
    		<td align="right"><%= FormatNumber(wn_ipsum,0) %></td>
    		<% else %>
    		<td align="right"><%= FormatNumber(ipsum,0) %></td>
    		<% end if %>
    		<td colspan="3"></td>
    	</tr>
    <% else %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">�����</td>
		<td width="60">����</td>
		<td width="60">���걸��</td>
		<td width="70">������</td>
		<td width="40">������</td>
		<td width="120">�귣��ID</td>
      	<td width="150">������ (���׷�)</td>
		<td width="60">����</td>
		<td width="60">����</td>
		<td width="80">����</td>
		<% if (isWonChonFile) then %>
		<td width="80">Ȯ���ݾ�</td>
		<td width="80">(��õ)����ݾ�</td>
		<% else %>
		<td width="80">����ݾ�</td>
		<% end if %>
		<td>��ü��</td>
		<td width="50">�׷��ڵ�</td>
		<td width="50">Erp�ڵ�</td>
		<td width="30">����</td>
		<td width="30">GRP</td>
	</tr>
	<%  For intLoop = 0 To UBound(arrDetailList,2) %>
	<%
	ipsum = ipsum + arrDetailList(4,intLoop)
	wn_ipsum = wn_ipsum + GetHoldingJungSanSum(arrDetailList(4,intLoop))  ''��õ¡����
	isGrp = FALSE
	isSubGrp  = FALSE
	isGrp = (arrDetailList(13,intLoop)>0)
	isSubGrp = Not ISNULL(arrDetailList(14,intLoop))
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= arrDetailList(6,intLoop) %></td>
		<td><%= arrDetailList(1,intLoop) %></td>
		<td><%= getJGubunName(arrDetailList(20,intLoop)) %></td>
		<td>
			<% if Left(arrDetailList(5,intLoop),7) = Left(CStr(now()),7) then %>
			<font color="red"><%= arrDetailList(5,intLoop) %></font>
			<% else %>
			<font color="blue"><%= arrDetailList(5,intLoop) %></font>
			<% end if %>
		</td>
		<td><%= arrDetailList(11,intLoop) %></td>
		<td><a href="javascript:PopUpcheBrandInfoEdit('<%= arrDetailList(8,intLoop) %>')"><%= arrDetailList(8,intLoop) %></a></td>
		<td><%= arrDetailList(15,intLoop) %></td>
		<td><font color="<%= getJGetStateColor(arrDetailList(12,intLoop)) %>"><%= getJGetStateName(arrDetailList(12,intLoop)) %></font></td>
		<td><%= arrDetailList(9,intLoop) %></td>
		<td><%= arrDetailList(10,intLoop) %></td>
		<% if (isWonChonFile) then %>
		<td align="right">
		<% if arrDetailList(4,intLoop)<1 then %><font color=red><% end if %>
		<%= FormatNumber(arrDetailList(4,intLoop),0) %></font></td>
		<td><%= FormatNumber(GetHoldingJungSanSum(arrDetailList(4,intLoop)),0) %></td>
		<% else %>
		<td align="right">
		<% if arrDetailList(4,intLoop)<1 then %><font color=red><% else %><font color="#000000"><% end if %>
		<% if Not isNULL(arrDetailList(4,intLoop)) then %><%= FormatNumber(arrDetailList(4,intLoop),0) %><% end if %>
		</font>
		</td>
		<% end if %>
		<td><%= arrDetailList(16,intLoop) %></td>
		<td><%= arrDetailList(7,intLoop) %></td>
		<td <%=CHKIIF(arrDetailList(18,intLoop)=0 or isNULL(arrDetailList(19,intLoop)),"bgcolor='#CCCCCC'","") %> >
		    <% if IsNULL(arrDetailList(19,intLoop)) then %>
		    <img src="/images/icon_arrow_link.gif" onclick="jsPopErpReceiveOrCustMap('<%= arrDetailList(7,intLoop) %>');" style="cursor:pointer">
		    <% else %>
		    <%= arrDetailList(19,intLoop) %>
		    <% end if %>
		</td>
		<td>
		<a href="javascript:delbankingup2('<%= arrDetailList(2,intLoop) %>','<%= arrDetailList(0,intLoop) %>')">
		    <% if (MipFileState<=3) then ''3���� ���� �����ϰ� %>
    		<a href="javascript:delbankingup2('<%= arrDetailList(2,intLoop) %>','<%= arrDetailList(0,intLoop) %>')">

    		<% IF (isGrp or isSubGrp) THEN %>
    		<!-- �׷����� -->
    		<% else %>
    		x
    		<% end if %>
    		</a>
    		<% else %>
    		    <% if (FALSE) then %>
    	        <% if arrDetailList(12,intLoop)<7 then %>
    	        <a href="javascript:delbankingup2('<%= arrDetailList(2,intLoop) %>','<%= arrDetailList(0,intLoop) %>')">x</a>
    	        <% end if %>
    	        <% end if %>

				<% if C_ADMIN_AUTH or C_MngPowerUser then %>
					x[������]
				<% end if %>
			<% end if %>
		</a>
		</td>
		<td>
		    <% IF (isGrp or isSubGrp) THEN %>
		    <%= CHKIIF(ISNULL(arrDetailList(14,intLoop)),arrDetailList(0,intLoop),arrDetailList(14,intLoop)) %>
		    <%if not isNull(arrDetailList(14,intLoop)) then%><span style="padding:3px;"><a href="javascript:jsDelGroup('<%= arrDetailList(0,intLoop) %>','<%=arrDetailList(14,intLoop)%>');">[x]</a></span><%end if%>
		    <% else %>
		        <% if arrDetailList(4,intLoop)<1 then %>
		            <img src="/images/icon_arrow_link.gif" onclick="jsGroupSelect('<%= arrDetailList(0,intLoop) %>','<%= arrDetailList(9,intLoop) %>','<%= arrDetailList(10,intLoop) %>');" style="cursor:pointer">
		        <% end if %>
		    <% end if %>
		</td>
	</tr>
	<%  next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9"></td>
		<% if (isWonChonFile) then %>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td align="right"><%= FormatNumber(wn_ipsum,0) %></td>
		<% else %>
		<% if isNULL(ipsum) then %>
		<td align="right">NULL</td>
		<% else %>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<% end if %>
		<% end if %>
		<td colspan="6"></td>
	</tr>
	<% end if %>
</table>
</form>
<% end if %>


<p>

<form name="frmSbmit" method="post" action="dobankingupflag.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="ipkumregdate" value="">
<input type="hidden" name="ipFileNo" value="<%= ipFileNo %>">
<input type="hidden" name="firstSel" value="">
<input type="hidden" name="secondSel" value="">
<input type="hidden" name="thirdSel" value="">
<input type="hidden" name="grpidx" value="">
</form>

<form name="frmErp" method="post" action="/admin/approval/payReqList/S_erpLink_Process.asp">
<input type="hidden" name="LTp" value="">
<input type="hidden" name="ipFileNo" value="<%= ipFileNo %>">
</form>

<iframe name="ifraXL" id="ifraXL" src="" width="1" height="1" frameborder="0" scrolling="no"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->