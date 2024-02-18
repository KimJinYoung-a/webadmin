<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/incLotteiMallFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'response.write "������"
'response.end

'''/// /admin/apps/outMallAutoJob.asp ���� �Լ� ���� ���ü������
function N_TenDlvCode2CommonDlvCode(imallname,itenCode)
    if (LCASE(imallname)="lottecom") or (LCASE(imallname)="lottecomm") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2LotteDlvCode(itenCode)
    elseif (LCASE(imallname)="lotteimall") then
		If (Now() > #09/01/2015 00:00:00#) Then
			N_TenDlvCode2CommonDlvCode = TenDlvCode2LotteiMallNewDlvCode(itenCode)
		Else
			N_TenDlvCode2CommonDlvCode = TenDlvCode2LotteiMallDlvCode(itenCode)
		End If
    elseif (LCASE(imallname)="interpark") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2InterParkDlvCode(itenCode)
    elseif (LCASE(imallname)="cjmall") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2cjMallDlvCode(itenCode)
    elseif (LCASE(imallname)="gseshop") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2GSShopDlvCode(itenCode)
    elseif (LCASE(imallname)="homeplus") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2HomeplusDlvCode(itenCode)
    elseif (LCASE(imallname)="ezwel") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2EzwelDlvCode(itenCode)
    elseif (LCASE(imallname)="auction1010") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2AuctionDlvCode(itenCode)
    end if
end function



dim sitename, research
dim matchState
Dim siteType, searchType
sitename = RequestCheckVar(Request("sitename"),32)
research = RequestCheckVar(Request("research"),32)
siteType = RequestCheckVar(Request("siteType"),32)
searchType = RequestCheckVar(Request("searchType"),32)

if (searchType="") then searchType="sendSongjang"

Dim sqlStr
Dim ArrList
CONST MAXROWS = 700
sqlStr = "select top "&MAXROWS&" T.orderserial, T.OutMallOrderSerial,T.matchitemid,T.matchitemoption "
sqlStr = sqlStr & " ,T.OrgDetailKey, IsNULL(T.sendState,0) as sendState"
sqlStr = sqlStr & " ,D.itemname,D.itemOptionName"
sqlStr = sqlStr & " ,D.itemNo, D.songjangDiv, D.songjangNo, D.cancelyn, M.cancelyn,M.ipkumdiv"
sqlStr = sqlStr & " ,V.divname, D.beasongdate, D.isUpchebeasong, T.sendReqCnt, T.sellsite, T.outMallGoodsNo"
sqlStr = sqlStr & " ,D.idx, IsNull(T.recvSendReqCnt, 0) as recvSendReqCnt, T.receiveName "
sqlStr = sqlStr & " from db_temp.dbo.tbl_xsite_mayDelOrder  T"
sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master M"
sqlStr = sqlStr & " 	on T.orderserial=M.orderserial"
sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D"
sqlStr = sqlStr & " 	on T.orderserial=D.orderserial"
sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid"					'// ���� �ֹ��� ������ ���(����1��,�Ķ�1�� -> �Ķ�2��)
sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption"
sqlStr = sqlStr & " 	and D.currstate=7"
sqlStr = sqlStr & " 	left join db_order.dbo.tbl_songjang_div V"
sqlStr = sqlStr & " 	on D.songjangDiv=V.divcd"
sqlStr = sqlStr & " where 1=1 and datediff(m,T.regdate,getdate())<7"						'// �ֱ� 6����
sqlStr = sqlStr & " and T.OrgDetailKey is Not NULL"
sqlStr = sqlStr & " and T.matchState not in ('R','D','B')"              ''��ȯ ��� ����.

if (sitename<>"") then
    sqlStr = sqlStr & " and T.sellsite='"&sitename&"'"
end if

if (siteType="") then
    sqlStr = sqlStr & " and T.sellsite in ('lotteCom','lotteimall','interpark','cjmall','lotteComM', 'gseshop', 'homeplus', 'ezwel', 'auction1010', 'gmarket1010', 'nvstorefarm', '11st1010','ssg', 'halfclub', 'gsisuper', 'coupang') "
    sqlStr = sqlStr & " and not (T.sellsite = 'gsisuper' and T.shoplinkerorderid is NULL) "
end if

if (searchType = "sendSongjang") then
	sqlStr = sqlStr & " and IsNULL(T.sendState,0)=0"
elseif (searchType = "sendRecvState") then
	sqlStr = sqlStr & " and ( "
	sqlStr = sqlStr & " 	((T.sellsite = 'cjmall') and (d.songjangDiv not in ('4', '3', '28', '2', '1', '13'))) "		'// ����Ʈ��ŷ(CJ���� �ٷ� ���μ� Ȯ��)
	sqlStr = sqlStr & " ) "
	sqlStr = sqlStr & " and IsNULL(T.sendState, 0) <> 0 "
	sqlStr = sqlStr & " and IsNULL(T.recvSendState, 0) < 100 "
	sqlStr = sqlStr & " and DateDiff(d, d.beasongdate, getdate()) >= 1 "
else
	'// ����
end if

sqlStr = sqlStr & " order by D.beasongdate"
''sqlStr = sqlStr & " order by T.OutMallOrderseq"


''rw sqlStr
''response.end

''IF (sitename<>"") then
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        ArrList = rsget.getRows
    end if
    rsget.Close
''ENd IF

dim i
%>


<script language='javascript'>
function sendSongJangCJMALL(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        alert('���޻� �ù���ڵ� ������');
        return;
    }

    if ((OutMallSDiv=="99")&&(songjangNo=="")){
        songjangNo="��Ÿ"
    }

    if (songjangNo==""){
        alert('�����ȣ ������..');
        return;
    }


    //proc_gubun=sfin:�߼ۿϷ� //dfin:��ۿϷ�

    var params = "ten_ord_no="+tenorderserial+"&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
	var popwin=window.open('/admin/etc/cjmall/actCJmallSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
//	var popwin=window.open('http://wapi.10x10.co.kr/outmall/cjmall/actCJmallSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangGSShop(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        alert('���޻� �ù���ڵ� ������');
        return;
    }

    if ((OutMallSDiv=="99")&&(songjangNo=="")){
        songjangNo="��Ÿ"
    }

    if (songjangNo==""){
        alert('�����ȣ ������..');
        return;
    }

    var params = "ten_ord_no="+tenorderserial+"&ordclmNo="+OutMallOrderSerial+"&ordSeq="+OrgDetailKey+"&delvEntrNo="+OutMallSDiv+"&invoNo="+songjangNo;
    var popwin=window.open('/admin/etc/gsshop/actGSShopSongjangInputProc.asp?' + params,'sendSongJangGSShop','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendRecvStateCJMALL(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName, errTakBae){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "ten_ord_no="+tenorderserial+"&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&rcv_nm="+receiveName;
	var popwin;

	if(errTakBae == '2' ){
		popwin=window.open("/admin/etc/cjmall/actCjmallRecvStateInputProc.asp?" + params,"sendRecvStateCJMALL","width=600,height=400,scrollbars=yes,resizable=yes");
	    popwin.focus();
	    comp.disabled=true;
	}else{
		if (confirm('�浿�ù� Ȩ���������� �μ�Ȯ�� �ϼ̳���?')){
			popwin=window.open("/admin/etc/cjmall/actCjmallRecvStateInputProc.asp?" + params,"sendRecvStateCJMALL","width=600,height=400,scrollbars=yes,resizable=yes");
		    popwin.focus();
		    comp.disabled=true;
		}
	}
//	var popwin=window.open("http://wapi.10x10.co.kr/outmall/cjmall/actCjmallRecvStateInputProc.asp?" + params,"sendRecvStateCJMALL","width=600,height=400,scrollbars=yes,resizable=yes");
}

function sendSongJangHomeplus(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
//alert('Wapi�� pop')
	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo;
    //var popwin=window.open("/admin/etc/homeplus/actHomeplusSongjangInputProc.asp?" + params,"sendRecvStateHomeplus","width=600,height=400,scrollbars=yes,resizable=yes");
    var popwin=window.open("<%=apiURL%>/outmall/proc/Homeplus_SongjangProc.asp?" + params,"sendRecvStateHomeplus","width=600,height=400,scrollbars=yes,resizable=yes"); //��ȭ��Ȯ���ʿ�
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangEzwel(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }

	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo;
    var popwin=window.open("/admin/etc/ezwel/actEzwelSongjangInputProc.asp?" + params,"sendRecvStateEzwel","width=600,height=400,scrollbars=yes,resizable=yes");
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangAuction(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }

	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&songjangDiv="+receiveName;
    var popwin=window.open('<%=apiURL%>/outmall/proc/Auction_DelSongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}


function sendSongJang(comp,OutMallOrderSerial,OrgDetailKey,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        alert('���޻� �ù���ڵ� ������');
        return;
    }

    if ((OutMallSDiv=="99")&&(songjangNo=="")){
        songjangNo="��Ÿ"
    }

    if (songjangNo==""){
        alert('�����ȣ ������..');
        return;
    }


    //proc_gubun=sfin:�߼ۿϷ� //dfin:��ۿϷ�
//alert('Wapi�� pop')
    var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
    //var popwin=window.open('/admin/etc/lotte/actLotteSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    var popwin=window.open('<%=apiURL%>/outmall/proc/LotteCom_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');


    popwin.focus();
    comp.disabled=true;
}

function sendSongJangiMall(comp,OutMallOrderSerial,OrgDetailKey,sendQnt,sendDate,outmallGoodsID,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        alert('���޻� �ù���ڵ� ������');
        return;
    }

    if (songjangNo==""){
        alert('�����ȣ ������');
        return;
    }

    //proc_gubun=sfin:�߼ۿϷ� //dfin:��ۿϷ�

    var params = "cmdparam=songjangip&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&sendQnt="+sendQnt+"&sendDate="+sendDate+"&outmallGoodsID="+outmallGoodsID+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
    var popwin=window.open('/admin/etc/lotteimall/actLotteiMallReq.asp?' + params,'xSiteSongjangInputProciMall','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

// �ű� �ý���
function sendSongJangiMallNew(comp,OutMallOrderSerial,OrgDetailKey,sendQnt,sendDate,outmallGoodsID,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������");
        return;
    }

    var params = "mode=sendsongjang&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&sendQnt="+sendQnt+"&sendDate="+sendDate+"&outmallGoodsID="+outmallGoodsID+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
	var popwin=window.open("/admin/etc/orderInput/xSiteCSOrder_lotteimall_Process.asp?" + params,"sendSongJangiMallNew","width=600,height=400,scrollbars=yes,resizable=yes");
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangInterpark(comp,OutMallOrderSerial,OrgDetailKey,yyyymmdd,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        //alert('���޻� �ù���ڵ� ������');
        //return;

        OutMallSDiv="169167";
    }

    if ((OutMallSDiv=="169167")&&(songjangNo=="")){
        songjangNo="��Ÿ"
    }

    if (songjangNo==""){
        alert('�����ȣ ������..');
        return;
    }



    var params = "ordclmNo="+OutMallOrderSerial+"&ordSeq="+OrgDetailKey;
    params=params+"&delvDt="+yyyymmdd+"&delvEntrNo="+OutMallSDiv+"&invoNo="+songjangNo;
    params=params+"&optPrdTp=01&optOrdSeqList="+OrgDetailKey
    //alert(params)
    var popwin=window.open('/admin/etc/interparkXML/actInterparkSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<!-- �˻� ���� -->
<% If session("ssBctID")="kjy8517" Then %>
UPDATE R
SET R.orderserial = B.orderserial
,matchState = 'O'
FROM db_temp.dbo.tbl_xsite_mayDelOrder as R
JOIN db_temp.dbo.tbl_xsite_tmpOrder as B on R.OutMallOrderSerial = B.OutMallOrderSerial
WHERE isnull(R.sendReqCnt, 0) = 0
<% End If %>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">

	<tr align="center">
		<td width="50" bgcolor="<%= adminColor("gray") %>" class="td_br">�˻�<br>����</td>
		<td align="left" class="td_br">
		    ���θ� ���� :
            <% CALL DrawApiMallSelect("sitename",sitename) %>


            <!--
		    &nbsp;&nbsp;
		    ó������ :
			<select class="select" name="matchState">
			<option value='' <%= chkIIF(matchState="","selected","") %> >��ü</option>
	     	<option value='I' <%= chkIIF(matchState="I","selected","") %> >�������</option>
	     	<option value='O' <%= chkIIF(matchState="O","selected","") %> >�ֹ��Է¿Ϸ�</option>
	     	<option value='D' <%= chkIIF(matchState="D","selected","") %> >���Է»���</option>
	     	</select>
	     	&nbsp;
            -->

            &nbsp;&nbsp;
			�� 30�и��� �ڵ� ó���� (15, 45)
		    &nbsp;&nbsp;
		    �˻����� :
			<select class="select" name="searchType">
	     	<option value="sendSongjang" <%= chkIIF(searchType="sendSongjang","selected","") %> >��������</option>
	     	<option value="sendRecvState" <%= chkIIF(searchType="sendRecvState","selected","") %> >���μ�����</option>
	     	</select>
		</td>

		<td width="50" bgcolor="<%= adminColor("gray") %>" class="td_br">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr height="25">
		<td align="left">
			<% if (IsArray(ArrList)) THEN %>
            ���� �˻��� : <%= UBound(ArrList,2)+1 %>�� (MAX <%= MAXROWS %> )
            <% else %>
            ���� �˻��� : 0 ��
            <% end if %>
		<!--
			<input type="button" class="button" value="���ó�����������" onClick="SubmitSongjangInput(frmSvArr)">
		-->
		</td>
	</tr>
</table>
<!-- �׼� �� -->
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl" >
<form name="frmSvArr" method="post" action="OrderInput_Process.asp">
	<input type="hidden" name="mode" value="add">
	<tr align="center" class="tr_tablebar">
	<!--
	    <td width="20" class="td_br"><input type="checkbox" name="chkAll" onclick="fnCheckValidAll(this.checked,frmSvArr.cksel);"></td>
	 -->
	    <td width="70" class="td_br">���޻�</td>
	    <td width="70" class="td_br">�ֹ���ȣ</td>
    	<td width="100" class="td_br">�����ֹ���ȣ</td>
      	<td width="60" class="td_br">����<br>��ǰ�ڵ�</td>
      	<td width="50" class="td_br">��ǰ�ڵ�</td>
      	<td width="50" class="td_br">�ɼ��ڵ�</td>
      	<td  class="td_br">��ǰ�� <font color="blue">[�ɼǸ�]</font></td>
        <td width="30" class="td_br">����</td>
        <td width="100" class="td_br">�����</td>
        <td width="40" class="td_br">���<br>����</td>
      	<td class="td_br">�ù��</td>
      	<td class="td_br">����</td>
      	<td class="td_br">���޻�<br>�ù��ڵ�</td>
      	<td class="td_br">����</td>
      	<td class="td_br">����<br>Ƚ��</td>
    </tr>
<% if (IsArray(ArrList)) THEN %>
<%
Dim intRows : intRows = UBound(ArrList,2)
dim iOutMallDlvCode
for i=0 to intRows
    iOutMallDlvCode = ""
    iOutMallDlvCode = N_TenDlvCode2CommonDlvCode(ArrList(18,i),ArrList(9,i))
%>
<tr>
    <!--<td class="td_br"><input type="checkbox" name="cksel" value=""></td> -->
    <td class="td_br"><%= ArrList(18,i) %></td>
    <td class="td_br"><%= ArrList(0,i) %></td>
    <td class="td_br"><%= ArrList(1,i) %></td>
    <td class="td_br"><%= ArrList(4,i) %></td>
    <td class="td_br"><%= ArrList(2,i) %></td>
    <td class="td_br"><%= ArrList(3,i) %></td>
    <td class="td_br"><%= ArrList(6,i) %>
    <% if ArrList(7,i)<>"" then %>
    <font color="blue">[<%= ArrList(7,i) %>]</font>
    <% end if %>
    </td>
    <td class="td_br" align="center"><%= ArrList(8,i) %></td>
    <td class="td_br"><%= ArrList(15,i) %></td>
    <td class="td_br" align="center">
        <% IF ArrList(16,i)="Y" THEN %>
        <font color="blue">��ü</font>
        <% End IF %>
    </td>
    <td class="td_br"><%= ArrList(14,i) %></td>
    <td class="td_br" <%=CHKIIF(isNULL(ArrList(10,i)),"bgcolor='#CC2222'","") %> ><%= ArrList(10,i) %></td>
    <td class="td_br" <%=CHKIIF(iOutMallDlvCode="","bgcolor='#CC2222'","")%> ><%= iOutMallDlvCode %>(<%=ArrList(9,i)%>)</td>
    <td class="td_br">
    <% if (ArrList(5,i)=0) Then %>
		<!-- �������� -->
		<%
		Select Case ArrList(18,i)
			Case "lotteCom", "lotteComM"
		%>
				<input type="button" value="����" onClick="sendSongJang(this,'<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
			Case "lotteimall"
				if InStr(ArrList(4,i),"-")>0 then
		%>
					<input type="button" value="����OLD" onClick="sendSongJangiMall(this,'<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= ArrList(8,i) %>','<%= replace(Left(ArrList(15,i),10),"-","") %>','<%= ArrList(19,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
				else
		%>
					<input type="button" value="����" onClick="sendSongJangiMallNew(this,'<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= ArrList(8,i) %>','<%= replace(Left(ArrList(15,i),10),"-","") %>','<%= ArrList(19,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
				end if
			Case "interpark"
				%>
				<input type="button" value="����" onClick="sendSongJangInterpark(this,'<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= replace(Left(ArrList(15,i),10),"-","") %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
			Case "cjmall"
			ArrList(10,i) = Replace(ArrList(10,i), "&nbsp;", Chr(32))
			ArrList(10,i) = trim(ArrList(10,i))
		%>
				<input type="button" value="����" onClick="sendSongJangCJMALL(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
			Case "gseshop"
		%>
				<input type="button" value="����" onClick="sendSongJangGSShop(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
			Case "homeplus"
		%>
				<input type="button" value="����" onClick="sendSongJangHomeplus(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
			Case "ezwel"
		%>
				<input type="button" value="����" onClick="sendSongJangEzwel(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
			Case "auction1010"
		%>
				<input type="button" value="����" onClick="sendSongJangAuction(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>', '<%=ArrList(9,i)%>');">
		<%
			Case Else
				response.write "ERR"
		End Select
		%>
	<% elseif ((ArrList(5,i) <> 0) and (searchType = "sendRecvState")) then %>
		<!-- ���μ� Ȯ�� �� ���� -->
        <%
        	if ArrList(18,i)="cjmall" then
        		If ArrList(9,i) = "21" Then		'�浿�ù�//Ȩ������ ��û����..�ٸ� ������� �ؾ� �� ��..2015-02-27 19:04 ������
		%>
		<input type="button" class="button_s" value="�浿Ȯ��" onClick="sendRecvStateCJMALL(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= ArrList(9,i) %>','<%= ArrList(10,i) %>','<%= ArrList(22,i) %>', '1');">
		<%
        		Else
        %>
        <input type="button" value="�μ�����" onClick="sendRecvStateCJMALL(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= ArrList(9,i) %>','<%= ArrList(10,i) %>','<%= ArrList(22,i) %>', '2');">
        <%
        		End If
        	end if
        %>
	<% end if %>
    </td>
    <td class="td_br">
		<% if (searchType = "sendSongjang") then %>
			<!-- �������� -->
			<% if ArrList(17,i)>2 then %>
			<b><%= ArrList(17,i) %></b>
			<% else %>
			<%= ArrList(17,i) %>
			<% end if %>
		<% else %>
			<!-- ���μ� Ȯ�� �� ���� -->
			<% if ArrList(21,i)>2 then %>
			<b><%= ArrList(21,i) %></b>
			<% else %>
			<%= ArrList(21,i) %>
			<% end if %>
		<% end if %>
    </td>
</tr>
<% next %>
<% ELSE %>
<tr>
    <td colspan="15" align="center">[�˻� ����� �����ϴ�.]</td>
</tr>
<% end if %>
</table>
<% if (sitename="") then %>
<script>//alert('���θ��� ���� �ϼ���.');</script>
<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
