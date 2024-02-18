<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
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
    elseif (LCASE(imallname)="lotteon") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2LotteonDlvCode(itenCode)
    elseif (LCASE(imallname)="shintvshopping") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2ShintvshoppingDlvCode(itenCode)
    elseif (LCASE(imallname)="skstoa") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2SkstoaDlvCode(itenCode)
    elseif (LCASE(imallname)="wetoo1300k") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2Wetoo1300kDlvCode(itenCode)
    elseif (LCASE(imallname)="cjmall") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2cjMallDlvCode(itenCode)
    elseif (LCASE(imallname)="gseshop") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2GSShopDlvCode(itenCode)
    elseif (LCASE(imallname)="homeplus") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2HomeplusDlvCode(itenCode)
    elseif (LCASE(imallname)="ezwel") OR (LCASE(imallname)="kakaostore") OR (LCASE(imallname)="boribori1010") OR (LCASE(imallname)="benepia1010") OR (LCASE(imallname)="wconcept1010") then
        N_TenDlvCode2CommonDlvCode = itenCode
    elseif (LCASE(imallname)="auction1010") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2AuctionDlvCode(itenCode)
    elseif (LCASE(imallname)="gmarket1010") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2GmarketDlvCode(itenCode)
    elseif (LCASE(imallname)="nvstorefarm") or (LCASE(imallname)="nvstoremoonbangu") or (LCASE(imallname)="mylittlewhoopee") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2NvstorefarmDlvCode(itenCode)
    elseif (LCASE(imallname)="11st1010") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode211stDlvCode(itenCode)
    elseif (LCASE(imallname)="ssg") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2SsgDlvCode(itenCode)
    elseif (LCASE(imallname)="halfclub") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2HalfClubDlvCode(itenCode)
    elseif (LCASE(imallname)="coupang") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2CoupangDlvCode(itenCode)
    elseif (LCASE(imallname)="hmall1010") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2HmallDlvCode(itenCode)
    elseif (LCASE(imallname)="wmp") or (LCASE(imallname)="wmpfashion") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2WMPDlvCode(itenCode)
    elseif (LCASE(imallname)="lfmall") then
        N_TenDlvCode2CommonDlvCode = itenCode
    elseif (LCASE(imallname)="gsisuper") OR (LCASE(imallname)="yes24") OR (LCASE(imallname)="alphamall") OR (LCASE(imallname)="ohou1010") OR (LCASE(imallname)="wadsmartstore") OR (LCASE(imallname)="withnature1010") OR (LCASE(imallname)="goodshop1010") OR (LCASE(imallname)="casamia_good_com") then
        '// gsisuper => ����API �̿�
        N_TenDlvCode2CommonDlvCode = TenDlvCode2SabangNetDlvCode(itenCode)
    end if
end function



dim sitename, research
dim matchState
Dim siteType, searchType
dim dlvchgexists, dlvfin
sitename = RequestCheckVar(Request("sitename"),32)
research = RequestCheckVar(Request("research"),32)
siteType = RequestCheckVar(Request("siteType"),32)
searchType = RequestCheckVar(Request("searchType"),32)
dlvchgexists= RequestCheckVar(Request("dlvchgexists"),10)
dlvfin      = RequestCheckVar(Request("dlvfin"),10)

if (searchType="") then searchType="sendSongjang"

Dim sqlStr
Dim ArrList
CONST MAXROWS = 500
sqlStr = "select top "&MAXROWS&" T.orderserial, T.OutMallOrderSerial,T.matchitemid,T.matchitemoption "
sqlStr = sqlStr & " ,T.OrgDetailKey, IsNULL(T.sendState,0) as sendState"
sqlStr = sqlStr & " ,D.itemname,D.itemOptionName"
'2023-05-03 ������..�����ȣ '-' -> '' �� ġȯó��
sqlStr = sqlStr & " ,T.itemOrderCount, D.songjangDiv, replace(isNull(D.songjangNo, ''), '-', '') as songjangNo, D.cancelyn, M.cancelyn,M.ipkumdiv"
sqlStr = sqlStr & " ,V.divname, convert(varchar(19),D.beasongdate,21) beasongdate, D.isUpchebeasong, T.sendReqCnt, T.sellsite, T.outMallGoodsNo"
sqlStr = sqlStr & " ,D.idx, IsNull(T.recvSendReqCnt, 0) as recvSendReqCnt, T.receiveName, T.beasongNum11st, T.reserve01 as shppSeq, T.orderItemOption, T.orderItemOptionName, T.shoplinkerorderid "
sqlStr = sqlStr & " ,T.outmalloptionno, T.requireDetail11stYN, T.orgOrderCNT "
sqlStr = sqlStr & " ,convert(varchar(10),d.dlvfinishdt,21) dlvfinishdt, T.outmallOrderseq "
sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T WITH(NOLOCK)"
sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master M WITH(NOLOCK)"
sqlStr = sqlStr & " 	on T.orderserial=M.orderserial"
sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D WITH(NOLOCK)"
sqlStr = sqlStr & " 	on T.orderserial=D.orderserial"
sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid"					'// ���� �ֹ��� ������ ���(����1��,�Ķ�1�� -> �Ķ�2��)
sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption"
sqlStr = sqlStr & " 	and D.currstate=7"
sqlStr = sqlStr & " 	left join db_order.dbo.tbl_songjang_div V WITH(NOLOCK)"
sqlStr = sqlStr & " 	on D.songjangDiv=V.divcd"
sqlStr = sqlStr & " where 1=1 and datediff(m,T.regdate,getdate())<7"						'// �ֱ� 6����
sqlStr = sqlStr & " and T.OrgDetailKey is Not NULL"
sqlStr = sqlStr & " and T.matchState not in ('R','D','B')"              ''��ȯ ��� ����.

if (sitename<>"") then
    sqlStr = sqlStr & " and T.sellsite='"&sitename&"'"
end if

if (siteType="") then
    sqlStr = sqlStr & " and T.sellsite in ('lotteCom','lotteimall', 'lotteon', 'shintvshopping', 'skstoa', 'wetoo1300k', 'interpark','cjmall','lotteComM', 'gseshop', 'homeplus', 'ezwel', 'auction1010', 'gmarket1010', 'nvstorefarm', 'nvstoremoonbangu', 'Mylittlewhoopee', '11st1010','ssg', 'halfclub', 'gsisuper', 'coupang', 'hmall1010', 'WMP', 'wmpfashion', 'lfmall', 'yes24', 'alphamall', 'wconcept1010', 'benepia1010', 'withnature1010', 'goodshop1010', 'ohou1010', 'wadsmartstore','casamia_good_com', 'kakaostore', 'boribori1010') "
    sqlStr = sqlStr & " and not ( (T.sellsite in ('gsisuper', 'yes24', 'alphamall', 'withnature1010', 'goodshop1010', 'ohou1010', 'wadsmartstore','casamia_good_com')) and T.shoplinkerorderid is NULL) "
end if

if (searchType = "sendSongjang") then
	sqlStr = sqlStr & " and IsNULL(T.sendState,0)=0"
elseif (searchType = "sendRecvState") then
	sqlStr = sqlStr & " and ( "
	sqlStr = sqlStr & " 	((T.sellsite = 'cjmall') and (d.songjangDiv not in ('4', '3', '28', '2', '1', '13','18','8'))) "		'// ����Ʈ��ŷ(CJ���� �ٷ� ���μ� Ȯ��) ����(18), ��ü��(8) Ȯ��
	sqlStr = sqlStr & " ) "
	sqlStr = sqlStr & " and IsNULL(T.sendState, 0) <> 0 "
	sqlStr = sqlStr & " and IsNULL(T.recvSendState, 0) < 100 "
    sqlStr = sqlStr & " and d.dlvfinishdt is not null"
	sqlStr = sqlStr & " and DateDiff(d, d.beasongdate, getdate()) >= 1 "
    sqlStr = sqlStr & " and DateDiff(d, d.dlvfinishdt, getdate()) >= 1 "
elseif (searchType = "sendChulgo") then
    sqlStr = sqlStr & " and T.sellsite='ssg'"
    sqlStr = sqlStr & " and T.sendState=2"
    sqlStr = sqlStr & " and IsNULL(T.recvSendState, 0)=0"
else
    sqlStr = sqlStr & " and 1=0 "
	'// ����
end if

if (dlvchgexists<>"") then
    sqlStr = sqlStr & " and Exists(select 1 from db_log.dbo.tbl_songjang_chglast sc WITH(NOLOCK) where sc.odetailidx=d.idx)"
end if

if (dlvfin<>"") then
    sqlStr = sqlStr & " and d.dlvfinishdt is not null"
end if

if (searchType = "sendRecvState") then
    sqlStr = sqlStr & " order by D.beasongdate desc"
else
    sqlStr = sqlStr & " order by D.beasongdate"
end if
''sqlStr = sqlStr & " order by T.OutMallOrderseq"


''rw sqlStr
''response.end


IF (searchType="optchg") then ''�ֱ� �ɼǺ��� ����� - CS �ɼ� ����� xSite_Tmp_Ooder.matchitemoption �� ���� ó���Ǹ� �̰����� ��ȸ�� �ʿ������..
    sqlStr = "exec [db_dataSummary].dbo.[usp_Ten_OUTAMLL_MayOptionChange] '"&sitename&"'"

    db3_dbget.CursorLocation = adUseClient
    db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
    if NOT db3_rsget.Eof then
        ArrList = db3_rsget.getRows
    end if
    db3_rsget.close()

else
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        ArrList = rsget.getRows
    end if
    rsget.Close
end if

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
	//var popwin=window.open('/admin/etc/cjmall/actCJmallSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
	var popwin=window.open('<%=apiURL%>/outmall/cjmall/actCJmallSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
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
	var popwin=window.open('<%=apiURL%>/outmall/gsshop/actGSShopSongjangInputJsonProc.asp?' + params,'sendSongJangGSShop','width=600,height=400,scrollbars=yes,resizable=yes');
    //var popwin=window.open('/admin/etc/gsshop/actGSShopSongjangInputProc.asp?' + params,'sendSongJangGSShop','width=600,height=400,scrollbars=yes,resizable=yes');
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

    /*
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
    */
	var popwin=window.open("<%=apiURL%>/outmall/cjmall/actCjmallRecvStateInputProc.asp?" + params,"sendRecvStateCJMALL","width=600,height=400,scrollbars=yes,resizable=yes");
    popwin.focus();
	comp.disabled=true;
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
    var popwin=window.open('/admin/etc/ezwel/Ezwel_SongjangProc.asp?' + params,'sendRecvStateEzwel','width=600,height=400,scrollbars=yes,resizable=yes');
    // var popwin=window.open('<%=apiURL%>/outmall/proc/Ezwel_SongjangProc.asp?' + params,'sendRecvStateEzwel','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangKakaostore(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }

	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo;
    var popwin=window.open('/admin/etc/kakaostore/kakaostore_SongjangProc.asp?' + params,'sendRecvStatekakaostore','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangBenepia(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }

	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo;
    var popwin=window.open('/admin/etc/benepia/benepia_SongjangProc.asp?' + params,'sendRecvStatebenepia','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangBoribori(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }

	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo;
    var popwin=window.open('/admin/etc/boribori/boribori_SongjangProc.asp?' + params,'sendRecvStateboribori','width=600,height=400,scrollbars=yes,resizable=yes');
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
    var popwin=window.open('<%=apiURL%>/outmall/proc/Auction_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}


function sendSongJangGmarket(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }

	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&songjangDiv="+receiveName;
    var popwin=window.open('<%=apiURL%>/outmall/proc/Gmarket_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangStorefarm(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&songjangDiv="+receiveName;
    var popwin=window.open('<%=apiURL%>/outmall/proc/Nvstorefarm_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangMylittlewhoopee(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&songjangDiv="+receiveName;
    var popwin=window.open('<%=apiURL%>/outmall/proc/Mylittlewhoopee_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangStoremoonbangu(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&songjangDiv="+receiveName;
    var popwin=window.open('<%=apiURL%>/outmall/proc/Nvstoremoonbangu_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJang11st(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, beasongNum){
/*
	alert('11���������۾���');
	alert(beasongNum);
	alert(songjangDiv);
	alert(songjangNo);
	return;
*/
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&songjangDiv="+beasongNum;
    var popwin=window.open('<%=apiURL%>/outmall/proc/11st_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangSSG(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, shppNo, shppSeq, itemno, dlvfinishdt){

    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
//        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
    var params = "shppNo="+shppNo+"&shppSeq="+shppSeq+"&delicoVenId="+songjangDiv+"&wblno="+songjangNo+"&itemno="+itemno+"&outmallorderserial="+OutMallOrderSerial+"&orgdetailKey="+OrgDetailKey+"&dlvfinishdt="+dlvfinishdt;

    <% if (searchType = "sendChulgo") then %>
        params = params + "&prctp=2"
    <% end if %>

   //alert(params);
   //return;
    var popwin=window.open('<%=apiURL%>/outmall/ssg/ssg_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;

}
function sendSongJangHalfclub(comp, OutMallOrderSerial, OrgDetailKey, outmallGoodNo, outmallOptionCode, outmallOptionName, itemno, songjangDiv, songjangNo){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&outmallGoodNo="+outmallGoodNo+"&outmallOptionCode="+outmallOptionCode+"&outmallOptionName="+outmallOptionName+"&itemno="+itemno+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo;
    var popwin=window.open('<%=apiURL%>/outmall/proc/halfclub_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangLfmall(comp, OutMallOrderSerial, OrgDetailKey, outmallGoodNo, outmallOptionCode, outmallOptionName, itemno, songjangDiv, songjangNo){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&outmallGoodNo="+outmallGoodNo+"&outmallOptionCode="+outmallOptionCode+"&outmallOptionName="+outmallOptionName+"&itemno="+itemno+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo;
    var popwin=window.open('<%=apiURL%>/outmall/proc/lfmall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangSabangNet(comp, OutMallOrderSerial, OrgDetailKey, outmallGoodNo, outmallOptionCode, outmallOptionName, itemno, songjangDiv, songjangNo, shoplinkerorderid){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&outmallGoodNo="+outmallGoodNo+"&outmallOptionCode="+outmallOptionCode+"&outmallOptionName="+outmallOptionName+"&itemno="+itemno+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&shoplinkerorderid=" + shoplinkerorderid;
    var popwin=window.open('<%=apiURL%>/outmall/proc/sabangnet_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}
function sendSongJangCoupang(comp, OutMallOrderSerial, OrgDetailKey, outmallGoodNo, outmalloptionno, songjangDiv, songjangNo, beasongNum, splitrequire){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&outmallGoodNo="+outmallGoodNo+"&outmallOptionCode="+outmalloptionno+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&beasongNum="+beasongNum+"&splitrequire="+splitrequire;
    var popwin=window.open('<%=apiURL%>/outmall/proc/coupang_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangHmall(comp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo, beasongNum, reserve01){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&beasongNum="+beasongNum+"&reserve01="+reserve01;
    var popwin=window.open('<%=apiURL%>/outmall/proc/hmall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangWMP(comp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo;
    var popwin=window.open('<%=apiURL%>/outmall/proc/wmp_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangWMPfashion(comp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo;
    var popwin=window.open('<%=apiURL%>/outmall/proc/wmpfashion_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function jsChulgoTargetMake(imallname){
    if (imallname=="ssg"){
        var params = "prctp=999";
        var popwin=window.open('<%=apiURL%>/outmall/ssg/ssg_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }
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
    //var popwin=window.open('/admin/etc/lotteimall/actLotteiMallReq.asp?' + params,'xSiteSongjangInputProciMall','width=600,height=400,scrollbars=yes,resizable=yes');
    var popwin=window.open('<%=apiURL%>/outmall/proc/Lotteimall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
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
	//var popwin=window.open("/admin/etc/orderInput/xSiteCSOrder_lotteimall_Process.asp?" + params,"sendSongJangiMallNew","width=600,height=400,scrollbars=yes,resizable=yes");
	var popwin=window.open('<%=apiURL%>/outmall/proc/Lotteimall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangLotteon(comp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo, outmallGoodNo, outmalloptionCode, beasongNum, orderCount){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
    var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&outmallGoodNo="+outmallGoodNo+"&outmallOptionCode="+outmalloptionCode+"&beasongNum="+beasongNum+"&sendQnt="+orderCount;
    var popwin=window.open('<%=apiURL%>/outmall/proc/Lotteon_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangShintvshopping(comp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo, outmallGoodNo, outmalloptionCode, beasongNum){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
    var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&outmallGoodNo="+outmallGoodNo+"&outmallOptionCode="+outmalloptionCode+"&beasongNum="+beasongNum;
    var popwin=window.open('<%=apiURL%>/outmall/proc/shintvshopping_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangSkstoa(comp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo, outmallGoodNo, outmalloptionCode){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
    var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&outmallGoodNo="+outmallGoodNo+"&outmallOptionCode="+outmalloptionCode;
    var popwin=window.open('<%=apiURL%>/outmall/proc/skstoa_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangWetoo1300k(comp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo, outmallGoodNo){
    if (songjangDiv==""){
        alert("���޻� �ù���ڵ� ������");
        return;
    }

    if (songjangNo==""){
        alert("�����ȣ ������..");
        return;
    }
    var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&outmallGoodNo="+outmallGoodNo;
    var popwin=window.open('<%=apiURL%>/outmall/proc/wetoo1300k_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
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
    //var popwin=window.open('/admin/etc/interparkXML/actInterparkSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    var popwin=window.open('<%=apiURL%>/outmall/interpark/actInterparkSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}
function pop_mayDelOrder(sitename){
    var popwin=window.open('/admin/etc/orderInput/pop_etcSiteSongjangInput.asp?sitename=' + sitename,'pop_mayDelOrder','width=1500,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=<%=menupos%>&page=1&research=on";
	iUrl += "&sellsite=<%=sitename%>"
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,crollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function fnsugiProc(outmallorderseq) {
	if (confirm('���� ó���Ǿ��� �� ���޸����� Ȯ���� �ʿ��մϴ�.\n\n���� ó�� �Ͻðڽ��ϱ�?')){
        var popupInvoice = window.open('/admin/etc/orderInput/invoice_sugi_process.asp?outmallorderseq=' + outmallorderseq,'fnsugiProc','width=1500,height=500,scrollbars=yes,resizable=yes');
        popupInvoice.focus();
    }
}
</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<!-- �˻� ���� -->



<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
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
			�� 20~30�и��� �ڵ� ó���� (15, 45)
		    &nbsp;&nbsp;
		    �˻����� :
			<select class="select" name="searchType">
	     	<option value="sendSongjang" <%= chkIIF(searchType="sendSongjang","selected","") %> >��������</option>
	     	<option value="sendRecvState" <%= chkIIF(searchType="sendRecvState","selected","") %> >���μ�����(Cjmall:�Ϻ��ù��)</option>
	     	<option value="sendChulgo" <%= chkIIF(searchType="sendChulgo","selected","") %> >�������(SSG)</option>
            <option value="optchg" <%= chkIIF(searchType="optchg","selected","") %> >�ɼǺ��濹�󳻿�</option>
	     	</select>

	     	<% if (searchType="sendChulgo") then %>
	     	<input type="button" value="������ۼ�" onClick="jsChulgoTargetMake('ssg');">
	        <% end if %>
            &nbsp;&nbsp;
            <input type="checkbox" name="dlvchgexists" <%=CHKIIF(dlvchgexists="on","checked","")%>>���庯�����系����
            &nbsp;&nbsp;
            <input type="checkbox" name="dlvfin" <%=CHKIIF(dlvfin="on","checked","")%>>��ۿϷ᳻����
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
	     	<% If (sitename = "auction1010" or sitename = "gseshop") and (session("ssBctID")="kjy8517") Then %>
				&nbsp;&nbsp;<input type="button" value="�����ֹ���" class="button" onclick="pop_mayDelOrder('<%=sitename%>');">
     		<% End If %>
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
        <td width="80" class="td_br">����Ŀ<br />�ֹ���ȣ</td>
      	<td width="60" class="td_br">����<br>��ǰ�ڵ�</td>
      	<td width="50" class="td_br">��ǰ�ڵ�</td>
      	<td width="50" class="td_br">�ɼ��ڵ�</td>
      	<td  class="td_br">��ǰ�� <font color="blue">[�ɼǸ�]</font></td>
        <td width="30" class="td_br">����</td>
        <td width="100" class="td_br">�����</td>
        <td width="80" class="td_br">��ۿϷ���</td>
        <td width="40" class="td_br">���<br>����</td>
      	<td class="td_br">�ù��</td>
      	<td class="td_br">����</td>
      	<td class="td_br">���޻�<br>�ù��ڵ�</td>
      	<td class="td_br">����</td>
      	<td class="td_br">����<br>Ƚ��</td>
        <td class="td_br">���</td>
    </tr>
<% if (IsArray(ArrList)) THEN %>
<%
Dim intRows : intRows = UBound(ArrList,2)
dim iOutMallDlvCode
for i=0 to intRows
    iOutMallDlvCode = ""
    iOutMallDlvCode = N_TenDlvCode2CommonDlvCode(ArrList(18,i),ArrList(9,i))

    if (ArrList(18,i)="ssg") then
        if (ArrList(9,i)="99") then
           '' ArrList(10,i) = ArrList(10,i)&RIGHT(replace(replace(replace(ArrList(15,i),":",""),"-","")," ",""),6)
        end if
    end if
%>
<tr>
    <!--<td class="td_br"><input type="checkbox" name="cksel" value=""></td> -->
    <td class="td_br"><%= ArrList(18,i) %></td>
    <td class="td_br"><%= ArrList(0,i) %></td>
    <td class="td_br"><a href="#" onClick="popByExtorderserial('<%= ArrList(1,i) %>');return false;"><%= ArrList(1,i) %></a></td>
    <td class="td_br"><%= ArrList(27,i) %></td>
    <td class="td_br"><%= ArrList(4,i) %></td>
    <td class="td_br"><%= ArrList(2,i) %></td>
    <td class="td_br"><%= ArrList(3,i) %></td>
    <td class="td_br"><%= ArrList(6,i) %>
    <% if ArrList(7,i)<>"" then %>
    <font color="blue">[<%= ArrList(7,i) %>]</font>
    <% end if %>
    </td>
    <td class="td_br" align="center"><%= ArrList(8,i) %></td>
    <td class="td_br" align="center"><%= ArrList(15,i) %></td>
    <td class="td_br" align="center"><%= ArrList(31,i) %></td>
    <td class="td_br" align="center">
        <% IF ArrList(16,i)="Y" THEN %>
        <font color="blue">��ü</font>
        <% End IF %>
    </td>
    <td class="td_br"><%= ArrList(14,i) %></td>
    <td class="td_br" <%=CHKIIF(isNULL(ArrList(10,i)),"bgcolor='#CC2222'","") %> ><%= ArrList(10,i) %></td>
    <% if (searchType="optchg") then %>
    <td class="td_br"></td>
    <td class="td_br"></td>
    <td class="td_br"></td>
    <% else %>
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
			Case "benepia1010"
		%>
				<input type="button" value="����" onClick="sendSongJangBenepia(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
			Case "kakaostore"
		%>
				<input type="button" value="����" onClick="sendSongJangKakaostore(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
			Case "boribori1010"
		%>
				<input type="button" value="����" onClick="sendSongJangBoribori(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>');">
		<%
			Case "auction1010"
		%>
				<input type="button" value="����" onClick="sendSongJangAuction(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>', '<%=ArrList(9,i)%>');">
		<%
			Case "gmarket1010"
		%>
				<input type="button" value="����" onClick="sendSongJangGmarket(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>', '<%=ArrList(9,i)%>');">

		<%
			Case "nvstorefarm"
		%>
				<input type="button" value="����" onClick="sendSongJangStorefarm(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>', '<%=ArrList(9,i)%>');">
		<%
			Case "nvstoremoonbangu"
		%>
				<input type="button" value="����" onClick="sendSongJangStoremoonbangu(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>', '<%=ArrList(9,i)%>');">
		<%
			Case "Mylittlewhoopee"
		%>
				<input type="button" value="����" onClick="sendSongJangMylittlewhoopee(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>', '<%=ArrList(9,i)%>');">
		<%
			Case "11st1010"
		%>
				<input type="button" value="����" onClick="sendSongJang11st(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>', '<%=ArrList(23,i)%>');">
		<%
			Case "ssg"
		%>
				<input type="button" value="����" onClick="sendSongJangSSG(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>', '<%=ArrList(23,i)%>', '<%=ArrList(24,i)%>', '<%=ArrList(8,i)%>', '<%=ArrList(31,i)%>');">
		<%
			Case "halfclub"
		%>
				<input type="button" value="����" onClick="sendSongJangHalfclub(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= ArrList(19,i) %>', '<%= ArrList(25,i) %>', '<%= ArrList(26,i) %>', '<%= ArrList(30,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>');" >
		<%
			Case "lfmall"
		%>
				<input type="button" value="����" onClick="sendSongJangLfmall(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= ArrList(19,i) %>', '<%= ArrList(25,i) %>', '<%= ArrList(26,i) %>', '<%= ArrList(30,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>');" >
		<%
			Case "gsisuper", "yes24", "alphamall", "ohou1010", "wadsmartstore", "casamia_good_com", "wconcept1010", "withnature1010", "goodshop1010"
                '// ����API �� ����
		%>
				<input type="button" value="����" onClick="sendSongJangSabangNet(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= ArrList(19,i) %>', '<%= ArrList(25,i) %>', '<%= ArrList(26,i) %>', '<%= ArrList(8,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>', '<%= ArrList(27,i) %>');" >
		<%
			Case "coupang"
		%>
				<input type="button" value="����" onClick="sendSongJangCoupang(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= ArrList(19,i) %>', '<%= ArrList(28,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>', '<%= ArrList(23,i) %>', '<%= ArrList(29,i) %>');" >
		<%
			Case "hmall1010"
		%>
				<input type="button" value="����" onClick="sendSongJangHmall(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>', '<%= ArrList(23,i) %>', '<%= ArrList(24,i) %>');" >
		<%
			Case "WMP"
		%>
				<input type="button" value="����" onClick="sendSongJangWMP(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>');" >
		<%
			Case "wmpfashion"
		%>
				<input type="button" value="����" onClick="sendSongJangWMPfashion(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>');" >
		<%
			Case "lotteon"
		%>
				<input type="button" value="����" onClick="sendSongJangLotteon(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>', '<%= ArrList(19,i) %>', '<%= ArrList(28,i) %>', '<%= ArrList(23,i) %>', '<%= ArrList(30,i) %>');" >
		<%
			Case "shintvshopping"
		%>
				<input type="button" value="����" onClick="sendSongJangShintvshopping(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>', '<%= ArrList(19,i) %>', '<%= ArrList(28,i) %>', '<%= ArrList(23,i) %>');" >
		<%
			Case "skstoa"
		%>
				<input type="button" value="����" onClick="sendSongJangSkstoa(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>', '<%= ArrList(19,i) %>', '<%= ArrList(28,i) %>');" >
		<%
			Case "wetoo1300k"
		%>
				<input type="button" value="����" onClick="sendSongJangWetoo1300k(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>', '<%= ArrList(19,i) %>');" >
		<%
			Case Else
				response.write "ERR"
		End Select
		%>
	<% elseif ((ArrList(5,i) <> 0) and (searchType = "sendRecvState")) then %>
		<!-- ���μ� Ȯ�� �� ���� -->
        <%
        	if ArrList(18,i)="cjmall" then
        		If (FALSE) and ArrList(9,i) = "21" Then		'�浿�ù�//Ȩ������ ��û����..�ٸ� ������� �ؾ� �� ��..2015-02-27 19:04 ������
		%>
		<input type="button" class="button_s" value="�浿Ȯ��" onClick="sendRecvStateCJMALL(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= ArrList(9,i) %>','<%= ArrList(10,i) %>','<%= ArrList(22,i) %>', '1');">
		<%
        		Else
        %>
        <input type="button" value="�μ�����" onClick="sendRecvStateCJMALL(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= ArrList(9,i) %>','<%= ArrList(10,i) %>','.', '2');">
        <%
        		End If
        	end if
        %>
    <% elseif ((ArrList(5,i) <> 0) and (searchType = "sendChulgo")) then %>
        <%
        	if ArrList(18,i)="ssg" then
		%>
				<input type="button" value="���ó��" onClick="sendSongJangSSG(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","������",ArrList(10,i)) %>', '<%=ArrList(23,i)%>', '<%=ArrList(24,i)%>', '<%=ArrList(8,i)%>', '<%=ArrList(31,i)%>');">
		<%
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
		    <% if ArrList(17,i)>2 then %>
			<b><%= ArrList(17,i) %></b>
			<% else %>
			<%= ArrList(17,i) %>
			<% end if %>
			/
			<!-- ���μ� Ȯ�� �� ���� -->
			<% if ArrList(21,i)>2 then %>
			<b><%= ArrList(21,i) %></b>
			<% else %>
			<%= ArrList(21,i) %>
			<% end if %>
		<% end if %>
    </td>
    <% end if %>
    <td>
        <input type="button" class="button" value="����" onclick="fnsugiProc('<%= ArrList(32,i) %>');">
    </td>
</tr>
<% next %>
<% ELSE %>
<tr>
    <td colspan="17" align="center">[�˻� ����� �����ϴ�.]</td>
</tr>
<% end if %>
</table>
<% if (sitename="") then %>
<script>//alert('���θ��� ���� �ϼ���.');</script>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->