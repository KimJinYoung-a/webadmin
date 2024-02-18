<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs 메모
' History : 2007.10.26 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%

dim i, userid, orderserial, phoneNumer
userid      = requestCheckVar(request("userid"),32)
orderserial = requestCheckVar(request("orderserial"),32)
phoneNumer  = requestCheckVar(request("phoneNumer"),32)

'==============================================================================
dim ocsmemo
set ocsmemo = New CCSMemo

ocsmemo.FRectUserID = userid
ocsmemo.FRectOrderserial = orderserial
ocsmemo.FRectPhoneNumber = phoneNumer

if (userid <> "") or (orderserial<>"") or (phoneNumer<>"") then
    ocsmemo.GetCSMemoList
end if

%>
<link href="/cscenter/js/css/custom-theme/jquery-ui-1.9.2.css" rel="stylesheet">
<script src="/cscenter/js/jquery-1.8.3.js"></script>
<script src="/cscenter/js/jquery-ui-1.9.2.js"></script>
<script>
function GotoHistoryMemoMidify(divcd,id,userid,orderserial) {
    //err 처리 추가.
    if (top.callring){
        top.document.all.callring.src = "/cscenter/ippbxmng/CallRingWithOrderFrame.asp?id=" + id;
    }else{
        top.opener.top.header.i_ippbxmng.popCallRing('','','',id,'','');
    }
	//var popwin = window.open("/cscenter/history/history_memo_write.asp?divcd="+divcd+"&id=" + id + "&backwindow=" + "opener" + "&userid=" + userid + "&orderserial=" + orderserial,"GotoHistoryMemoMidify","width=600 height=400 scrollbars=yes resizable=yes");
	//popwin.focus();
}

function ShowThisItem(frm) {
    var e, t;

    if (!frm.orderdetailidx) return;

    for (var i = 0; i < frm.orderdetailidx.length; i++) {
        e = frm.orderdetailidx[i];
        t = frm.orderdetailidx[i];

        if (e.type == "checkbox") {
			while (t.tagName != "TR") {
				t = t.parentElement;
			}

			if (e.checked == true) {
				t.style.display = '';
			} else {
				t.style.display = 'none';
			}
        }
    }
}

function ShowThisItem(comp, obj) {
	var x = document.getElementById(obj);

    if (comp.value == "보기") {
        x.style.display = "inline";
		comp.value = "닫기";
    }else{
        x.style.display = "none";
		comp.value = "보기";
    }
}

</script>
<link rel="stylesheet" href="/cscenter/css/cs.css" type="text/css">
<style>
body {
    background-color: #FFFFFF;
}

.listSep {
	border-top:0px #CCCCCC solid; height:1px; margin:0; padding:0;
}
</style>
<table width="100%" border=0 cellspacing=1 cellpadding=1 class=a bgcolor="F4F4F4">
<% if ocsmemo.FResultCount > 0 then %>
    <tr align="center" bgcolor="F3F3FF">
        <td height="20" width="50">구분</td>
    	<td width="50">idx</td>
     	<td width="80">고객ID</td>
    	<td width="80">주문번호</td>
    	<td>내용</td>
        <td width="80">등록자</td>
    	<td width="80">접수일</td>
    	<td width="30">완료</td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC" style="border-top:1px"></td>
    </tr>

	<% for i = 0 to (ocsmemo.FResultCount - 1) %>
    <tr align="center" bgcolor="FFFFFF">
        <td height="20" >
     	  	 <%= ocsmemo.FItemList(i).GetDivCDName %>
        </td>
    	<td>
    		<%= ocsmemo.FItemList(i).Fid %>
    	</td>
     	<td><%= ocsmemo.FItemList(i).Fuserid %></td>
    	<td><%= ocsmemo.FItemList(i).Forderserial %></td>
    	<td align="left">
			<a href="javascript:GotoHistoryMemoMidify('<%= ocsmemo.FItemList(i).fdivcd %>','<%= ocsmemo.FItemList(i).Fid %>','<%= ocsmemo.FItemList(i).Fuserid %>','<%= ocsmemo.FItemList(i).Forderserial %>')">
    			<span onmouseover="csMemo<%= ocsmemo.FItemList(i).Fid %>.style.display='block';" onmouseout="csMemo<%= ocsmemo.FItemList(i).Fid %>.style.display='none';"><%= DDotFormat(ocsmemo.FItemList(i).Fcontents_jupsu,25) %></span>
				<div id='csMemo<%= ocsmemo.FItemList(i).Fid %>' style='display:none; position:absolute; border:solid 1px #000000; width:400px; padding:5px; background-color:#ffffff; text-decoration: none;'><%= nl2br(ocsmemo.FItemList(i).Fcontents_jupsu) %></div>
			</a>
		</td>
        <td><%= ocsmemo.FItemList(i).Fwriteuser %></td>
    	<td><acronym title="<%= ocsmemo.FItemList(i).Fregdate %>"><%= Left(ocsmemo.FItemList(i).Fregdate,10) %></acronym></td>
    	<td><% if (ocsmemo.FItemList(i).Ffinishyn = "Y") then %>완료<% end if %></td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC"></td>
    </tr>
	<% next %>

<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="6" align="center">검색결과가 없습니다.</td>
    </tr>
<% end if %>

</table>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
