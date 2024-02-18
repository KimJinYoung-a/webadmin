<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �ֹ��α���ȸ
' History : 2020.01.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<%
dim page, research, yyyy1, mm1, yyyy2, mm2, dd1,dd2, stdate, nowdate,searchnextdate, i, orderserial
dim userid, paygateid, authno, price, buyname, buyhp, reqname, reqhp, isqueryexecute, temp_idx, resultMessage, isPay, isSuccess
    page = requestCheckvar(getNumeric(request("page")),10)
    yyyy1   = requestCheckvar(getNumeric(request("yyyy1")),4)
    yyyy2   = requestCheckvar(getNumeric(request("yyyy2")),4)
    mm1     = requestCheckvar(getNumeric(request("mm1")),2)
    mm2     = requestCheckvar(getNumeric(request("mm2")),2)
    dd1     = requestCheckvar(getNumeric(request("dd1")),2)
    dd2     = requestCheckvar(getNumeric(request("dd2")),2)
    research= requestCheckvar(request("research"),2)
    orderserial= requestCheckvar(request("orderserial"),11)
    userid= requestCheckvar(request("userid"),32)
    paygateid= requestCheckvar(request("paygateid"),50)
    authno= requestCheckvar(request("authno"),32)
    price= requestCheckvar(getNumeric(request("price")),10)
    buyname= requestCheckvar(request("buyname"),32)
    buyhp= requestCheckvar(request("buyhp"),16)
    reqname= requestCheckvar(request("reqname"),32)
    reqhp= requestCheckvar(request("reqhp"),16)
    temp_idx = requestCheckvar(getNumeric(request("temp_idx")),10)
    resultMessage = requestCheckvar(request("resultMessage"),20)
    isPay = requestCheckvar(request("isPay"),1)
    isSuccess = requestCheckvar(request("isSuccess"),1)

isqueryexecute=false
nowdate = Left(CStr(now()),10)
stdate = Left(CStr(DateAdd("d",-365,now())),10)
if page="" then page=1
if (yyyy1="") then
    yyyy1 = Left(stdate,4)
	mm1   = Mid(stdate,6,2)
	dd1   = Mid(stdate,9,2)

	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

'searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)
searchnextdate = DateAdd("d",dateserial(yyyy2 , mm2 , dd2),1)

if not(orderserial="" and userid="" and paygateid="" and authno="" and price="" and buyname="" and buyhp="" and reqname="" and reqhp="" and temp_idx="" and resultMessage="" and isPay="" and isSuccess="") then
    isqueryexecute=true
end if

dim ojumun
set ojumun = new CJumunMaster
    ojumun.FCurrPage = page
    ojumun.FPageSize = 100
    ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
    ojumun.FRectRegEnd = searchnextdate
    ojumun.FRectorderserial = orderserial
    ojumun.FRectuserid = userid
    ojumun.FRectpaygateid = paygateid
    ojumun.FRectauthcode = authno
    ojumun.FRectprice = price
    ojumun.FRectbuyname = buyname
    ojumun.FRectbuyhp = buyhp
    ojumun.FRectreqname = reqname
    ojumun.FRectreqhp = reqhp
    ojumun.FRecttemp_idx = temp_idx
    ojumun.FRectResultMessage = resultMessage
    ojumun.FRectIsPay = isPay
    ojumun.FRectIsSuccess = isSuccess
    

    if isqueryexecute then
        ojumun.getordertemplist

        ' �����Ͱ� ���°�� ������̺� �˻�
        if ojumun.FResultCount<1 then
            ojumun.FRectOldJumun = "on"
            ojumun.getordertemplist
        end if
    end if
%>

<script type='text/javascript'>

function NextPage(ipage){
    //�ϴ���
    var startdate = frm_search.yyyy1.value + "-" + frm_search.mm1.value + "-" + frm_search.dd1.value;
    var enddate = frm_search.yyyy2.value + "-" + frm_search.mm2.value + "-" + frm_search.dd2.value;
    var diffDay = 0;
    var start_yyyy = startdate.substring(0,4);
    var start_mm = startdate.substring(5,7);
    var start_dd = startdate.substring(8,startdate.length);
    var sDate = new Date(start_yyyy, start_mm-1, start_dd);
    var end_yyyy = enddate.substring(0,4);
    var end_mm = enddate.substring(5,7);
    var end_dd = enddate.substring(8,enddate.length);
    var eDate = new Date(end_yyyy, end_mm-1, end_dd);

    diffDay = Math.ceil((eDate.getTime() - sDate.getTime())/(1000*60*60*24));

    if (diffDay > 365){
        alert('1�� ������ �˻��� �����մϴ�.');
        return;
    }

	document.frm_search.page.value= ipage;
	document.frm_search.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm_search" method="GET" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
    <td align="left">
        * �ֹ��� : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
    </td>  
    <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="�˻�" onClick="NextPage('');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        <label>* �ֹ���ȣ : <input type="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="11" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
        <label style="margin-left:7px;">* ���̵� : <input type="text" name="userid" value="<%= userid %>" size="15" maxlength="32" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
        <label style="margin-left:7px;">* TID : <input type="text" name="paygateid" value="<%= paygateid %>" size="25" maxlength="50" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
        <label style="margin-left:7px;">* ���ι�ȣ : <input type="text" name="authno" value="<%= authno %>" size="15" maxlength="32" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
        <label style="margin-left:7px;">* �������ֹ���ȣ : <input type="text" name="temp_idx" value="<%= temp_idx %>" size="10" maxlength="10" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
        <label style="margin-left:7px;">* �ݾ� : <input type="text" name="price" value="<%= price %>" size="10" maxlength="10" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
        <Br>
        <label>* �ֹ��� : <input type="text" name="buyname" value="<%= buyname %>" size="10" maxlength="32" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
        <label style="margin-left:7px;">* �ֹ����޴��� : <input type="text" name="buyhp" value="<%= buyhp %>" size="15" maxlength="16" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
        <label style="margin-left:7px;">* ������ : <input type="text" name="reqname" value="<%= reqname %>" size="10" maxlength="32" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
        <label style="margin-left:7px;">* �������޴��� : <input type="text" name="reqhp" value="<%= reqhp %>" size="15" maxlength="16" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
        <label style="margin-left:7px;">
            * ���俩�� :
            <select name="isPay" class="select">
            <option value="">��ü</option>
            <option value="Y" <%=chkIIF(isPay="Y","selected","")%>>����</option>
            <option value="N" <%=chkIIF(isPay="N","selected","")%>>����</option>
            </select>
        </label>
        <label style="margin-left:7px;">
            * �ֹ���� :
            <select name="isSuccess" class="select">
            <option value="">��ü</option>
            <option value="Y" <%=chkIIF(isSuccess="Y","selected","")%>>����</option>
            <option value="N" <%=chkIIF(isSuccess="N","selected","")%>>����</option>
            </select>
        </label>
        <label style="margin-left:7px;">* ����޽��� : <input type="text" name="resultMessage" value="<%= resultMessage %>" size="20" onKeyPress="if(window.event.keyCode==13) NextPage('');" class="text" /></label>
    </td>
</tr>
</table>
</form>

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
        <% if not(isqueryexecute) then %>
            <strong><font color="red">�˻������� 1�� �̻� �Է��ϼž� �˻��� �˴ϴ�.</font></strong>
        <% end if %>
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="20">
        �˻���� : <b><%= ojumun.FTotalCount %></b>
        &nbsp;
        ������ : <b><%= page %> / <%= ojumun.FTotalPage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td align="center" width="60">������<Br>�ֹ���ȣ</td>
    <td align="center" width="80">�ֹ���</td>
    <td align="center" width="80">�ֹ���ȣ</td>
    <td align="center" width="90">���̵�</td>
    <td align="center" width="90">PG��<Br>(PG���̵�)</td>
    <td align="center" width="80">�������</td>
    <td align="center" width="50">PG��<Br>���俩��</td>
    <td align="center">����޼���</td>
    <td align="center" width="30">�ֹ�<Br>���</td>
    <td align="center" width="200">TID</td>
    <td align="center" width="80">���ι�ȣ</td>
    <td align="center" width="60">�ݾ�</td>
    <td align="center" width="60">�ֹ���</td>
    <td align="center" width="100">�ֹ����޴���</td>
    <td align="center" width="60">������</td>
    <td align="center" width="100">�������޴���</td>
    <td align="center" width="60">ä��</td>
</tr>

<% if ojumun.FResultCount>0 then %>
    <% for i=0 to ojumun.FResultCount -1 %>
    <tr bgcolor="#FFFFFF" align="center">
        <td><%= ojumun.FItemList(i).ftemp_idx %></td>
        <td><%= FormatDate(ojumun.FItemList(i).fregdate,"0000-00-00 00:00:00") %></td>
        <td><%= ojumun.FItemList(i).forderserial %></td>
        <td>
            <% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 or C_CriticInfoUserLV3 then %>
                <%= ojumun.FItemList(i).fuserid %>
            <% else %>
                <%= printUserId(ojumun.FItemList(i).fuserid,1,"*") %>
            <% end if %>
        </td>
        <td>
            <%= fnGetPggubunName(ojumun.FItemList(i).fpggubun) %>
            <% if ojumun.FItemList(i).fmid<>"" and not(isnull(ojumun.FItemList(i).fmid)) then %>
                (<%= ojumun.FItemList(i).fmid %>)
            <% end if %>
        </td>
        <td><%= ojumun.FItemList(i).JumunMethodName %></td>
        <td><%= ojumun.FItemList(i).fispay %></td>
        <td align="left"><%= ojumun.FItemList(i).fp_rmesg1 %></td>
        <td>
            <% if ojumun.FItemList(i).fissuccess<>"" and not(isnull(ojumun.FItemList(i).fissuccess)) then %>
                <% if ojumun.FItemList(i).fissuccess then %>
                    Y
                <% else %>
                    N
                <% end if %>
            <% end if %>
        </td>
        <td><%= ojumun.FItemList(i).ftid %></td>
        <td><%= ojumun.FItemList(i).fauth_no %></td>
        <td><%= FormatNumber(ojumun.FItemList(i).fprice,0) %></td>
        <td>
            <% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 or C_CriticInfoUserLV3 then %>
                <%= ojumun.FItemList(i).fbuyname %>
            <% else %>
                <%= printUserId(ojumun.FItemList(i).fbuyname,1,"*") %>
            <% end if %>
        </td>
        <td><%= ojumun.FItemList(i).fbuyhp %></td>
        <td>
            <% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 or C_CriticInfoUserLV3 then %>
                <%= ojumun.FItemList(i).freqname %>
            <% else %>
                <%= printUserId(ojumun.FItemList(i).freqname,1,"*") %>
            <% end if %>
        </td>
        <td><%= ojumun.FItemList(i).freqhp %></td>
        <td><%= ojumun.FItemList(i).frdsite %></td>
    </tr>
    <% next %>

    <tr bgcolor="#FFFFFF" height="25">
        <td colspan="20" align="center">
            <% if ojumun.HasPreScroll then %>
                <a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
            <% else %>
                [pre]
            <% end if %>

            <% for i=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
                <% if i>ojumun.FTotalpage then Exit for %>
                <% if CStr(page)=CStr(i) then %>
                <font color="red">[<%= i %>]</font>
                <% else %>
                <a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
                <% end if %>
            <% next %>

            <% if ojumun.HasNextScroll then %>
                <a href="javascript:NextPage('<%= i %>')">[next]</a>
            <% else %>
                [next]
            <% end if %>
        </td>
    </tr>
<% else %>
    <tr height="25" bgcolor="FFFFFF" align="center">
        <td colspan="20">
            [�˻������ �����ϴ�.]
        </td>
    </tr>
<% end if %>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

