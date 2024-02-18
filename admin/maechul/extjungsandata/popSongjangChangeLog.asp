<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���庯�泻�� ����
' Hieditor : 2019.08.30 eastone
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
Dim i
dim research : research = requestCheckvar(request("research"),10)
dim sellsite : sellsite = requestCheckvar(request("sellsite"),32)
dim page : page = requestCheckvar(request("page"),10)

dim sitescope : sitescope = requestCheckvar(request("sitescope"),10)
dim noinccmt : noinccmt = requestCheckvar(request("noinccmt"),10)
dim noxjungsan : noxjungsan = requestCheckvar(request("noxjungsan"),10)

dim yyyy1, mm1
yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)

if (yyyy1="") then yyyy1=LEFT(NOW(),4)
if (mm1="") then mm1=MID(NOW(),6,2)
if (page="") then page=1

dim oSongjangChgLog
SET oSongjangChgLog = new CDeliveryTrack
oSongjangChgLog.FPageSize = 1000
oSongjangChgLog.FCurrPage = page
oSongjangChgLog.FRectSitename = sellsite
oSongjangChgLog.FRectStartDate = yyyy1+"-"+mm1+"-01"
oSongjangChgLog.FRectEndDate = CStr(dateadd("d",-1,dateadd("m",1,yyyy1+"-"+mm1+"-01")))
oSongjangChgLog.FRectSiteScope = sitescope
oSongjangChgLog.FRectNotIncComment = noinccmt
oSongjangChgLog.FRectNotIncMapXjungsan = noxjungsan

oSongjangChgLog.getSongjangChangeLogListWithCmt

dim FormatDotNo : FormatDotNo=0
%>
<script language='javascript'>
function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=<%=menupos%>&page=1&research=on";
	iUrl += "&sellsite=<%=sellsite%>"
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,crollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function popJcomment(iorderserial,iitemid,iitemoption,isadd){
    var addcmt = "";
   // if (isadd){
        addcmt = prompt("���� comment", "");
        if (addcmt == null) return;

        if (addcmt.length<1){
            alert("�ڸ�Ʈ�� �ۼ����ּ���.");
            return;
        }

        var frm = document.frmcmt;
        frm.orderserial.value=iorderserial;
        frm.itemid.value=iitemid;
        frm.itemoption.value=iitemoption;
        frm.addcomment.value=addcmt;

        frm.submit();
   // }else{

   // }
}

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���޸�:
		<%= getJungsanXsiteComboHTML("sellsite",sellsite,"") %>
		&nbsp;
		
		* ���庯���:
		<% DrawYMBox yyyy1,mm1 %>
        &nbsp;
        * �˻� ����
        <select class="select" name="sitescope">
        <option value="" <%=CHKIIF(sitescope="","selected","") %> >��ü
        <option value="50" <%=CHKIIF(sitescope="50","selected","") %> >���޸�
		<option value="10" <%=CHKIIF(sitescope="10","selected","") %> >�ڻ��
        </select>
        &nbsp;
		<input type="checkbox" name="noxjungsan" <%=CHKIIF(noxjungsan<>"","checked","")%> >�������곻�����ξ��³�����
		&nbsp;
		<input type="checkbox" name="noinccmt" <%=CHKIIF(noinccmt<>"","checked","")%> >�����ڸ�Ʈ���³�����

	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" style="width:70px;height:50px;" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF" >
	<td>
		<%= getExtsongjangInputNOTIStr %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<p  >
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oSongjangChgLog.FTotalcount %></b>
		&nbsp;
		<% if oSongjangChgLog.FTotalcount>=oSongjangChgLog.FPageSize then %>
        (�ִ� <%=FormatNumber(oSongjangChgLog.FPageSize,0)%> ��)
        <% end if %>
	</td>
</tr>

<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="50">DtlIdx</td>
    <td width="60">SITE</td>
    <td width="60">����<br>�ֹ���ȣ</td>
    <td width="80">�ּ�</td>
    <td width="60">�ֹ���ȣ</td>
    <td width="60">��ǰ�ڵ�<br>�ɼ��ڵ�</td>
    <!-- RPA �ν� �����ϰ� �ϱ� ���� �и� -->
    <!--td width="80">��ǰ��<br>(�ɼǸ�)</td-->
    <td width="80">��ǰ��</td>
    <td width="80">�ɼǸ�</td>
    <!-- RPA �ν� �����ϰ� �ϱ� ���� �и� -->
    <!--td width="80">�����ù��<br>�����ù��</td>
    <td width="80">���������ȣ<br>��������ȣ</td-->
    <td width="80">�����ù��</td>
    <td width="80">�����ù��</td>
    <td width="80">���������ȣ</td>
    <td width="80">��������ȣ</td>

    <td width="80">������</td>
    <td width="70">����������</td>

    <td width="70">�����ù��</td>
    <td width="80">��������ȣ</td>
    <td width="70">�����</td>
    <td width="70">�����</td>
    <td width="70">(�ڻ�)������</td>
    
    <td width="60">���</td>
</tr>

<% if oSongjangChgLog.FresultCount<1 then %>
<tr align="center" bgcolor="FFFFFF" >
    <td colspan="17">
        [�˻������ �����ϴ�.]
    </td>
</tr>
<% else %>
<% for i=0 to oSongjangChgLog.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" >
    <td ><%=oSongjangChgLog.FItemList(i).Fodetailidx %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fsitename %></td>
    <td ><a href="#" onClick="popByExtorderserial('<%= oSongjangChgLog.FItemList(i).Fauthcode %>');return false;"><%= oSongjangChgLog.FItemList(i).Fauthcode %></a></td>
    <td ><%=oSongjangChgLog.FItemList(i).Freqzipaddr %></td>
    <td><%=oSongjangChgLog.FItemList(i).Forderserial %></td>
    <td>
        <%=oSongjangChgLog.FItemList(i).FItemid %>
        <br><%=oSongjangChgLog.FItemList(i).FItemOption %>
    </td>
    <!--td >
        <%'oSongjangChgLog.FItemList(i).FItemName %>
        <%' if oSongjangChgLog.FItemList(i).FItemOptionName<>"" then %>
        <br><font color="blue"><%'oSongjangChgLog.FItemList(i).FItemOptionName %></font>
        <%' end if %>
    </td-->
    <td>
        <%=oSongjangChgLog.FItemList(i).FItemName %>
    </td>
    <td>
        <% if oSongjangChgLog.FItemList(i).FItemOptionName<>"" then %>
        <font color="blue"><%=oSongjangChgLog.FItemList(i).FItemOptionName %></font>
        <% end if %>    
    </td>
    <!--td>
        <%'getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fpsongjangdiv,1) %><br>
        <%'getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fchgsongjangdiv,1) %>
    </td-->
    <td>
        <%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fpsongjangdiv,1) %>
    </td>
    <td>
        <%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fchgsongjangdiv,1) %>    
    </td>
    <!--td>
        <%'oSongjangChgLog.FItemList(i).Fpsongjangno %><br>
        <%'oSongjangChgLog.FItemList(i).Fchgsongjangno %>
    </td-->
    <td>
        <%=oSongjangChgLog.FItemList(i).Fpsongjangno %>
    </td>
    <td>
        <%=oSongjangChgLog.FItemList(i).Fchgsongjangno %>
    </td>
    <td><%=oSongjangChgLog.FItemList(i).Fchguserid %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fupddt %></td> 
    <td><%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fsongjangdiv,1) %></td>
    <td>
        <a target="_dlv2" href="<%=getTrackNaverURIByTrName(oSongjangChgLog.FItemList(i).Fsongjangdiv,oSongjangChgLog.FItemList(i).Fsongjangno)%>"><%=oSongjangChgLog.FItemList(i).Fsongjangno %></a>
    </td>
    <td><%=oSongjangChgLog.FItemList(i).Fbeasongdate %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fdlvfinishdt %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fjungsanfixdate %></td> 
    
    <td>
        <% if isNULL(oSongjangChgLog.FItemList(i).Fcomment) or (oSongjangChgLog.FItemList(i).Fcomment="") then %>
            <a href="#" onClick="popJcomment('<%=oSongjangChgLog.FItemList(i).FOrderserial%>','<%=oSongjangChgLog.FItemList(i).Fitemid%>','<%=oSongjangChgLog.FItemList(i).Fitemoption%>',true);return false;"><img src="/images/icon_new.gif" alt="�ڸ�Ʈ�ۼ�"></a>
        <% else %>
            <a href="#" onClick="popJcomment('<%=oSongjangChgLog.FItemList(i).FOrderserial%>','<%=oSongjangChgLog.FItemList(i).Fitemid%>','<%=oSongjangChgLog.FItemList(i).Fitemoption%>',false);return false;"><%=oSongjangChgLog.FItemList(i).Fcomment%></a>
        <% end if %>

        <% if oSongjangChgLog.FItemList(i).Fcancelyn<>"N" then %><br><strong>�ֹ����<strong><% end if %>
        <% if oSongjangChgLog.FItemList(i).Fdcancelyn="Y" then %><br><strong>��ǰ���<strong><% end if %>
        <% if oSongjangChgLog.FItemList(i).Fdcancelyn="A" then %><br><strong>��ǰ�߰�<strong><% end if %>
    </td>
</tr>
<% next %>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
    <% if (FALSE) then %>
		<% if oSongjangChgLog.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oSongjangChgLog.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oSongjangChgLog.StartScrollPage to oSongjangChgLog.FScrollCount + oSongjangChgLog.StartScrollPage - 1 %>
			<% if i>oSongjangChgLog.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oSongjangChgLog.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    <% end if %>
	</td>
</tr>

</table>

<p>
<form name="frmcmt" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="addcmt">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="addcomment" value="">
</form>

<%
set oSongjangChgLog = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->

