<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����뿪����
' Hieditor : 2021.04.14 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/logistics/invoice_band_cls.asp"-->
<%
Dim i, page, osongjang, isusing, reload, siteseq, gubuncd, osongjanglog, songjangdiv
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
    isusing = requestcheckvar(request("isusing"),10)
    siteseq = requestcheckvar(getNumeric(request("siteseq")),10)
    reload = requestcheckvar(request("reload"),2)
    gubuncd = requestcheckvar(trim(request("gubuncd")),3)
    songjangdiv = requestcheckvar(trim(request("songjangdiv")),32)

if page = "" then page = 1
if reload="" and isusing="" then isusing="Y"
if siteseq="" and siteseq="" then siteseq="10"
''if gubuncd="" and gubuncd="" then gubuncd="00"
if reload="" and songjangdiv="" then
    'songjangdiv = "1"   ' �����ù�
    songjangdiv = "2"   ' �Ե��ù�
end if

set osongjang = new cinvoice_band_list
	osongjang.FPageSize = 50
	osongjang.FCurrPage = page
    osongjang.Frectisusing = isusing
    osongjang.Frectsiteseq = siteseq
    osongjang.Frectgubuncd = gubuncd
    osongjang.FRectSongjangDiv = songjangdiv

    osongjang.finvoice_band()

set osongjanglog = new cinvoice_band_list
	osongjanglog.FPageSize = 5
	osongjanglog.FCurrPage = 1
    osongjanglog.Frectsiteseq = siteseq
    osongjanglog.Frectgubuncd = gubuncd
    osongjanglog.FRectSongjangDiv = songjangdiv

	osongjanglog.finvoice_band_log()
%>

<script type="text/javascript">

function invoice_band_reg(iidx){
	var popwin = window.open('/admin/logics/invoice_band_reg.asp?iidx='+iidx+'&menupos=<%=menupos%>','addreg','width=1200,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function NextPage(page){
	document.frm.page.value= page;
	document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method=get action="" style="margin:0px;">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reload" value="on">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
    <td align="left">
        * ��ü : <% drawSelectBoxSiteSeq "siteseq",siteseq,"" %>
        &nbsp;
        * �ù�� :
        <% Call drawSelectBoxDeliverCompany ("songjangdiv", songjangdiv) %>
        &nbsp;
        * ����� : <% drawSelectBoxgubuncd "gubuncd",gubuncd,"" %>
        &nbsp;
        * ��뿩�� : <% drawSelectBoxisusingYN "isusing",isusing,"" %>
    </td>
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="�˻�" onClick="NextPage('');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left"></td>
    <td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
    <td colspan="3">
        �� �ֱ� ������ ���� ���峻�� �α� 5��
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>�����ȣ(����Ű����)</td>
    <td>���������ȣ</td>
    <td>�ֹ���ȣ</td>
</tr>
<% if osongjanglog.FresultCount>0 then %>
<% for i=0 to osongjanglog.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" align="center">
    <td>
        <%= osongjanglog.FItemList(i).fSONGJANGNO %>
    </td>
    <td>
        <%= osongjanglog.FItemList(i).fREALSONGJANGNO %>
    </td>
    <td>
        <%= osongjanglog.FItemList(i).fORDERSERIAL %>
    </td>
</tr>
<% next %>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
</table>
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
    </td>
    <td align="right">
        <input type="button" class="button" value="�űԵ��" onclick="invoice_band_reg('');">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
    <td colspan="15">
        �˻���� : <b><%= osongjang.FTotalCount %></b>
        &nbsp;
        ������ : <b><%= page %>/ <%= osongjang.FTotalPage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>��ȣ</td>
    <td>��ü</td>
    <td>�ù��</td>
    <td>�����</td>
    <td>�����ȣ(����Ű����)</td>
    <td>���������ȣ</td>
    <td>
        ���������
        <br>(8�ð��ֱ������Ʈ)
    </td>
    <td>
        �⺻���忩��
        <br>(����������������뿪)
    </td>
    <td>��뿩��</td>
    <td>���ʵ��</td>
    <td>��������</td>
    <td>���</td>
</tr>
<% if osongjang.FresultCount>0 then %>
<% for i=0 to osongjang.FresultCount-1 %>
<% if osongjang.FItemList(i).fbasicsongjangyn = "Y" then %>
<tr align="center" bgcolor="#FFFFaa" align="center">
<% else %>
<tr align="center" bgcolor="#FFFFFF" align="center">
<% end if %>
    <td>
        <%= osongjang.FItemList(i).fiidx %>
    </td>
    <td>
        <%= getSiteSeqnamestr(osongjang.FItemList(i).fsiteseq) %>
    </td>
    <td>
        <%= osongjang.FItemList(i).Fdivname %>
    </td>
    <td>
        <%= getgubuncdname(osongjang.FItemList(i).fgubuncd) %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fstartsongjangno %> - <%= osongjang.FItemList(i).fendsongjangno %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fstartrealsongjangno %> - <%= osongjang.FItemList(i).fendrealsongjangno %>
    </td>
    <td>
        <% if osongjang.FItemList(i).fbasicsongjangyn="Y" then %>
            <% if osongjang.FItemList(i).fremainsongjangcount="0" then %>
                <%= osongjang.FItemList(i).fendrealsongjangno-osongjang.FItemList(i).fstartrealsongjangno %>
            <% else %>
                <%= osongjang.FItemList(i).fremainsongjangcount %>
            <% end if %>
        <% else %>
            <% 'if osongjang.fcurrentbasicsongjangidx > osongjang.FItemList(i).fiidx then %>
                <%= osongjang.FItemList(i).fremainsongjangcount %>
            <% 'else %>
                <%'= osongjang.FItemList(i).fendrealsongjangno-osongjang.FItemList(i).fstartrealsongjangno %>
            <% 'end if %>
        <% end if %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fbasicsongjangyn %>
    </td>
    <td>
        <%= osongjang.FItemList(i).fisusing %>
    </td>
    <td>
        <%= osongjang.FItemList(i).freguserid %>
        <br><%= left(osongjang.FItemList(i).fregdate,10) %>
        <br><%= mid(osongjang.FItemList(i).fregdate,12,22) %>
    </td>
    <td>
        <%= osongjang.FItemList(i).flastuserid %>
        <br><%= left(osongjang.FItemList(i).flastupdate,10) %>
        <br><%= mid(osongjang.FItemList(i).flastupdate,12,22) %>
    </td>
    <td>
        <input type="button" class="button" value="����" onclick="invoice_band_reg('<%= osongjang.FItemList(i).fiidx %>');">
    </td>
</tr>
<% next %>
<tr bgcolor="FFFFFF">
    <td colspan="15" align="center">
        <% if osongjang.HasPreScroll then %>
        <a href="javascript:NextPage('<%= osongjang.StartScrollPage-1 %>')">[pre]</a>
        <% else %>
            [pre]
        <% end if %>

        <% for i=0 + osongjang.StartScrollPage to osongjang.FScrollCount + osongjang.StartScrollPage - 1 %>
            <% if i>osongjang.FTotalpage then Exit for %>
            <% if CStr(page)=CStr(i) then %>
            <font color="red">[<%= i %>]</font>
            <% else %>
            <a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
            <% end if %>
        <% next %>

        <% if osongjang.HasNextScroll then %>
            <a href="javascript:NextPage('<%= i %>')">[next]</a>
        <% else %>
            [next]
        <% end if %>
    </td>
</tr>

<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>

</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
