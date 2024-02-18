<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����뿪����
' Hieditor : 2021.04.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/admin/incsessionadmin.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/logistics/invoice_band_cls.asp"-->

<%
dim iidx,siteseq,gubuncd,startsongjangno,endsongjangno,startrealsongjangno,endrealsongjangno
dim remainsongjangcount,basicsongjangyn,isusing,regdate,lastupdate,reguserid,lastuserid, songjangdiv
dim osongjangedit, i, mode, menupos
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	iidx = requestcheckvar(getNumeric(request("iidx")),10)

mode="add"

set osongjangedit = new cinvoice_band_list
	osongjangedit.frectiidx = iidx

	if iidx <> "" then
		osongjangedit.finvoice_band_one()

        if osongjangedit.FResultCount>0 then
            mode="edit"
            iidx = osongjangedit.FOneItem.fiidx
            siteseq = osongjangedit.FOneItem.fsiteseq
            gubuncd = osongjangedit.FOneItem.fgubuncd
            startsongjangno = osongjangedit.FOneItem.fstartsongjangno
            endsongjangno = osongjangedit.FOneItem.fendsongjangno
            startrealsongjangno = osongjangedit.FOneItem.fstartrealsongjangno
            endrealsongjangno = osongjangedit.FOneItem.fendrealsongjangno
            remainsongjangcount = osongjangedit.FOneItem.fremainsongjangcount
            basicsongjangyn = osongjangedit.FOneItem.fbasicsongjangyn
            isusing = osongjangedit.FOneItem.fisusing
            regdate = osongjangedit.FOneItem.fregdate
            lastupdate = osongjangedit.FOneItem.flastupdate
            reguserid = osongjangedit.FOneItem.freguserid
            lastuserid = osongjangedit.FOneItem.flastuserid
            songjangdiv = osongjangedit.FOneItem.Fsongjangdiv
        end if
	end if
set osongjangedit = nothing

if remainsongjangcount="" or isnull(remainsongjangcount) then remainsongjangcount=0
if basicsongjangyn="" or isnull(basicsongjangyn) then basicsongjangyn="N"
if isusing="" or isnull(isusing) then isusing="Y"
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

	function invoice_band_reg(){
		if ($('#frm select[name="siteseq"] option:selected').val()==''){
			alert('��ü�� �����ϼ���.');
			$('#frm select[name="siteseq"]').focus();
			return;
		}
		if ($('#frm select[name="gubuncd"] option:selected').val()==''){
			alert('������� �����ϼ���.');
			$('#frm select[name="gubuncd"]').focus();
			return;
		}
		if ($('#frm input[name="startsongjangno"]').val()==''){
			alert('���ۼ����ȣ(����Ű����)�� �Է��ϼ���.');
			$('#frm input[name="startsongjangno"]').focus();
			return;
		}
		if ($('#frm input[name="endsongjangno"]').val()==''){
			alert('��������ȣ(����Ű����)�� �Է��ϼ���.');
			$('#frm input[name="endsongjangno"]').focus();
			return;
		}
		if ($('#frm input[name="startrealsongjangno"]').val()==''){
			alert('���۽��������ȣ�� �Է��ϼ���.');
			$('#frm input[name="startrealsongjangno"]').focus();
			return;
		}
		if ($('#frm input[name="endrealsongjangno"]').val()==''){
			alert('������������ȣ�� �Է��ϼ���.');
			$('#frm input[name="endrealsongjangno"]').focus();
			return;
		}
		if ($('#frm select[name="basicsongjangyn"] option:selected').val()==''){
			alert('�⺻���忩�θ� �����ϼ���.');
			$('#frm select[name="basicsongjangyn"]').focus();
			return;
		}
		if ($('#frm select[name="isusing"] option:selected').val()==''){
			alert('��뿩�θ� �����ϼ���.');
			$('#frm select[name="isusing"]').focus();
			return;
		}
		frm.submit();
	}

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left"></td>
    <td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frm" id="frm" method="post" action="/admin/logics/invoice_band_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<% if mode="edit" then %>
<tr bgcolor="#FFFFFF">
    <td align="center">��ȣ</td>
    <td>
        <%= iidx %>
		<input type="hidden" name="iidx" value="<%= iidx %>">
    </td>
</tr>
<% else %>
    <input type="hidden" name="iidx" value="<%= iidx %>">
<% end if %>
<tr bgcolor="#FFFFFF">
    <td align="center">��ü</td>
    <td>
        <select class="select" name="siteseq" >
            <option value="10" selected>�ٹ�����</option>
        </select>
        <!-- ��ü�� �߰��Ϸ��� �¶������, ��Ÿ��� ��� �����Ŀ� ��ü�� �߰��ؾ� �Ѵ�.
        <% drawSelectBoxSiteSeq "siteseq",siteseq,"" %>
        -->
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">�ù��</td>
    <td>
        <% Call drawSelectBoxDeliverCompany ("songjangdiv", songjangdiv) %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">�����</td>
    <td>
        <% drawSelectBoxgubuncd "gubuncd",gubuncd,"" %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">�����ȣ(����Ű����)</td>
    <td>
        <input type="text" name="startsongjangno" value="<%= startsongjangno %>" size=11 maxlength=12>
		- <input type="text" name="endsongjangno" value="<%= endsongjangno %>" size=11 maxlength=12>
		<br>�ǳ��� 1�ڸ��� ����Ű �Դϴ�. �����ȣ�� ���� ��� �Է��� �ּž� �մϴ�.
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">���������ȣ</td>
    <td>
        <input type="text" name="startrealsongjangno" value="<%= startrealsongjangno %>" size=11 maxlength=12>
		- <input type="text" name="endrealsongjangno" value="<%= endrealsongjangno %>" size=11 maxlength=12>
		<br>���忡 ����Ǵ� ���������ȣ �Դϴ�. �ǳ��� 1�ڸ� ����Ű�� �����ϰ� �Է��� �ּ���.
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">�⺻���忩��</td>
    <td>
		<% if mode="edit" and basicsongjangyn="Y" then %>
			<input type="hidden" name="basicsongjangyn" value="<%= basicsongjangyn %>">
			<%= basicsongjangyn %>
			<br>�⺻���忩�θ� N���� ������ �Ұ��մϴ�.<br>�⺻������ 1�� �̻� ���� �ؾ� �մϴ�.<br>����Ͻ� ����뿪���� �⺻���忩�θ� Y �� ���ּ���.
		<% else %>
			<% drawSelectBoxisusingYN "basicsongjangyn",basicsongjangyn,"" %>
			<br>Y �� ��� �ش� ����뿪���� ���������� ��� �˴ϴ�.���� ������ ���� ���뿪
		<% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">��뿩��</td>
    <td>
		<% drawSelectBoxisusingYN "isusing",isusing,"" %>
    </td>
</tr>
<% if mode="edit" then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">���������</td>
		<td>
			<%= remainsongjangcount %>
			<br>8�ð� �ֱ�� ������Ʈ �˴ϴ�.
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">���ʵ��</td>
		<td>
			<%= reguserid %>
			<br><%= regdate %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">��������</td>
		<td>
			<%= lastuserid %>
			<br><%= lastupdate %>
		</td>
	</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="2">
		<input type="button" value="����" onclick="invoice_band_reg();" class="button">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
