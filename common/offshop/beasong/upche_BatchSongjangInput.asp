<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim detailidxArr,iSall
	detailidxArr = request.Form("detailidxArr")
	detailidxArr = Trim(detailidxArr)
	iSall   =  request("isall")

if (Right(detailidxArr,1)=",") then detailidxArr=Left(detailidxArr,Len(detailidxArr)-1)

if (Len(detailidxArr)<1) and (iSall="") then
    response.write "<script>alert('���õ� �ֹ����� �����ϴ�.');</script>"
    dbget.close()	:	response.End
end if
%>

<script language='javascript'>

function popDeliverCode(){
    var popwin = window.open('/designer/jumunmaster/popDeliverCode.asp','popDeliverCode','width=400,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function downloadDeliveXL(){
    var xlfrm = document.xlfrm;
	xlfrm.target="iiframeXL";
	xlfrm.action="/common/offshop/beasong/upche_dosongjanglistexcel.asp";
	xlfrm.submit();
}

function NextStep(frm){
    if (frm.songjangfile.value.length<1){
        alert('���ε��� CSV������ �����ϼ���.')
        return;
    }

    if (confirm('���� �ܰ�� ���� �Ͻðڽ��ϱ�?')){
        frm.action="/common/offshop/beasong/upche_BatchSongjangInputStep2.asp";
        frm.submit();
    }
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td>
        1. ���� ����� �ٿ�޾� <strong>�ù�� �ڵ�</strong>�� <strong>���� ��ȣ</strong>�� �Է� �Ͻ� �� ���ε� �Ͻø� �ϰ� �߼� ó�� �˴ϴ�.<br>
        2. �����ȣ �Է� �� ������ <strong>CSV �������� ����</strong> �Ͻñ� �ٶ��ϴ�.<br>
        3. ��ü ���� �������� <strong>�⺻ �ù�縦 ����</strong> �� �����ø� �ù�� �ڵ尡 �⺻���� �����Ǿ� �ٿ�ε� �� �� �ֽ��ϴ�.<br>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="right">
    <a href="javascript:popDeliverCode();"><font color="blue">[�ù�� �ڵ� ����]</font></a>
    </td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="50">
    <td width="100" align="center">Step 1</td>
    <td>���� ����� �ٿ� ��������. (����� ������ �ۼ��˴ϴ�.)<a href="javascript:downloadDeliveXL();"><font color="blue">[�ٿ�ε�]</font></a>
        <br>������ĳ��� �ù���ڵ�� �⺻�ù��� �����ǿ���, �����Ͻñ� �ٶ��ϴ�.
        <br>�⺻�ù��� ��ü���� �������� ���������մϴ�.
    </td>
</tr>
</table>

<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="50">
    <td width="100" align="center">Step 2</td>
    <td>������ ���� ���� ��ȣ�� �Է� �Ͻ� �� ���� PC�� CSV ���Ϸ� �����ϼ���.</td>
</tr>
</table>

<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmNext" method="post" onsubmit="return false;" enctype="multipart/form-data">
<tr bgcolor="#FFFFFF" height="50">
    <td width="100" align="center">Step 3</td>
    <td>������ CSV ������ �����Ͻ� �� �����ܰ�� �̵��ϼ���.
        <input type="file" name="songjangfile" size="30" value="" class="file">
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
    <td colspan="2" align="center">
    <input type="button" value="�����ܰ�� ����" onClick="NextStep(frmNext)" class="button">
    </td>
</tr>
</form>
</table>

<iframe name="iiframeXL" name="iiframeXL" width="0" height="0" frameborder=0 scrolling="auto"></iframe>
<form name="xlfrm" method="post">
	<input type="hidden" name="detailidxArr" value="<%= detailidxArr %>">
	<input type="hidden" name="iSall" value="<%= iSall %>">
</form>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
