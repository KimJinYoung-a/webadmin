<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idxArr,iSall
idxArr = Replace(request.Form("idxArr"), " ", "")
idxArr = Trim(idxArr)
iSall   =  requestCheckVar(request("isall"), 32)

if (Right(idxArr,1)=",") then idxArr=Left(idxArr,Len(idxArr)-1)

if (Len(idxArr)<1) and (iSall="") then
    response.write "<script>alert('���õ� �ֹ����� �����ϴ�.');</script>"
    dbget.close()	:	response.End
end if
%>
<script language='javascript'>
function popDeliverCode(){
    var popwin = window.open('popDeliverCode.asp','popDeliverCode','width=400,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function downloadDeliveXL(){
    var xlfrm = document.xlfrm;
	xlfrm.target="iiframeXL";
	xlfrm.action="upchecs_songjanglistexcel.asp";
	xlfrm.submit();
}

function NextStep(frm){
    if (frm.songjangfile.value.length<1){
        alert('���ε��� CSV������ �����ϼ���.')
        return;
    }

    if (confirm('���� �ܰ�� ���� �Ͻðڽ��ϱ�?')){
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
<p>
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

<p>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="50">
    <td width="100" align="center">Step 2</td>
    <td>������ ���� ���� ��ȣ�� �Է� �Ͻ� �� ���� PC�� CSV ���Ϸ� �����ϼ���.</td>
</tr>
</table>
<p>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmNext" method="post" action="upchecs_pop_BatchSongjangInputStep2.asp" onsubmit="return false;" enctype="multipart/form-data">
<tr bgcolor="#FFFFFF" height="50">
    <td width="100" align="center">Step 3</td>
    <td>������ CSV ������ �����Ͻ� �� �����ܰ�� �̵��ϼ���.
        <input type="file" name="songjangfile" size="30" value="">
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
    <td colspan="2" align="center">
    <input type="button" value="�����ܰ�� ����" onClick="NextStep(frmNext)">
    </td>
</tr>
</form>
</table>
<iframe name="iiframeXL" name="iiframeXL" width="110" height="110" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>
<form name=xlfrm method=post action="">
<input type="hidden" name="idxArr" value="<%= idxArr %>">
<input type="hidden" name="iSall" value="<%= iSall %>">
</form>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
