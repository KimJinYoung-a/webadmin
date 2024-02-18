<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �������� ���� �ı� ����
' History : 2019.08.13 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/isms/personaldata_cls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim i, idxarr, userid, menupos
    menupos = requestcheckvar(request("menupos"),10)

userid = session("ssBctId")

if request.form("idx").count<1 then
    response.write "<script type='text/javascript'>"
    response.write "    alert('�����Ұ��� �����ϴ�.');"
    response.write "    self.close();"
    response.write "</script>"
    dbget.close() : response.end
end if

for i=1 to request.form("idx").count
    idxarr = idxarr & request.form("idx")(i) & ","
next

%>
<script type='text/javascript'>

function downFileconfirm(){
    if ( !frmupdate.ck1.checked ){
        alert('�����м⿡ ���� üũ�� ���ּ���.');
        frmupdate.ck1.focus();
        return;
    }
    if ( !frmupdate.ck2.checked ){
        alert('CD,USB��⿡ ���� üũ�� ���ּ���.');
        frmupdate.ck2.focus();
        return;
    }
    if ( !frmupdate.ck3.checked ){
        alert('��ǻ�� ���� ������ ���� üũ�� ���ּ���.');
        frmupdate.ck3.focus();
        return;
    }
    if ( !frmupdate.ck4.checked ){
        alert('��Ÿ����(���ϵ�,�̸��� ��) ��ü �ı�/������ ���� üũ�� ���ּ���.');
        frmupdate.ck4.focus();
        return;
    }

    frmupdate.mode.value = "downFileconfirmArr";
    frmupdate.target="_self"
    frmupdate.action="/admin/isms/personaldata_process.asp";
    frmupdate.submit();
}

// ��ü����
function totalCheck(tmpval){
    if (tmpval.checked){
        frmupdate.ck1.checked = true
        frmupdate.ck2.checked = true
        frmupdate.ck3.checked = true
        frmupdate.ck4.checked = true
    }else{
        frmupdate.ck1.checked = false
        frmupdate.ck2.checked = false
        frmupdate.ck3.checked = false
        frmupdate.ck4.checked = false
    }
}

</script>
</head>
<body>
<form name="frmupdate" method="post" action="/admin/isms/personaldata_process.asp" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idxarr" value="<%= idxarr %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td><h2>������ �ı� Ȯ�μ�</h2></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>
        1. <%= session("ssBctCname") %>�� �ٹ������� ���Ͽ� ��ǰ�Ǹ�, ���, ��ǰ ���� ���� ������ �����ϴµ� �־� ����� ���� ������ �ı�� �����Ͽ�
        <br>&nbsp;&nbsp;&nbsp;������ ������ Ȯ���Ͽ� �ֽñ� �ٶ��ϴ�.
        <Br>2. <%= session("ssBctCname") %>�� 2���� �̻� �������� �ı�� Ȯ�μ��� ���ۼ���, ���� ��뿡 ������ ������ �˷� �帳�ϴ�.
        <br>3. �� ������ȣ �� ������ �ŷ����� Ȯ���� ���Ͽ� ����� �ֽñ� �ٶ��ϴ�.
    </td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>
        <br><%= session("ssBctCname") %>�� �ٹ������� ���Ͽ� �Ǹ��� ��ǰ�� ������ ��� ������ ���� ���� ������ � ���Ͽ� ������ ���� �ı� ������ �����Ͽ����� Ȯ�� �մϴ�.
    </td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>
        1. �ı� ���� : ��� ������ ���� ���� �� ����(����, ��ȭ��ȣ, �ּ� ��)
        <br>2. �ı� ���
        <br>&nbsp;&nbsp;&nbsp;
        &nbsp;&nbsp;<input type="checkbox" name="ck1"> �����м�
        &nbsp;&nbsp;<input type="checkbox" name="ck2"> CD,USB���
        &nbsp;&nbsp;<input type="checkbox" name="ck3"> ��ǻ�� ���� ����
        &nbsp;&nbsp;<input type="checkbox" name="ck4"> ��Ÿ����(���ϵ�,�̸��� ��) ��ü �ı�/����
    </td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>
        <br><%= session("ssBctCname") %>�� �������� ���� ������ �ؼ��� ���̸�, ��� ������ �ı� ������ ������ �Ǵ� �ҿ��� �������� �����ϴ� ������,
        ������� å���� �� �δ��� ���� Ȯ�� �մϴ�.
        <br><br><p align="center"><%= year(date()) %>�� <%= month(date()) %>�� <%= day(date()) %>�� <%= session("ssBctCname") %></p>
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
    <td>
        <input type="checkbox" name="ckall" onClick="totalCheck(this);">���Ȯ��
        &nbsp;
        <input type="button" value="Ȯ��"  onclick="downFileconfirm();" class="button" />
    </td>
</tr>
</table>
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->