<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
'###############################################
' PageName : main_md_recommand_flash_writes.asp
' Discription : ��ǰ�ڵ� �ϰ� ���
'###############################################
'// ���� ����
Dim realdate : realdate = request("realdate")
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
// ���˻�
function SaveForm(frm) {
	var selChk=true;
	if(frm.linkitemid.value=="") {
		alert("�ϰ� ����Ͻ� ��ǰ�ڵ带 �Է����ּ���");
		frm.linkitemid.focus();
		return;
	}

	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}
</script>
<form name="frmSub" method="post" action="main_md_recommend_flash_proc.asp" style="margin:0px;">
    <table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
        <tr bgcolor="#FFFFFF">
            <td height="25" colspan="4" bgcolor="#F8F8F8"><b>���� ���� - ��ǰ�ڵ� �ϰ� ���</b></td>
        </tr>
        <colgroup>
            <col width="100" />
            <col width="*" />
            <col width="100" />
            <col width="*" />
        </colgroup>
        <tr bgcolor="#FFFFFF">
            <td bgcolor="#DDDDFF">��ǰ�ڵ�</td>
            <td colspan="3">
                <textarea name="linkitemid" class="textarea" title="��ǰ�ڵ�" style="width:95%; height:80px;"></textarea>
                <p>�� ��ǰ�ڵ带 ��ǥ(,) �Ǵ� ���ͷ� �����Ͽ� �Է�</p>
                <p>�� ��ǰ���� �⺻ ��ǰ������ �Է� �˴ϴ�. (���� �ʿ�)</p>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
            <td colspan="3">
                <input id="startdate" name="startdate" value="<%=Left(realdate,10)%>" class="text" size="10" maxlength="10" />
                <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
                <input type="text" name="startdatetime" size="2" maxlength="2" value="00" />(�� 00~23)
                <input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
                <script type="text/javascript">
                var CAL_Start = new Calendar({
                    inputField : "startdate",
                    trigger    : "startdate_trigger",
                    onSelect: function() {
                        var date = Calendar.intToDate(this.selection.get());
                        CAL_End.args.min = date;
                        CAL_End.redraw();
                        this.hide();
                    },
                    bottomBar: true,
                    dateFormat: "%Y-%m-%d"
                });
                </script>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
            <td colspan="3">
                <input id="enddate" name="enddate" value="<%=Left(realdate,10)%>" class="text" size="10" maxlength="10" />
                <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
                <input type="text" name="enddatetime" size="2" maxlength="2" value="23">(�� 00~23)
                <input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
                <script type="text/javascript">
                var CAL_End = new Calendar({
                    inputField : "enddate",
                    trigger    : "enddate_trigger",
                    onSelect: function() {
                        var date = Calendar.intToDate(this.selection.get());
                        CAL_Start.args.max = date;
                        CAL_Start.redraw();
                        this.hide();
                    },
                    bottomBar: true,
                    dateFormat: "%Y-%m-%d"
                });
                </script>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td bgcolor="#DDDDFF">���ü���</td>
            <td>
                <input type="text" name="disporder" class="text" size="4" value="99" />
            </td>
            <td bgcolor="#DDDDFF">��뿩��</td>
            <td>
                <span id="rdoUsing">
                <input type="radio" name="isusing" id="rdoUsing1" value="Y" checked /><label for="rdoUsing1">���</label>
                <input type="radio" name="isusing" id="rdoUsing2" value="N" /><label for="rdoUsing2">����</label>
                </span>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td colspan="4" align="center"><input type="button" value=" �� �� " onClick="SaveForm(this.form);"></td>
        </tr>
    </table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->