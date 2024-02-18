<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2019.10.16 �ѿ�� ����
'	Description : Link �߼�
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/rndSerial.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/LinkSendCls.asp"-->
<%
dim linkidx, title, linkurl, isusing, viewcount, regdate, lastupdate, lastadminid, menupos, oLink
	linkidx = requestCheckVar(getNumeric(request("linkidx")),10)
    menupos = requestCheckVar(getNumeric(request("menupos")),10)

set oLink = New CLinkSend
    oLink.FRectlinkidx = linkidx

	if linkidx<>"" then
    	oLink.GetLinkSend_one
	end if

if oLink.FTotalCount > 0 then
    title		= oLink.FOneItem.ftitle
    linkurl		= oLink.FOneItem.flinkurl
    isusing		= oLink.FOneItem.fisusing
    viewcount	= oLink.FOneItem.fviewcount
    regdate		= oLink.FOneItem.fregdate
    lastupdate		= oLink.FOneItem.flastupdate
    lastadminid		= oLink.FOneItem.flastadminid
end if
set oLink = Nothing
%>
<script type='text/javascript'>

function jsRegLink(){
	var frm = document.frmReg;
	if(!frm.title.value) {
		alert("��ũ���� �Է��� �ּ���");
		frm.title.focus();
		return;
	}
	if(!frm.linkurl.value) {
		alert("������ũ �Է��� �ּ���");
		frm.linkurl.focus();
		return;
	}
	if(!frm.isusing.value) {
		alert("��뿩�θ� ������ �ּ���");
		frm.isusing.focus();
		return;
	}

	if(confirm("�����Ͻðڽ��ϱ�?")) {
		frm.action="/admin/sitemaster/link/LinkSend_process.asp"
		frm.mode.value="RegLink";
		frm.target="view";
		frm.submit();
	}
}

function fnLinkURLCopy(link) {
	copyStringToClipboard(link);
	alert('��ũ�� ����Ǿ����ϴ�.\n���Ͻô� ���� Ctrl+V �Ͻø�˴ϴ�.');
}

</script>

<form name="frmReg" method="post" action="" style="margin:0px;">
<input type="hidden" name="linkidx" value="<%= linkidx %>">
<input type="hidden" name="mode" value="">
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2">
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<tr>			
			<td><b>������ ��ũ ���</b></td>
		</tr>
		<tr>
			<td>	
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<% IF linkidx <> "" THEN %>
                    <tr>
                        <td bgcolor="#EFEFEF" width="100" align="center">��ũ��ȣ</td>
                        <td bgcolor="#FFFFFF"><%= linkidx %></td>
                    </tr>
				<% END IF %>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">��ũ��</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="100" maxlength="128" name="title" value="<%= ReplaceBracket(title) %>">
					</td>
				</tr>
				<% IF linkidx <> "" THEN %>
					<tr>
						<td bgcolor="#EFEFEF" width="100" align="center">�ܺγ��⸵ũ</td>
						<td bgcolor="#FFFFFF">
							http://www.10x10.co.kr/apps/Link/LinkSend.asp?key=<%= rdmSerialEnc(linkidx) %>
							<input type="button" value="��ũ����" onclick="fnLinkURLCopy('http://www.10x10.co.kr/apps/Link/LinkSend.asp?key=<%= rdmSerialEnc(linkidx) %>')" class="button">
						</td>
					</tr>
				<% END IF %>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">������ũ</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="140" maxlength="512" name="linkurl" value="<%= ReplaceBracket(linkurl) %>">
						<br>ex) http://www.10x10.co.kr
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF" width="100" align="center">��뿩��</td>
					<td bgcolor="#FFFFFF">
						<% drawSelectBoxisusingYN "isusing",isusing,"" %>
					</td>
				</tr>
                <% IF linkidx <> "" THEN %>
                    <tr>
                        <td bgcolor="#EFEFEF" width="100" align="center">Ŭ����</td>
                        <td bgcolor="#FFFFFF"><%= viewcount %></td>
                    </tr>
                    <tr>
                        <td bgcolor="#EFEFEF" width="100" align="center">�����</td>
                        <td bgcolor="#FFFFFF"><%= regdate %></td>
                    </tr>
                    <tr>
                        <td bgcolor="#EFEFEF" width="100" align="center">��������</td>
                        <td bgcolor="#FFFFFF">
                            <%= left(lastupdate,10) %>
                            <br>
                            <%= mid(lastupdate,11,22) %>
                            <% if lastadminid <> "" then %>
                                <br>(<%= lastadminid %>)
                            <% end if %>
                        </td>
                    </tr>
				<% END IF %>
				</table>		
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td align="center">
                        <input type="button" value="����" onclick="jsRegLink();" class="button">
                    </td>
				</tr>
				</table>
			</td>
		</tr>	
		</table>
	</td>
</tr>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 allowtransparency="true"  frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height=0 allowtransparency="true"  frameborder="0" scrolling="no"></iframe>
<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->