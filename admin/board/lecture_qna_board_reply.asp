<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/myqnacls.asp" -->
<%

dim i, j
dim reffrom
reffrom = request("reffrom")

'==============================================================================
'���� 1:1�����亯
dim boardqna
set boardqna = New CMyQNA

boardqna.read(request("id"))

if boardqna.results(0).userid <> "" then
dim orderinfo
set orderinfo = New CMyQNAOrderInfo
orderinfo.UserOrderInfo (boardqna.results(0).userid)
orderinfo.UserMinusOrderInfo (boardqna.results(0).userid)
end if

if boardqna.results(0).userid <> "" or boardqna.results(0).orderserial <> "" then
dim myqnalist
set myqnalist = New CMyQNA
if boardqna.results(0).userid <> "" then
myqnalist.SearchUserID = boardqna.results(0).userid
end if
if boardqna.results(0).orderserial <> "" then
myqnalist.SearchOrderSerial = boardqna.results(0).orderserial
end if
myqnalist.PageSize = 20
myqnalist.CurrPage = 1
myqnalist.list
end if
%>

<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
����Ÿ - ���ֹ�������<br><br>
<script>
function SubmitForm()
{
        if (document.f.replytitle.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }
        if (document.f.replycontents.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }

        if (confirm("�Է��� ��Ȯ�մϱ�?") == true) { document.f.submit(); }
}

function updateqadiv(){
	if (confirm("�����Ͻðڽ��ϱ�?")){
		document.updateform.submit();
	}
}

function delqadiv(){
	if (confirm("�����Ͻðڽ��ϱ�?")){
		document.delform.submit();
	}
}
</script>
<table width="580" border="0" align="center" cellpadding="0" cellspacing="0">
<form method=post name="updateform" action="doeditqadiv.asp">
<input type="hidden" name="id" value="<% = boardqna.results(0).id %>">
<tr>
	<td align="right">������������ :
		  <select name="qadiv">
			<option>����</option>
			<option value="00" <% if boardqna.results(0).qadiv = "00" then response.write "selected" %>>��۹���</option>
			<option value="01" <% if boardqna.results(0).qadiv = "01" then response.write "selected" %>>�ֹ�����</option>
			<option value="02" <% if boardqna.results(0).qadiv = "02" then response.write "selected" %>>��ǰ����</option>
			<option value="03" <% if boardqna.results(0).qadiv = "03" then response.write "selected" %>>�����</option>
			<option value="04" <% if boardqna.results(0).qadiv = "04" then response.write "selected" %>>���,ȯ�ҹ���</option>
			<option value="06" <% if boardqna.results(0).qadiv = "06" then response.write "selected" %>>��ȯ����</option>
			<option value="08" <% if boardqna.results(0).qadiv = "08" then response.write "selected" %>>����ǰ����</option>
			<option value="10" <% if boardqna.results(0).qadiv = "10" then response.write "selected" %>>�ý��۹���</option>
			<option value="12" <% if boardqna.results(0).qadiv = "12" then response.write "selected" %>>������������</option>
			<option value="20" <% if boardqna.results(0).qadiv = "20" then response.write "selected" %>>��Ÿ����</option>
		  </select>
		  &nbsp; <a href="http://www.10x10.co.kr/shopping/category.asp?itemid=<%= boardqna.results(i).FItemID %>" target="_blank" >��ǰ��ȣ</a> : <input type="text" name="itemid" size="10" value="<% = boardqna.results(0).itemid %>">
		  <!--
		  <input type="hidden" name="itemid" value="0">
		  -->
		  &nbsp;<input type="button" value="����" onclick="updateqadiv();">&nbsp;
		  <% if reffrom = "itemqa" then %>
		  <input type="button" value="����" onclick="delqadiv();">&nbsp;
		  <% end if %>
		  </td>
</tr>
</form>
</table>

<form method="post" name="delform" action="cscenter_qna_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="mode" value="del">
</form>

<table width="580" border="0" align="center" cellpadding="0" cellspacing="3">
<form method="post" name="f" action="cscenter_qna_board_act.asp" onsubmit="return false">
<% if boardqna.results(0).replyuser<>"" then %>
<input type="hidden" name="mode" value="reply">
<% else %>
<input type="hidden" name="mode" value="firstreply">
<% end if %>

<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="email" value="<%= boardqna.results(0).usermail %>">
<input type="hidden" name="emailok" value="<%= boardqna.results(0).emailok %>">
<input type="hidden" name="extsitename" value="<%= boardqna.results(0).Fextsitename %>">

	<tr>
      <td background="/admin/images/topbar_bg.gif" height="26" valign="middle" colspan="2">
        <div align="left">
          <table width="520" border="0" cellpadding="0" cellspacing="0" class="a">
            <tr>
              <td>
                <div align="left">&nbsp;<span class="a"><b>�� <%= boardqna.results(0).title %></b></span></div>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
    <tr>
      <td class="a" height="5">&nbsp;<b>�ۼ���</b> : <%= boardqna.results(0).username %>(<%= boardqna.results(0).userid %>/<%= boardqna.results(0).orderserial %>)</td>
	  <td align="right" class="a"><b>�ۼ���</b> : <%= FormatDate(boardqna.results(i).regdate, "0000-00-00") %>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2"><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
<% if boardqna.results(0).userid <> "" then %>
    <tr>
      <td class="a" height="5" colspan="2">&nbsp;<b>���ֹ��Ǽ�</b> : <%= orderinfo.OrderCount %>&nbsp;<b>���ֹ��ݾ�</b> : <% = FormatNumber(orderinfo.TotalPrice,0) %>��</td>
    </tr>
    <tr>
      <td colspan="2"><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
    <tr>
      <td class="a" height="5" colspan="2">&nbsp;<b>�ֹ���ҰǼ�</b> : <%= orderinfo.MOrderCount %>&nbsp;<b>�ֹ���ұݾ�</b> : <% = FormatNumber(orderinfo.MTotalPrice,0) %>��</td>
    </tr>
    <tr>
      <td colspan="2"><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
<% end if %>
	<tr>
      <td class="a" height="5" colspan="2">&nbsp;<b>Site</b> : <%= boardqna.results(0).FextSiteName %>&nbsp;</td>
    </tr>
	<tr>
      <td colspan="2"><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
    <tr>
      <td colspan="2" class="a"><b>����</b> : <br><%= nl2br(db2html(boardqna.results(0).contents)) %><br><br></td>
    </tr>
    <tr>
      <td colspan="2"><hr></td>
    </tr>
    <tr>
      <td colspan="2"><input type="text" name="replytitle" size="30" value="<%= boardqna.results(0).replytitle %>"></td>
    </tr>
    <tr>
      <td colspan="2"><textarea name="replycontents" cols="80" rows="10"><%= db2html(boardqna.results(0).replycontents) %></textarea></td>
    </tr>
    <tr>
      <td colspan="2" align="right"><input type="button" value=" ��۴ޱ� " onclick="SubmitForm()">
      <a href="lecqna_list.asp">������� �̵�</a>
      </td>
    </tr>
</form>
</table>
<% if boardqna.results(0).userid <> "" or boardqna.results(0).orderserial <> "" then %>
<table cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" border="1">
<tr>
	<td colspan="3" bgcolor="#DDDDFF" align="center">���� ������ ���</td>
</tr>
<tr>
	<td width="200" align="center">����</td>
	<td width="100" align="center">����</td>
	<td width="100" align="center">�ۼ���</td>
</tr>
<tr>
<% if myqnalist.ResultCount < 0 then %>
<% else %>
<% for i = 0 to (myqnalist.ResultCount - 1) %>
	<td><a href="cscenter_qna_board_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>"><%= myqnalist.results(i).title %></a></td>
	<td align="center"><%= myqnalist.code2name(myqnalist.results(i).qadiv) %></td>
	<td align="center"><%= FormatDate(myqnalist.results(i).regdate, "0000-00-00") %></td>
</tr>
<% next %>
<% end if %>
</table>
<% end if %>
<!--
<% if (boardqna.results(0).emailok = "Y") then %>
<b>�̸���</b> : ������<br><br>
<% else %>
<b>�̸���</b> : ���ž���<br><br>
<% end if %>
-->
<!-- #include virtual="/lib/db/dbclose.asp" -->