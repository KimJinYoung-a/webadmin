<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/board/lib/classes/offshopqnacls.asp" -->
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
        if (document.frm.replytitle.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }
        if (document.frm.replycontents.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }

        if (confirm("�Է��� ��Ȯ�մϱ�?") == true) { document.frm.submit(); }
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
<form method="post" name="delform" action="college_offshop_qna_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="mode" value="del">
</form>

<table width="580" border="0" align="center" cellpadding="0" cellspacing="3">
<form method="post" name="frm" action="college_offshop_qna_board_act.asp" onsubmit="return false">
<input type="hidden" name="imsitxt">
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
      <td colspan="2">
		<input type="text" name="replytitle" size="30" value="<%= boardqna.results(0).replytitle %>">
		  <% SelectBoxQnaPreface "02" %>&nbsp;
		  <% SelectBoxQnaCompliment "" %>
		</td>
    </tr>
    <tr>
      <td colspan="2"><textarea name="replycontents" cols="80" rows="10"><%= db2html(boardqna.results(0).replycontents) %></textarea></td>
    </tr>
    <tr>
      <td colspan="2" align="right"><input type="button" value=" ��۴ޱ� " onclick="SubmitForm()">
      <% if reffrom="itemqa" then %>
      <a href="itemqna_list.asp">������� �̵�</a>
      <% else %>
      <a href="college_offshop_qna_board_list.asp">������� �̵�</a>
      <% end if %>
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
	<td><a href="college_offshop_qna_board_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>"><%= myqnalist.results(i).title %></a></td>
	<td align="center"><%= myqnalist.code2name(myqnalist.results(i).qadiv) %></td>
	<td align="center"><%= FormatDate(myqnalist.results(i).regdate, "0000-00-00") %></td>
</tr>
<% next %>
<% end if %>
</table>
<% end if %>
<iframe name="PrefaceFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">
<!--

 function TnChangePreface(SelectGubun){
	PrefaceFrame.location.href="/admin/board/preface_select.asp?gubun=" + SelectGubun + "&userid=<%= boardqna.results(0).userid %>&masterid=02";
 }

 function TnChangeText(str){
var basictext;
basictext = "�ȳ��ϼ���, <%= boardqna.results(0).userid %>��.\n"
basictext = basictext + "�ٹ����� �ø������� ���ϰ� �ִ� <%= session("ssBctCname") %>�Դϴ�.\n"
basictext = basictext + "(����)\n"
basictext = basictext + "���������� �亯�� �Ǽ̴�����?\n"
basictext = basictext + "Ȥ, �ٸ� ���ǰ� �����ôٸ� �ٹ����� ��ī����(02-741-9070 : ����~��11��)�� ��ȭ�ּ���.\n"
basictext = basictext + "ģ���ϰ� �ż��ϰ� �亯�帮�ڽ��ϴ�.\n"
basictext = basictext + "�����մϴ�. ���� �Ϸ� �Ǽ���.\n"

	if(str == ''){
		document.frm.replycontents.value = basictext;
	}
	else{
		document.frm.replycontents.value = str;
	}
 }
<% if boardqna.results(0).replyuser = "" then %>
TnChangeText('');
<% end if %>
//-->
</script>
<iframe name="ComplimentFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">
<!--

 function TnChangeCompliment(SelectGubun){
	ComplimentFrame.location.href="/admin/board/compliment_select.asp?masterid=02&gubun=" + SelectGubun;
 }

 function TnChangeText2(str){

	if(str == ''){
	}
	else{
		document.frm.replycontents.value = document.frm.imsitxt.value + "\n" + str;
	}
 }
//-->
</script>
<!--
<% if (boardqna.results(0).emailok = "Y") then %>
<b>�̸���</b> : ������<br><br>
<% else %>
<b>�̸���</b> : ���ž���<br><br>
<% end if %>
-->
<!-- #include virtual="/lib/db/dbclose.asp" -->