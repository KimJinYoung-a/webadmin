<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ������
' Hieditor : 2011.03.14 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->

<%
dim i, j , orderno ,oordermaster, oorderdetail ,masteridx
	masteridx = requestCheckVar(request("masteridx"),10)

set oordermaster = new COrder
	oordermaster.frectmasteridx = masteridx
	oordermaster.fQuickSearchOrderMaster

set oorderdetail = new COrder
oorderdetail.frectmasteridx = masteridx
oorderdetail.fQuickSearchOrderDetail
%>

<html>
<head>
<title>[�ٹ�����] ��ſ��� ������ ���θ� 10x10 = tenbyten</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="http://www.10x10.co.kr/lib/css/2009ten.css" rel="stylesheet" type="text/css">
<body  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="701" border="0" align="center" cellpadding="0" cellspacing="0">
<% if oordermaster.ftotalcount>0 then %>
<tr>
    <td width="701" style="padding-top:15">
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td style="border:1px solid #E1E1E1">
                <img src="http://fiximage.10x10.co.kr/web2007/cs_center/receipt_top.gif">
            </td>
        </tr>
        </table>
    </td>
</tr>
<tr>
    <td>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            <!-- ������ ���� ���� -->
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td style="padding-top:15">
                  <font color="#000000"><b>* �ֹ�����</b></font></td>
              </tr>
              <tr>
                <td>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="30" bgcolor="#F7F7F7" style="padding:5 0 5 10;border-top:1px solid #E1E1E1"><font color="#000000">�ֹ��Ͻ� ��</font></td>
                              <td style="padding:5 0 5 10;border-top:1px solid #E1E1E1">
                                [<%= oordermaster.FOneItem.FBuyName %>]<br>
                                <%= oordermaster.FOneItem.FBuyPhone %> / <%= oordermaster.FOneItem.FBuyHp %>                      </td>
                              <td width="70" bgcolor="#F7F7F7" style="padding:5 0 5 10;border-top:1px solid #E1E1E1"><font color="#000000">�����ô� ��</font></td>
                              <td style="padding:5 0 5 10;border-top:1px solid #E1E1E1">
                                [<%= oordermaster.FOneItem.FReqName %>]<br>
                                <%= oordermaster.FOneItem.Freqzipaddr %><br><%= oordermaster.FOneItem.Freqaddress %><br>
                                <%= oordermaster.FOneItem.FReqPhone %> / <%= oordermaster.FOneItem.FReqHp %>                      </td>
                    </tr>
                    <tr>
						<td width="70" height="25" bgcolor="#F7F7F7" style="padding:5 0 5 10;border-top:1px solid #E1E1E1;border-bottom:1px solid #E1E1E1">
							<font color="#000000">�ϷĹ�ȣ(�ֹ���ȣ)</font>
						</td>
						<td width="70" style="padding:5 0 5 10;border-top:1px solid #E1E1E1;border-bottom:1px solid #E1E1E1">
							<%= oordermaster.FOneItem.fmasteridx %> (<%= oordermaster.FOneItem.Forderno %>)
						</td>
						<td width="70" bgcolor="#F7F7F7" style="padding:5 0 5 10;border-top:1px solid #E1E1E1;border-bottom:1px solid #E1E1E1"><font color="#000000">�ֹ�����</font></td>
						<td width="70" style="padding:5 0 5 10;border-top:1px solid #E1E1E1;border-bottom:1px solid #E1E1E1"><%= left(oordermaster.FOneItem.FRegDate,10) %></td>
                    </tr>
                  </table>
                  </td>
              </tr>
            </table>
                  <!-- ������ ���� �� -->          </td>
        </tr>
        <tr>
          <td>
            <!-- ���Ż�ǰ ���� ���� -->
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td style="padding-top:15">
                  <font color="#000000"><b>* �ֹ�����</b></font>                </td>
              </tr>
              <tr>
                <td style="padding:2 0 0 0 ;border-top:1px solid #E1E1E1;border-bottom:2px solid #D4E4D1"  background="http://fiximage.10x10.co.kr/web2007/cs_center/top_bg.gif" height="30">
                  <table width="100%" cellspacing="0" cellpadding="0" height="10">
                    <tr>
						      <td width="60" height="0" style="border-right:1px solid #E1E1E1;padding:0 5 0 5"><div align="center"><font color="#000000">��ǰ�ڵ�</font></div></td>
                              <td style="border-right:1px solid #E1E1E1;padding:0 5 0 5"><div align="center"><font color="#000000">��ǰ��[�ɼ�]</font></div></td>
                              <td width="65" height="0" style="border-right:1px solid #E1E1E1;padding:0 5 0 5"><div align="center"><font color="#000000">�ǸŰ�</font></div></td>
						      <td width="30" height="0" style="border-right:1px solid #E1E1E1;padding:0 5 0 5"><div align="center"><font color="#000000">����</font></div></td>
						      <td width="60" height="0" style="padding:0 5 0 5"><div align="center"><font color="#000000">�Ұ�ݾ�</font></div></td>
                    </tr>
                  </table>                </td>
              </tr>
              <tr>
                <td>
                  <% for i=0 to oorderdetail.FResultCount-1 %>
                  <% if oorderdetail.FItemList(i).Fitemid <>0 then %>

                  <table width="100%" style="border-bottom:1px solid #DCDCDC" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="60" align="center" valign="middle" style="padding:0 5 0 5">
                        <%= oorderdetail.FItemList(i).fitemgubun%>-<%=CHKIIF(oorderdetail.FItemList(i).fitemid>=1000000,Format00(8,oorderdetail.FItemList(i).fitemid),Format00(6,oorderdetail.FItemList(i).fitemid))%>-<%=oorderdetail.FItemList(i).fitemoption %>
                        <br>
                        <% if oorderdetail.FItemList(i).Fisupchebeasong="N" then %>
                        �ٹ�����
                        <% elseif oorderdetail.FItemList(i).Fisupchebeasong="Y" then %>
                        <font color="red">��ü����</font>
                        <% end if %>
                      </td>
                      <td align="left" valign="middle" style="padding:0 5 0 5">
                        <%= oorderdetail.FItemList(i).FItemName %>
                        <br>
                        <font color="blue"><%= oorderdetail.FItemList(i).FItemoptionName %></font>
                      </td>
                      <td width="65" align="right" valign="middle" style="padding:0 5 0 5">
                        <% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
                            <%= FormatNumber(oorderdetail.FItemList(i).fsellprice,0) %> ��
                        <% else %>
                            <font color="red">���</font>
                        <% end if %>
                      </td>
                      <td width="30" align="center" valign="middle" style="padding:0 5 0 5"><%= oorderdetail.FItemList(i).FItemNo %></td>
                      <td width="60" align="right" valign="middle" style="padding:0 5 0 5">
                        <% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
                            <%= FormatNumber((oorderdetail.FItemList(i).fsellprice * oorderdetail.FItemList(i).FItemNo),0) %> ��
                        <% else %>
                            <font color="red">���</font>
                        <% end if %>
                      </td>
                    </tr>
                  </table>
                  <% end if %>
                  <% next %>

                </td>
              </tr>
            </table>
                  <!-- ���Ż�ǰ ���� �� -->
          </td>
        </tr>
    </table>
  </td>
</tr>

<tr>
  <td align="left">
    <table width="700" border="0" align="left" cellpadding="0" cellspacing="0">
    <tr>
      <td height="80" align="left" valign="bottom" style="border-top:1px solid #dddddd"><img src="http://fiximage.10x10.co.kr/web2007/cs_center/receipt_bottom.gif"></td>
    </tr>
    </table>
  </td>
</tr>
<tr>
    <td align="center" style="padding:15">
		<a href="javascript:window.print();"><img src="http://fiximage.10x10.co.kr/web2007/cs_center/print_btn.gif" border="0"></a>
    </td>
</tr>
<% else %>
�ֹ������� �����ϴ�
<% end if %>
</table>

</body>
</html>

<%
set oorderdetail = Nothing
%>
<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->