<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%
dim orderserial : orderserial= requestCheckVar(request("orderserial"),11)
dim id : id = requestCheckVar(request("id"),10)

dim oneOrder
set oneOrder = new COrderMaster
oneOrder.FRectOrderserial = orderserial
oneOrder.QuickSearchOrderMaster

dim oneOrderdetail
set oneOrderdetail = new COrderMaster
oneOrderdetail.FRectOrderserial = orderserial
oneOrderdetail.FRectDetailIdx = id

if oneOrder.FResultCount>0 then
    oneOrderdetail.GetOneOrderDetail
end if


dim i

if ((oneOrder.FResultCount<1) or (oneOrderdetail.FResultCount<1)) then
    response.write "<script language='javascript'>alert('�ֹ� ������ �������� �ʽ��ϴ�.');</script>"
    response.write "<script language='javascript'>window.close();</script>"
    dbget.close()	:	response.End
end if


dim IsRequireDetailEditEnable
IsRequireDetailEditEnable = (oneOrderdetail.FOneItem.IsRequireDetailExistsItem) and (oneOrderdetail.FOneItem.Fcurrstate>2)
%>
<script language='javascript'>

function editHandMadeRequire(frm){
    var detailArr='';
<% if Not (IsRequireDetailEditEnable) then %>
    if (!confirm('���� ���� ���°� �ƴմϴ�. \n��� ���� �Ͻðڽ��ϱ�?')){
        return;
    }
<% end if %>

    if (frm.requiredetail!=undefined){
        if (frm.requiredetail.value.length<1){
            alert('�ֹ� ���� ������ �Է��� �ּ���.');
            frm.requiredetail.focus();
            return;
        }

		// 255 -> 512(2013-06-19, skyer9)
        if(GetByteLength(frm.requiredetail.value)>512) {
    		alert('���� �Է��� �ѱ� �ִ� 240�� ���� �����մϴ�.');
    		frm.requiredetailedit.focus();
    		return;
    	}
	}else{
	    <% if (oneOrderdetail.FOneItem.FItemNo>1) then %>
        for (i=0;i<<%=oneOrderdetail.FOneItem.FItemNo%>;i++){
			// 255 -> 512(2013-06-19, skyer9)
            if(GetByteLength(eval("frm.requiredetail" + i).value)>512){
    			alert('���� �Է��� �ѱ� �ִ� 240�� ���� �����մϴ�.');
    			eval("frm.requiredetailedit" + i).focus();
    			return;
    		}

            detailArr = detailArr + eval("frm.requiredetail" + i).value+'||';

        }

        if(GetByteLength(detailArr)>1024){
			alert('���� �Է��հ�� �ѱ� �ִ� 512�� ���� �����մϴ�.');
			frm.requiredetail.focus();
			return;
		}
        <% end if %>
	}

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }

}

window.onload = function()
{
	popupResize(380);
}
</script>

<% if (oneOrderdetail.FResultCount>0) then %>
<form name="frm" method="post" action="/cscenter/ordermaster/order_info_edit_process.asp">
<input type="hidden" name="mode" value="edithandmadereq">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="detailidx" value="<%= id %>">

<!---- �˾�ũ�� 360x340 ----->
<table width="360" border="0" cellspacing="0" cellpadding="0" class="a">
  <tr>
    <td><img src="http://fiximage.10x10.co.kr/web2009/order/popup_wordmodify.gif" width="360" height="60" /></td>
  </tr>
    <tr>
        <td align="center">
            <table width="300" border="0" cellpadding="2" cellspacing="0" class="a">
                <tr>
                  <td width="50"><img src="<%= oneOrderdetail.FOneItem.FSmallImage %>" width="50" height="50"></td>
                  <td>
                    <%= oneOrderdetail.FOneItem.FItemName %>
                    <% if oneOrderdetail.FOneItem.FItemoptionName<>"" then %>
    		        <font color="blue">[<%= oneOrderdetail.FOneItem.FItemoptionName %>]</font>
    		        <% end if %>
                  </td>
                  <td width="40"><%= oneOrderdetail.FOneItem.FItemNo %>��</td>
                  <td width="50" align="right"><%= FormatNumber(oneOrderdetail.FOneItem.FItemCost,0) %>��</td>
                </tr>
            </table>
        </td>
  </tr>
  <% if oneOrderdetail.FOneItem.FItemNo=1 then %>
  <tr>
    <td align="center"><span style="padding-top:1px;">
    	    <textarea name="requiredetail" cols="40" rows="5" class="txt_b1"><%= oneOrderdetail.FOneItem.Frequiredetail %></textarea>
    </span></td>
  </tr>
  <% else %>
  <% for i=0 to oneOrderdetail.FOneItem.FItemNo-1 %>
  <tr><td style="padding-left:30px"><%= i+1 %>�� ��ǰ ����</td></tr>
  <tr>
    <td align="center"><span style="padding-top:1px;">
    	    <textarea name="requiredetail<%= i %>" cols="40" rows="5" class="txt_b1"><%= splitValue(oneOrderdetail.FOneItem.Frequiredetail,CAddDetailSpliter,i) %></textarea>
    </span></td>
  </tr>
  <tr height="10"><td style="padding-top:20px"></td></tr>
  <% next %>
  <% end if %>
  <tr>
    <td style="padding:15px;">
    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
      <tr>
        <td width="10" valign="top" style="padding-top:2px"><img src="http://fiximage.10x10.co.kr/web2009/order/bullet_gray02.gif" width="10" height="7"></td>
        <td>���� ��ǰ�� 2�� �ֹ��Ͻ� ��� ���� ������ �ٸ��� ���� �Է��� �ֽñ� �ٶ��ϴ�.</td>
      </tr>
      <tr>
        <td width="10" valign="top" style="padding-top:2px"><img src="http://fiximage.10x10.co.kr/web2009/order/bullet_gray02.gif" width="10" height="7"></td>
        <td>
            ���� ������ ������� 1��° ��ǰ���� �Է��Ͻñ� �ٶ��ϴ�.
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td align="center" style="padding-bottom:10px;"><table border="0" cellspacing="0" cellpadding="0" class="a">
        <tr>
          <td style="padding-right:7px;"><a href="javascript:editHandMadeRequire(frm)" onfocus="blur()"><img src="http://fiximage.10x10.co.kr/web2009/order/btn_modiry02.gif" width="58" height="24" border="0"/></a></td>
          <td><a href="javascript:window.close();" onfocus="blur()"><img src="http://fiximage.10x10.co.kr/web2009/order/btn_cancel02.gif" width="58" height="24" border="0"/></a></td>
        </tr>
    </table></td>
  </tr>
</table>

<% end if %>

<%
set oneOrder       = Nothing
set oneOrderdetail = Nothing
%>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
