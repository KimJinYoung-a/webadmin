<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ������
' Hieditor : 2012.03.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/shopcscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->

<%
dim sqlStr
Dim mode	: mode = requestCheckVar(req("mode",""),32)
Dim CsAsID	: CsAsID = requestCheckVar(req("masteridx",""),10)

Dim songjangDiv	: songjangDiv = requestCheckVar(req("songjangDiv",""),3)
Dim songjangNo	: songjangNo  = requestCheckVar(req("songjangNo",""),32)

If mode = "SONGJANG" Then 

	sqlStr = "UPDATE db_shop.dbo.tbl_shopjumun_cs_master" & VbCrlf	
	sqlStr = sqlStr + " SET currState	  = (case when currState < 'B004' then 'B004' else currState end)" & VbCrlf
    sqlStr = sqlStr + " WHERE masteridx =" & CsAsID
    
    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr

	sqlStr = "UPDATE db_shop.dbo.tbl_shopjumun_cs_delivery" & VbCrlf	
	sqlStr = sqlStr + " SET songjangDiv ="&songjangDiv&"" & VbCrlf
	sqlStr = sqlStr + " , songjangNo ='"&songjangNo&"'" & VbCrlf
	sqlStr = sqlStr + " , senddate=getdate()" & VbCrlf
    sqlStr = sqlStr + " WHERE asid =" & CsAsID
    
    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
    
	response.write "<script type='text/javascript'>" & vbCrLf
	response.write "	alert('��ϵǾ����ϴ�.');" & vbCrLf
	response.write "	opener.location.reload();" & vbCrLf
	response.write "	window.close();" & vbCrLf
	response.write "</script>" & vbCrLf
	dbget.close()	:	response.End 
End If 

dim ocsaslist
set ocsaslist = New corder
	ocsaslist.FRectCsAsID = CsAsID
	
	if (CsAsID<>"") then
	    ocsaslist.fGetOneCSASMaster
	end if

if (ocsaslist.ftotalcount < 1) then
    response.write "��ϵ� as ������ �����ϴ�"
    response.end
end if
%>

<script type='text/javascript'>

function jsSubmit()
{
	var f = document.frmWrite;

	if (!f.songjangDiv.value)
	{
		alert("�ù�ȸ�縦 ������ �ּ���.");
		f.songjangDiv.focus();
		return;
	}
	if (!f.songjangNo.value || f.songjangNo.value.length < 8)
	{
		alert("�����ȣ�� �Է��� �ּ���.");
		f.songjangNo.focus();
		return;
	}

	f.submit();
}

</script>

<!---- �˾�ũ�� 400x195 ---->
<form name="frmWrite" action="popChangeSongjang.asp">
<input type="hidden" name="mode" value="SONGJANG">
<input type="hidden" name="masteridx" value="<%=CsAsID%>">
<table width="400" border="0" cellspacing="0" cellpadding="0">
<tr>
	<!---- �˾����� ���� ---->
	<td valign="top" bgcolor="#af1414">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td align="right"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/popup_logo.gif" width="254" height="15"></td>
		</tr>
		<tr>
			<td height="42" valign="bottom" style="padding:0 0 5px 20px"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/title_invoice.gif" width="107" height="23"></td>
		</tr>
		</table>
	</td>
	<!---- �˾����� �� ---->
</tr>
<tr>
	<td><br><Br></td>
</tr>
<tr>
	<td align="center" class="gray11px02" style="padding:0 0 20px 0px;">
		<table border="0" cellspacing="0" cellpadding="0" style="border-top:3px solid #be0808;" class="a">
		<tr>
			<td width="100" height="31" bgcolor="#fcf6f6" class="bbstxt01" style="border-bottom:1px solid #eaeaea;">�ù� ȸ��</td>
			<td width="200" style="border-bottom:1px solid #eaeaea;padding:0 1px 0 20px;">
				<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td>
						<%Call drawSelectBoxDeliverCompany("songjangDiv",ocsaslist.FOneItem.FsongjangDiv)%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="100" height="31" bgcolor="#fcf6f6" class="bbstxt01" style="border-bottom:1px solid #eaeaea;">�ù� ��ȣ</td>
			<td style="border-bottom:1px solid #eaeaea;padding:0 1px 0 20px;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><input name="songjangNo" type="text" class="input_02" style="width:140px;height:20px;" value="<%=ocsaslist.FOneItem.FsongjangNo%>" /></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" style="padding-bottom:10px;">
		<table border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td style="padding-right:7px;"><a href="javascript:jsSubmit();" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/btn_confirm.gif" width="58" height="24" border="0"/></a></td>
			<td><a href="javascript:window.close();" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2009/order/btn_cancel02.gif" width="58" height="24" border="0"/></a></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<%
Sub drawSelectBoxDeliverCompany(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>����</option><%
   query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' "
   query1 = query1 + " order by divcd"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Trim(Lcase(selectedId)) = Trim(Lcase(rsget("divcd"))) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcd")&"' "&tmp_str&">" & "" & replace(db2html(rsget("divname")),"'","") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub
%>
<%
set ocsaslist = Nothing
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->