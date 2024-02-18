<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ü����
' History : 2007.10.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/user/partnerusercls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->

<script language='javascript'>
window.resizeTo(600,700);
</script>
<%
dim ogroup,opartner,i
dim makerid , takbae
dim groupid

makerid = RequestCheckvar(request("makerid"),32)
takbae = RequestCheckvar(request("takbaebox"),16)

set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid
opartner.GetOnePartnerNUser


set ogroup = new CPartnerGroup

if opartner.FResultCount>0 then
	ogroup.FRectGroupid = opartner.FOneItem.FGroupid
	ogroup.GetOneGroupInfo
end if


dim OReturnAddr
set OReturnAddr = new CCSReturnAddress

OReturnAddr.FRectMakerid = makerid
OReturnAddr.GetBrandReturnAddress


dim OCSBrandMemo
set OCSBrandMemo = new CCSBrandMemo

OCSBrandMemo.FRectMakerid = makerid
OCSBrandMemo.GetBrandMemo

dim brandmemo_found
if (OCSBrandMemo.Fbrandid = "") then
	brandmemo_found = "N"
else
	brandmemo_found = "Y"
end if


%>
<script language="javascript">

function SaveBrandInfo(frm){
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4">
			<b>�귣�� ����</b>
		</td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ �귣�� �⺻���� ] (������ ��ü�� �귣�庰�� ��ǰ������ �ٸ� �� �ֽ��ϴ�.)</td>
	</tr>

	<tr height="25">
		<td width="18%" bgcolor="<%= adminColor("tabletop") %>" >�귣��ID</td>
		<td width="40%" bgcolor="#FFFFFF"><b><%= opartner.FOneItem.FID %></b></td>
		<td width="18%" bgcolor="<%= adminColor("tabletop") %>">��Ʈ��Ʈ��</td>
		<td bgcolor="#FFFFFF"><b><%= opartner.FOneItem.Fsocname_kor %></b></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�����</td>
		<td bgcolor="#FFFFFF"><%= OReturnAddr.FreturnName %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��ȭ</td>
		<td bgcolor="#FFFFFF"><%= OReturnAddr.FreturnPhone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ���</td>
		<td bgcolor="#FFFFFF"><%= OReturnAddr.Freturnhp %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�̸���</td>
		<td bgcolor="#FFFFFF"><%= OReturnAddr.FreturnEmail %></td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ �ּ�</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			[<%= OReturnAddr.FreturnZipcode %>] <%= OReturnAddr.FreturnZipaddr %> <%= OReturnAddr.FreturnEtcaddr %>
		</td>
	</tr>

	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ �귣�� ������� ]</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���ǹ�ۿ���</td>
		<td bgcolor="#FFFFFF">
			<% if (opartner.FOneItem.IsFreeBeasong) then %>
				�׻� ������
			<% end if %>
			<% if (opartner.FOneItem.IsUpcheReceivePayDeliverItem) then %>
				���ҹ��
			<% end if %>
			<% if opartner.FOneItem.IsUpcheParticleDeliverItem then %>
				���ݺ� ������
			<% end if %>
			<% if ((opartner.FOneItem.IsUpcheParticleDeliverItem) or (opartner.FOneItem.IsUpcheReceivePayDeliverItem)) and Not(opartner.FOneItem.IsFreeBeasong) then %>
			<% else %>
				N
			<% end if %>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<td bgcolor="#FFFFFF">
			<% if opartner.FOneItem.IsUpcheParticleDeliverItem then %>
			<b><%=FormatNumber(opartner.FOneItem.FdefaultFreeBeasongLimit,0)%></b>�� �̻� ���Ž� ����<br>
			��ۺ� <b><%=FormatNumber(opartner.FOneItem.FdefaultDeliverPay,0)%></b>��
			<% end if %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�ŷ��ù��</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= opartner.FOneItem.Ftakbae_name %> (<%= opartner.FOneItem.Ftakbae_tel %>)</td>
	</tr>

	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ �귣�� �߰����� ]</td>
	</tr>
	<tr height="25">
		<form name=brandmemo method=post action="do_brandmemo_input.asp">
		<input type=hidden name=makerid value="<%= makerid %>">
		<input type=hidden name=mode value="<% if brandmemo_found = "Y" then %>modify<% else %>insert<% end if %>">
		<td bgcolor="<%= adminColor("tabletop") %>">ȸ�����ɿ���</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select class="select" name="is_return_allow">
		     	<option value='-' >-</option>
		     	<option value='Y' <% if (OCSBrandMemo.Fis_return_allow = "Y") then %>selected<% end if %>>Y</option>
		     	<option value='N' <% if (OCSBrandMemo.Fis_return_allow = "N") then %>selected<% end if %>>N</option>
	     	</select>
	    </td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��㰡�ɽð�</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select class="select" name="tel_start">
				<option value='0'>-- : --</option>
		     	<% for i = 6 to 15 %>
		     	<option value='<%= i %>' <% if (OCSBrandMemo.Ftel_start = i) then %>selected<% end if %>><%= i %>:00</option>
		    	<% next %>
	     	</select>
	     	~
			<select class="select" name="tel_end">
				<option value='0'>-- : --</option>
		     	<% for i = 12 to 21 %>
		     	<option value='<%= i %>' <% if (OCSBrandMemo.Ftel_end = i) then %>selected<% end if %>><%= i %>:00</option>
		    	<% next %>
	     	</select>
	      	(����� �ٹ�����
	      	<select class="select" name="is_saturday_work">
		     	<option value='-' >-</option>
		     	<option value='Y' <% if (OCSBrandMemo.Fis_saturday_work = "Y") then %>selected<% end if %>>Y</option>
		     	<option value='N' <% if (OCSBrandMemo.Fis_saturday_work = "N") then %>selected<% end if %>>N</option>
	     	</select>)
	     </td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�ް�����</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="text" size="10" name="vacation_startday" value="<%= OCSBrandMemo.Fvacation_startday %>" onClick="jsPopCal('brandmemo','vacation_startday');" style="cursor:hand;"> - <input type="text" size="10" name="vacation_endday" value="<%= OCSBrandMemo.Fvacation_endday %>" onClick="jsPopCal('brandmemo','vacation_endday');" style="cursor:hand;">
	     </td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ�޸�</td>
		<td bgcolor="#FFFFFF" colspan="3">
		<textarea class="textarea" name=brand_comment cols="70" rows="5"><% if (OCSBrandMemo.Fbrand_comment = "") then %>�����޸�(��󿬶���,ȯ�Ұ���,�±�ȯ���ɿ��� ��)<% else %><%= OCSBrandMemo.Fbrand_comment %><% end if %></textarea>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����������</td>
		<td bgcolor="#FFFFFF" colspan="3">
		<% if Len(OCSBrandMemo.Flast_modifyday) > 10 then %>
			<%= Left(OCSBrandMemo.Flast_modifyday) %>
		<% else %>
			<%= (OCSBrandMemo.Flast_modifyday) %>
		<% end if %>
		</td>
	</tr>
	<tr height="25" align="center">
		<td colspan="4" bgcolor="#FFFFFF" height="25">
			<input type="button" class="button" value="�߰���������" onclick="SaveBrandInfo(brandmemo)"></td>
		</td>
	</tr>
	</form>


	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ ��ü�⺻���� ]</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF"><b><%= ogroup.FOneItem.FCompany_name %></b></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�׷��ڵ�</td>
		<td bgcolor="#FFFFFF"><b><%= opartner.FOneItem.FGroupid %></b></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��ȭ</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_tel %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ѽ�</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_fax %></td>
	</tr height="25">
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�繫�� �ּ�</td>
		<td bgcolor="#FFFFFF" colspan=3>[<%= ogroup.FOneItem.Freturn_zipcode %>] <%= ogroup.FOneItem.Freturn_address %> <%= ogroup.FOneItem.Freturn_address2 %></td>
	</tr height="25">



	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ ��ü ��������� ]</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����ڸ�</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_phone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_email %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_hp %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��۴���ڸ�</td>
		<td bgcolor="#FFFFFF" colspan="3">�귣�庰�� ��ȸ �����մϴ�</td>
	</tr>
	<!--
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��۴���ڸ�</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= ogroup.FOneItem.Fdeliver_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_phone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_email %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_hp %></td>
	</tr>
	-->
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�������ڸ�</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_phone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_email %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_hp %></td>
	</tr>














	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ ��ü ����ڵ������ ]</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.FCompany_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fceoname %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����ڹ�ȣ</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= ogroup.FOneItem.Fcompany_no %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����������</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<%= ogroup.FOneItem.Fcompany_zipcode %>&nbsp;
			<%= ogroup.FOneItem.Fcompany_address %>&nbsp;
			<%= ogroup.FOneItem.Fcompany_address2 %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_uptae %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_upjong %></td>
	</tr>


	<tr align="center">
		<td colspan="4" bgcolor="#FFFFFF" height="25">
			<input type="button" class="button" value="�ݱ�" onclick="self.close();"></td>
		</td>
	</tr>

</table>

<%
set opartner = Nothing
set ogroup = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->