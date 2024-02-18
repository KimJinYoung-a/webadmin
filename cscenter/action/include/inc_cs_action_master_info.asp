<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<% if (IsDisplayCSMaster = true) then %>
	<%
	dim jupsugubun, jupsudefaulttitle

	jupsugubun = GetCSCommName("Z001", divcd)
	jupsudefaulttitle = GetDefaultTitle(divcd, id, orderserial)

	'CS �����ÿ��� ��ǰ����(��ü���)/ȸ����û(�ٹ����ٹ��) �� �������� �ʰ�
	'����� �귣�������� �ִ°�� ��ǰ����(��ü���), ���°�� ȸ����û(�ٹ����ٹ��) ���� �����Ѵ�.
	if (IsStatusRegister = true) and (divcd = "A004" or divcd = "A010") then
		jupsugubun = "��ǰ����"
		jupsudefaulttitle = "��ǰ����"
	end if
	%>
	<tr >
	    <td >
	        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	        <tr>
	            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">��������</td>
	            <td bgcolor="#FFFFFF">
				    	<font style='line-height:100%; font-size:15px; color:blue; font-family:����; font-weight:bold'><%= jupsugubun %></font>
				    	&nbsp;
	                <% if (Not IsStatusRegister) then %>
				    	<font style='line-height:100%; font-size:15px; color:#CC3333; font-family:����; font-weight:bold'>[<%= ocsaslist.FOneItem.GetCurrstateName %>]</font>
				    	<% if ocsaslist.FOneITem.FDeleteyn<>"N" then %>
							<font style='line-height:100%; font-size:15px; color:#FF0000; font-family:����; font-weight:bold'>- ������ ����</font>
				    	<% end if %>
			    	<% end if %>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">�ֹ���ȣ</td>
	            <td bgcolor="#FFFFFF" width="200" >
	                <%= orderserial %>
	                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font>]
	                [<font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font>]
	            </td>
	        </tr>
	        <tr height="20">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">������</td>
	            <td bgcolor="#FFFFFF" >
	                <% if (IsStatusRegister) then %>
	                    <%= session("ssbctid") %>
	                <% else %>
	                    <%= ocsaslist.FOneItem.Fwriteuser %>
	                <% end if %>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">�ֹ���ID</td>
	            <td bgcolor="#FFFFFF">
	                <%= oordermaster.FOneItem.FUserID %>
	                (<font color="<%= getUserLevelColorByDate(oordermaster.FOneItem.fUserLevel, left(oordermaster.FOneItem.Fregdate,10)) %>">
					<%= getUserLevelStrByDate(oordermaster.FOneItem.fUserLevel, left(oordermaster.FOneItem.Fregdate,10)) %></font>)
	            </td>
	        </tr>
	        <tr height="20">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">�����Ͻ�</td>
	            <td bgcolor="#FFFFFF" >
	                <% if (IsStatusRegister) then %>
	                	<%= now() %>
	                <% else %>
	                	<%= ocsaslist.FOneItem.Fregdate %>
	                <% end if %>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">�ֹ�������</td>
	            <td bgcolor="#FFFFFF">
	                <%= oordermaster.FOneItem.FBuyname %>
	                 &nbsp;
	                 [<%= oordermaster.FOneItem.FBuyHp %>]
	            </td>
	        </tr>
	        <tr height="20">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
	            <td bgcolor="#FFFFFF" >
	                <% if (IsStatusRegister) then %>
						<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= jupsudefaulttitle %>" size="56" maxlength="56">
						<% SelectBoxCSTemplateGubunNew "30", "csreg_template", "" %>
						<iframe name="CSTemplateFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
	                <% else %>
	                	<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= ocsaslist.FOneItem.Ftitle %>" size="56" maxlength="56">
	                <% end if %>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">����������</td>
	            <td bgcolor="#FFFFFF">
	                 <%= oordermaster.FOneItem.FReqName %>
	                 &nbsp;
	                 [<%= oordermaster.FOneItem.FReqHp %>]
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
	            <td bgcolor="#FFFFFF">
	                <input type="hidden" name="gubun01" value="<%= ocsaslist.FOneItem.Fgubun01 %>">
	                <input type="hidden" name="gubun02" value="<%= ocsaslist.FOneItem.Fgubun02 %>">
	                <input class="text_ro" type="text" name="gubun01name" value="<%= ocsaslist.FOneItem.Fgubun01name %>" size="16" Readonly >
	                &gt;
	                <input class="text_ro" type="text" name="gubun02name" value="<%= ocsaslist.FOneItem.Fgubun02name %>" size="16" Readonly >
	                <input class="csbutton" type="button" value="����" onClick="divCsAsGubunSelect(frmaction.gubun01.value, frmaction.gubun02.value, frmaction.gubun01.name, frmaction.gubun02.name, frmaction.gubun01name.name, frmaction.gubun02name.name,'frmaction','causepop');">
	                <div id="causepop" style="position:absolute;"></div>

	                <!-- �Ϻ� ���� �̸� ǥ�� -->
	                <%
	                '��������
					'select top 100 m.comm_cd, m.comm_name, d.comm_cd, d.comm_name
					'from
					'	db_cs.dbo.tbl_cs_comm_code m
					'	left join db_cs.dbo.tbl_cs_comm_code d
					'	on
					'		m.comm_cd = d.comm_group
					'where
					'	1 = 1
					'	and m.comm_group = 'Z020'
					'	and m.comm_isdel <> 'Y'
					'	and d.comm_isdel <> 'Y'
					'order by m.comm_cd, d.comm_cd
	                %>
	                <% if (ocsaslist.FOneItem.IsCancelProcess) then %>
		                [<a href="javascript:selectGubun('C004','CD01','����','������','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">������</a>]
		                [<a href="javascript:selectGubun('C004','CD05','����','ǰ��','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">ǰ��</a>]
						<!--
						[<a href="javascript:selectGubun('C005','CE02','��ǰ����','��ǰ�Ҹ���','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�Ҹ���</a>]
						-->
						[<a href="javascript:selectGubun('C006','CF06','��������','�������','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�������</a>]
		                [<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��Ÿ</a>]
		                <% if IsStatusRegister then %>
		                	&nbsp; &nbsp; &nbsp;
		                	<div id="chkmodifyitemstockoutyn" style="display: inline;"><input type="checkbox" name="modifyitemstockoutyn" value="Y" checked> ǰ������ ����(�����ǰ)</div>
		                <% end if %>

	                <% elseif (ocsaslist.FOneItem.IsReturnProcess) then %>
		                [<a href="javascript:selectGubun('C004','CD01','����','������','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">������</a>]
		                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ҷ�</a>]
		                [<a href="javascript:selectGubun('C006','CF01','��������','���߼�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�����</a>]
						<!--
	                    [<a href="javascript:selectGubun('C004','CD04','����','�����ȯ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�����ȯ</a>]
						-->
	                    [<a href="javascript:selectGubun('C004','CD06','����','������ �ȸ���','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">������ �ȸ���(������)</a>]
						[<a href="javascript:selectGubun('C006','CF06','��������','�������','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�������</a>]
					<% elseif (divcd="A009") or (divcd="A006") or (divcd="A700") or (divcd="A900") then %>
						<% if (divcd="A700") then %>
							[<a href="javascript:selectGubun('C004','CD10','����','��ü��ǰ�Ұ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ü��ǰ�Ұ�</a>]
						<% end if %>
	                	[<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��Ÿ</a>]
					<% elseif (divcd="A060") then %>
	                	[<a href="javascript:selectGubun('C011','CK01','��޹���','��ҹ���','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ҹ���</a>]
	                	[<a href="javascript:selectGubun('C011','CK02','��޹���','��ȯ��ǰ����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ȯ��ǰ����</a>]
	                	[<a href="javascript:selectGubun('C011','CK03','��޹���','AS����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">AS����</a>]
	                	[<a href="javascript:selectGubun('C011','CK04','��޹���','��۹���','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��۹���</a>]
						[<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��Ÿ</a>]

	                <% elseif (divcd="A001") then %>
	                	[<a href="javascript:selectGubun('C006','CF03','��������','���Ż�ǰ����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ����</a>]

	                <% elseif (divcd="A002") then %>
		                [<a href="javascript:selectGubun('C006','CF04','��������','����ǰ����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(����)����ǰ����</a>]
		                [<a href="javascript:selectGubun('C005','CE05','��ǰ����','�̺�Ʈ�����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(MD)�̺�Ʈ�����</a>]

	                <% elseif (divcd="A000") then %>
		                [<a href="javascript:selectGubun('C004','CD08','����','ȸ�������','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">ȸ�������</a>]
		                [<a href="javascript:selectGubun('C004','CD09','����','������û','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">������û</a>]

		                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ҷ�</a>]
		                [<a href="javascript:selectGubun('C006','CF01','��������','���߼�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">���߼�</a>]
		                [<a href="javascript:selectGubun('C006','CF02','��������','��ǰ�ļ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ļ�</a>]
		                <!--
		                [<a href="javascript:selectGubun('C004','CD04','����','�����ȯ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�����ȯ</a>]
		                -->
		                <p>
		                * <font color="red"><b>�����ȯ</b></font>�� "�ɼǺ��� �±�ȯ" �Է� (�������뿡 �Է½� ������� �� ����)

	                <% elseif (divcd="A100") or (divcd="A111") then %>
	                	<!--
	                	* ������, ������ �ȸ���(������) �� ��� ȸ�����Ŀ� �±�ȯ ����Ѵ�.
	                	* ���� : http://logics.10x10.co.kr/v2/online/m_re_chulgo.asp
	                	-->
	                	[<a href="javascript:selectGubun('C004','CD01','����','������','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">������</a>]
		                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ҷ�</a>]
		                [<a href="javascript:selectGubun('C006','CF01','��������','���߼�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">���߼�</a>]
		                [<a href="javascript:selectGubun('C006','CF02','��������','��ǰ�ļ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ļ�</a>]
		                [<a href="javascript:selectGubun('C004','CD06','����','������ �ȸ���','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">������ �ȸ���(������)</a>]
					<% elseif (divcd="A999") then %>
						[<a href="javascript:selectGubun('C012','CL01','�߰�����','��ǰ����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ����</a>]
	                	[<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��Ÿ</a>]
	                <% end if %>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
	            <td bgcolor="#FFFFFF">
	            	<% if oordermaster.FOneItem.IsErrSubtotalPrice then %>
	            		<font color="red"><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice-realSubPaymentSum,0) %>��</font>
	            	<% else %>
	            		<%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice-realSubPaymentSum,0) %>��
					<% end if %>
	            	&nbsp;
	                [<%= oordermaster.FOneItem.JumunMethodName %>]

	                <% if (realdepositsum>0) then %>
	                   /&nbsp; <strong><%= FormatNumber(realdepositsum,0) %></strong>��&nbsp; [��ġ��]
	                <% end if %>
	                <% if (realgiftcardsum>0) then %>
	                   /&nbsp; <strong><%= FormatNumber(realgiftcardsum,0) %></strong>��&nbsp; [��ǰ��]
	                <% end if %>


	                <% if (oordermaster.FOneItem.Faccountdiv="110") then %>
	                	(OK Cashbag��� : <strong><%= FormatNumber(oordermaster.FOneItem.FokcashbagSpend,0) %></strong> ��)
	                <% end if %>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="6">
					��������<br><br>
	    			<input type="button" class="button" value="�ð�" onClick="WriteNowDateString(document.frmaction.contents_jupsu)">
				</td>
	            <td bgcolor="#FFFFFF" rowspan="6">
					<table width="100%" height="100%" border="0" align="center" cellpadding="2" cellspacing="0"  class="a">
						<tr>
							<td width="420">
								<textarea <% if IsStatusFinishing then response.write "class='textarea_ro' ReadOnly" else response.write "class='textarea'" end if %> id="contents_jupsu" name="contents_jupsu" cols="68" rows="12"><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea>
							</td>
							<td align="left">
								<%
								if (IsTempEventAvail = True) or (IsTempEventAvail_Str <> "") then
									response.Write "<br>�����ǰ �̺�Ʈ �ֹ�<br>"
									response.Write "&nbsp; - &nbsp; �귣�� : " & IsTempEventAvail_Makerid & "<br>"
									if (IsTempEventAvail_Str <> "") then
										response.Write "&nbsp; - &nbsp; ����Ұ� : " & IsTempEventAvail_Str & "<br>"
									else
										%>
										&nbsp; - &nbsp; <input type="button" class="button" onClick="jsCheckApplyEvent(frmaction);" value="�����ǰ����"><br>
										<%
									end if
								end if
								%>
							</td>
						</tr>
					</table>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">���������</td>
	            <td bgcolor="#FFFFFF" valign="top">
	            	[<%= oordermaster.FOneItem.FReqZipCode %>]<br>
	                <%= oordermaster.FOneItem.FReqZipAddr %><br>
	                <%= oordermaster.FOneItem.FReqAddress %>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">������</td>
	            <td bgcolor="#FFFFFF">
	            	<% if ocsaslist.FOneItem.IsRequireSongjangNO and ocsOrderDetail.FResultCount > 0 and (divcd = "A004" or divcd = "A010") and (Not IsStatusRegister) then %>
					<% Call drawSelectBoxDeliverCompany ("songjangdiv_tmp",ocsOrderDetail.FItemList(ocsOrderDetail.FResultCount - 1).Fsongjangdiv) %>
					<%= ocsOrderDetail.FItemList(ocsOrderDetail.FResultCount - 1).Fsongjangno %>
			        <% end if %>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">�ù�����</td>
	            <td bgcolor="#FFFFFF">
	            	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
					<%
					Select Case ocsaslist.FOneItem.FsongjangRegGubun
						Case "U"
							Response.Write("�ٹ�����(��ü) ����")
						Case "C"
							Response.Write("����������")
						Case "T"
							Response.Write("���� ����")
						Case Else
							Response.Write ocsaslist.FOneItem.FsongjangRegGubun
					End Select
					%>
			        <% end if %>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">�ù�������</td>
	            <td bgcolor="#FFFFFF">
	            	<%
					if ocsaslist.FOneItem.IsRequireSongjangNO then
						if Not IsNull(ocsaslist.FOneItem.FsongjangRegUserID) and (ocsaslist.FOneItem.FsongjangRegUserID <> "") then
							Response.Write ocsaslist.FOneItem.FsongjangRegUserID
							if (ocsaslist.FOneItem.FsongjangRegUserID = oordermaster.FOneItem.FUserID) then
								Response.Write " (��)"
							elseif (ocsaslist.FOneItem.Frequireupche = "Y") and (ocsaslist.FOneItem.FsongjangRegUserID = ocsaslist.FOneItem.Fmakerid) then
								Response.Write " (��ü)"
							end if
						end if
					end if
					%>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">�����ȣ</td>
	            <td bgcolor="#FFFFFF">
					<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
					<%= ocsaslist.FOneItem.FsongjangPreNo %>
					<% end if %>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">�ù�����</td>
	            <td bgcolor="#FFFFFF">
	            	<!-- �ڵ� Ȯ���Ұ� -->
	            	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
				        <% Call drawSelectBoxDeliverCompany ("songjangdiv",ocsaslist.FOneItem.Fsongjangdiv) %>
				        <input type="text" class="text" name="songjangno" value="<%= ocsaslist.FOneItem.Fsongjangno %>" size="14" maxlength="16">
				        <% dim ifindurl : ifindurl = DeliverDivTrace(ocsaslist.FOneItem.Fsongjangdiv) %>
				        <% if (ocsaslist.FOneItem.Fsongjangdiv="24") then %>
	                		<a href="javascript:popDeliveryTrace('<%= ifindurl %>','<%= ocsaslist.FOneItem.Fsongjangno %>');">����</a>
	                	<% else %>
				            <a href="<%= ifindurl + ocsaslist.FOneItem.Fsongjangno %>" target="_blank">����</a>
				        <% end if %>
				        <input type="button" class="button" value="����" onClick="changeSongjang('<%= id %>');">
			        <% end if %>
	            </td>
	        </tr>

			<% if False and InStr(",A000,A100,A001,A002,A009,A006,A012,", divcd) > 0 and Not IsStatusFinishing and Not IsUpcheConfirmState then %>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">
					�Ϸᱸ��
				</td>
	            <td bgcolor="#FFFFFF" colspan="3">
					<input type="radio" id="needChkYN_X" name="needChkYN" value="X" <%= CHKIIF(ocsaslist.FOneItem.FneedChkYN="X", "checked", "") %> > ��üó���Ϸ�� ��ÿϷ�
					<input type="radio" id="needChkYN_F" name="needChkYN" value="F" <%= CHKIIF(ocsaslist.FOneItem.FneedChkYN="F", "checked", "") %> > ������ Ȯ�� �ʿ�
	            </td>
	        </tr>
			<% end if %>

	        <% if (IsStatusFinishing) or (IsUpcheConfirmState) or (IsStatusFinished) then %>
		        <tr bgcolor="#F4F4F4">
		            <td bgcolor="<%= adminColor("topbar") %>" align="center">
		            	ó������
		            	<% if (IsUpcheConfirmState) and (IsRefASExist) and (ocsaslist.FOneItem.Frequireupche = "Y") then %>
		            		<br><br>(��ü���)<br>+<br>(��üȸ��)
		            	<% end if %>
		            </td>
		            <td bgcolor="#FFFFFF">
			            <% if True or (IsUpcheConfirmState) then %>
							<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a"><tr><td width="450">
							<% if IsUpcheConfirmState and (IsRefASExist) and (ocsaslist.FOneItem.Frequireupche = "Y") then %>
				            	<textarea class='textarea_ro' readOnly name="contents_finish" cols="68" rows="4"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
				            	<textarea class='textarea_ro' readOnly name="contents_finish1" cols="68" rows="4"><%= ioneRefas.FOneItem.Fcontents_finish %></textarea>
							<% else %>
								<textarea class='textarea_ro' name="contents_finish" cols="68" rows="9"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
							<% end if %>
							</td>
							<td style="vertical-align: middle; text-align: left;">
								<%
								Select Case ocsaslist.FOneItem.FneedChkYN
									Case "Y"
										response.write "<font color='red'><b>Ȯ�� �� ó��(��ü ���)</b></font>"
									Case "N"
										response.write "<b>��ÿϷ�</b>(Ȯ�κ��ʿ�)"
									Case "F"
										response.write "<font color='red'><b>Ȯ�� �� ó��(CS ���)</b></font>"
									Case Else
										response.write "-"
								End Select
								%>
							</td></tr></table>
			            <% else %>
			            	<textarea class='textarea' name="contents_finish" cols="68" rows="9"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
			            <% end if %>
		            </td>
		            <td bgcolor="<%= adminColor("pink") %>" align="center">ó������<br>������<br>�����Է�</td>
		            <td bgcolor="#FFFFFF">
		            	<table border="0" cellspacing="0" cellpadding="0" class="a" valign="top">
		            	<tr>
						    <td>
						    	<input class="text" type="text" name="opentitle" value="<%= ocsaslist.FOneItem.Fopentitle %>" size="48" maxlength="60" readonly>
						    </td>
						</tr>
						<tr>
						    <td>
						    	<textarea class="textarea" name="opencontents" cols="48" rows="7" readonly><%= ocsaslist.FOneItem.Fopencontents %></textarea>
						    </td>
						</tr>
						</table>
					</td>
		        </tr>
	        <% end if %>
	        </table>
		</td>
	</tr>
<% end if %>
