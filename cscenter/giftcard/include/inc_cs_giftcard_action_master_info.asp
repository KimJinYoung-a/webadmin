<% if (IsDisplayCSMaster = true) then %>
	<%
	dim jupsugubun, jupsudefaulttitle

	jupsugubun = GetCSCommName("Z001", divcd)
	jupsudefaulttitle = ogiftcardordermaster.FOneItem.GetAccountdivName + " " + ogiftcardordermaster.FOneItem.GetJumunDivName + " ������ �ֹ����"

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
                <%= giftorderserial %>
                [<font color="<%= ogiftcardordermaster.FOneItem.CancelYnColor %>"><%= ogiftcardordermaster.FOneItem.CancelYnName %></font>]
                [<font color="<%= ogiftcardordermaster.FOneItem.IpkumDivColor %>"><%= ogiftcardordermaster.FOneItem.GetJumunDivName %></font>]
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
                <%= ogiftcardordermaster.FOneItem.FUserID %>(<font color="<%= ogiftcardordermaster.FOneItem.GetUserLevelColor %>"><%= ogiftcardordermaster.FOneItem.GetUserLevelName %></font>)
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
                <%= ogiftcardordermaster.FOneItem.FBuyname %>
                 &nbsp;
                 [<%= ogiftcardordermaster.FOneItem.FBuyHp %>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsStatusRegister) then %>
                	<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= jupsudefaulttitle %>" size="56" maxlength="56">
                <% else %>
                	<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= ocsaslist.FOneItem.Ftitle %>" size="56" maxlength="56">
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">����������</td>
            <td bgcolor="#FFFFFF">
                 [<%= ogiftcardordermaster.FOneItem.FReqHp %>]
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
	                [<a href="javascript:selectGubun('C004','CD01','����','�ܼ�����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�ܼ�����</a>]
	                [<a href="javascript:selectGubun('C004','CD05','����','ǰ��','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">ǰ��</a>]
	                [<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��Ÿ</a>]

                <% elseif (ocsaslist.FOneItem.IsReturnProcess) then %>
	                [<a href="javascript:selectGubun('C004','CD01','����','�ܼ�����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�ܼ�����</a>]
	                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ҷ�</a>]
	                [<a href="javascript:selectGubun('C006','CF01','��������','���߼�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�����</a>]
                    [<a href="javascript:selectGubun('C004','CD04','����','�����ȯ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�����ȯ</a>]
                    [<a href="javascript:selectGubun('C004','CD06','����','������ �ȸ���','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">������ �ȸ���(������)</a>]
                <% elseif (divcd="A009") or (divcd="A006") or (divcd="A700") or (divcd="A900") then %>
                	[<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��Ÿ</a>]

                <% elseif (divcd="A001") then %>
                	[<a href="javascript:selectGubun('C006','CF03','��������','���Ż�ǰ����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ����</a>]

                <% elseif (divcd="A002") then %>
	                [<a href="javascript:selectGubun('C006','CF04','��������','����ǰ����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(����)����ǰ����</a>]
	                [<a href="javascript:selectGubun('C005','CE05','��ǰ����','�̺�Ʈ�����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(MD)�̺�Ʈ�����</a>]

                <% elseif (divcd="A000") then %>
	                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ҷ�</a>]
	                [<a href="javascript:selectGubun('C006','CF01','��������','���߼�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">���߼�</a>]
	                [<a href="javascript:selectGubun('C006','CF02','��������','��ǰ�ļ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ļ�</a>]
	                [<a href="javascript:selectGubun('C004','CD04','����','�����ȯ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�����ȯ</a>]
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
            <td bgcolor="#FFFFFF">
            	<%= FormatNumber(ogiftcardordermaster.FOneItem.Fsubtotalprice,0) %>��
            	&nbsp;
                [<%= ogiftcardordermaster.FOneItem.GetAccountdivName %>]
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="2">��������</td>
            <td bgcolor="#FFFFFF" rowspan="2">
            	<textarea <% if IsStatusFinishing then response.write "class='textarea_ro' ReadOnly" else response.write "class='textarea'" end if %> name="contents_jupsu" cols="68" rows="6"><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">���������</td>
            <td bgcolor="#FFFFFF" valign="top">
            	[<%= ogiftcardordermaster.FOneItem.FReqEmail %>]<br>
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center"></td>
            <td bgcolor="#FFFFFF" valign="top">

            </td>
        </tr>
        <% if (IsStatusFinishing) or (IsStatusFinished) then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">ó������</td>
            <td bgcolor="#FFFFFF">
            	<textarea class='textarea' name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
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
				    	<textarea class="textarea" name="opencontents" cols="48" rows="5" readonly><%= ocsaslist.FOneItem.Fopencontents %></textarea>
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
