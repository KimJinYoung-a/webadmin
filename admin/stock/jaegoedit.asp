<%@ language = vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����ľ�
' History : 2007.07.13 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim fnow,idx, fmode , order , jaego,smallimage,itemid,makerid,itemname,itemoption		'��������
dim realstock,basicstock								'��������
	idx = html2db(request("idx"))							'���̺��� �ε������� �޾ƿ´�
	fmode = html2db(request("mode")	)						'��屸��
	order = left(now(),10)									'�۾�������
	jaego = html2db(request("jaego"))						'��������ľ������
	smallimage = html2db(request("smallimage"))				'�̹���
	itemid = request("itemid")						'��ǰid
	makerid = html2db(request("makerid"))					'�귣���
	itemname = html2db(request("itemname"))					'��ǰ��
	itemoption = html2db(request("itemoption"))				'��ǰ�ɼ��ڵ�
	realstock = request("realstock")						'�������
	basicstock = request("basicstock")						'����ľǿ����		
	
%>
<% 
dim sql , refer , sql111			'��������
%>			

<!--����������-->
<% if fmode = "edit" then %>
	<%	 
	dim sql101,fitemgubun1,fitemid1,fitemoption1,fitemname1,fitemoptionname1,fmakerid1
	dim fregdate1,freguserid1,forderingdate1,fbasicstock1,frealstock1,ffinishuserid1,fsmallimage1
	
	sql101 = "select"
	sql101 = sql101 & " b.smallimage,b.itemname,b.makerid,b.listimage,"
	sql101 = sql101 & " c.optionname , a.*"
	sql101 = sql101 & " from [db_summary].[dbo].tbl_req_realstock a"
	sql101 = sql101 & " join db_item.[dbo].tbl_item b"
	sql101 = sql101 & " on a.itemid = b.itemid"
	sql101 = sql101 & " left join [db_item].[dbo].tbl_item_option c" 
	sql101 = sql101 & " on a.itemid = c.itemid"
	sql101 = sql101 & " where 1=1 and idx = "& idx &""
	
	'response.write sql101&"<br>"	
	rsget.open sql101,dbget,1
		fitemgubun1 = rsget("itemgubun")				'��ǰ����
		fitemid1 = rsget("itemid")						'��ǰ��ȣ
		fitemoption1 = rsget("itemoption")		'�ɼ��ڵ�	
		fitemname1 = rsget("itemname")					'��ǰ��
		fitemoptionname1 = rsget("optionname")		'�ɼǸ�
		fmakerid1 = rsget("makerid")					'�귣��id
		fregdate1 = rsget("regdate")					'�����
		freguserid1 = rsget("reguserid")				'������id	
		fbasicstock1 = rsget("basicstock")				'����ľ����
		fsmallimage1 = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")				'��ǰ�̹���
		frealstock1 = rsget("realstock")				'�ǻ簹��
	rsget.close				
	%>	
	
	<script language="javascript">
	function sendit()
	{
	if(document.form1.jaego.value==""){
	alert("����ľ��Ͻ� ������ �Է��ϼ���.")
	document.form1.jaego.focus();
	}
	
	else
	document.form1.mode.value='edit'
	document.form1.submit();
	}
	</script>
	
	<!--ǥ ������-->
	<body>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
		<tr height="10" valign="bottom">
			<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
			<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
			<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
		</tr>
		<tr height="25" valign="top">
			<td background="/images/tbl_blue_round_04.gif"></td>
			<td background="/images/tbl_blue_round_06.gif">
				<img src="/images/icon_star.gif" align="absbottom">
				<font color="red" size=2><strong>������ </strong> / ���ǻ��� : �ý�����(�ѿ��) </font>
				</td>
				
			<td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
		<tr valign="top">
			<td background="/images/tbl_blue_round_04.gif"></td>
			<td></td>
			<td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
		<tr  height="10" valign="top">
			<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
			<td background="/images/tbl_blue_round_06.gif"></td>
			<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
		</tr>
		</tr>
	</table>
	<!--ǥ ��峡-->
	
	<!--��ǰ���̺����-->
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	 <form method="get" name="form1" action="jaegototalsubmit.asp">  	
	 <input type="hidden" name="itemid" value="<%= fitemid1 %>">
	 <input type="hidden" name="itemoption" value="<%= fitemoption1 %>">
	  <tr bgcolor="#FFFFFF">
	<td rowspan=5><input type="hidden" name="mode"><img src="<%= fsmallimage1 %>" width="100" height="100"></td>
	<td><font size=2>��������ȣ :</font></td>
	 <td><font size=2><%= idx %></font><input type="hidden" name="idx" value="<%= idx %>"></td>
	<td><font size=2>������ �ɼ� : </font></td>
	<td><font size=2><%= fitemoption1 %></font>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td><font size=2>��ǰ��ȣ :</font></td> 
	<td><font size=2><%= fitemid1 %></font></td>
	<td><font size=2>�۾������� :</font></td> 
	<td><font size=2><%= freguserid1 %></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td><font size=2>��ǰ�� : </font></td>
	<td><font size=2><%= fitemname1 %></font></td>
	<td><font size=2>�귣�� :</font></td>
	<td><font size=2><%= fmakerid1 %></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td><font size=2>��ǰ���� : </font></td> 
	<td><font size=2>
		<% if fitemgubun1 = 10 then %>
		�¶��λ�ǰ
		<% elseif fitemgubun1 = 90 then %>
			�������λ�ǰ
		<% elseif fitemgubun1 = 70 then %>
			�Ҹ�ǰ
		<% end if %>
	</font></td>
	<td></td>
	<td></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td><font size=2>����ľǽ���� : </font></td>
	<td><font size=2><%= fbasicstock1 %></font><input type="hidden" name="basicstock" value="<%= fbasicstock1 %>"></td>
	<td></td>
	<td></td>
	</tr>
	</table>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
	<td><font size=2>�ǻ���� : </font> <input type="text" name="jaego" size="12" value="<%= frealstock1 %>"></td>
	<td><input type="button" value="����" onclick="javascript:sendit()"></tr>
	</tr>
	</form>
	</table>
	<!--��ǰ���̺�-->
	
	<!-- ǥ �ϴܹ� ����-->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	    <tr valign="top" height="25">
	        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="bottom" align="right">&nbsp;</td>
	        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	    </tr>
	    <tr valign="bottom" height="10">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_08.gif"></td>
	        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	    </tr>
	</table>
	</body>
	<!-- ǥ �ϴܹ� ��-->
	<!--������� ��-->

<!--����ľǸ�����-->
<% elseif fmode = "" then %>					
	
	<%	 
	dim oip1 ,i					'Ŭ��������
		set oip1 = new Cfitemlist	'������ ��Ż�� �ֱ�
		oip1.Frectitemid = itemid
		if itemoption = "" then
			oip1.frectitemoption = "0000"
		else
			oip1.frectitemoption = itemoption		
		end if
		oip1.fjaegoinsert()			'Ŭ�������� 			
	%>	
	
	<script language="javascript">
	function sendit()
	{
	document.form1.submit();
	}
	</script>
	
	<!--ǥ ������-->
	<body>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	 <form method="get" name="frm" action=""> 
		<tr height="10" valign="bottom">
			<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
			<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
			<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
		</tr>
		<tr height="25" valign="top">
			<td background="/images/tbl_blue_round_04.gif"></td>
			<td background="/images/tbl_blue_round_06.gif">
				<img src="/images/icon_star.gif" align="absbottom">
				<font color="red" size=2><strong>����ľ��Է� </strong> </font>
				</td>
				
			<td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
		<tr valign="top">
			<td background="/images/tbl_blue_round_04.gif"></td>
			<td><br>
			��ǰ�ڵ� : <input type="text" name="itemid" value="<%= itemid %>" size="10">
			<a href="javascript:frm.submit();">
			<img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
			</td>
			<td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
		<tr height="10" valign="top">
			<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
			<td background="/images/tbl_blue_round_06.gif"></td>
			<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
		</tr>
		</tr>
		</form>
	</table>
	<!--ǥ ��峡-->
	
	<% if oip1.ftotalcount > 0 then %>
	<!--��ǰ���̺����-->
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	 <form method="post" name="form1" action="jaegototalsubmit.asp">  	
	  <tr bgcolor="#FFFFFF">
 	  
		<td rowspan=3><img src="<%= oip1.flist(i).fsmallimage %>" width="100" height="100"></td></td>
		<td><font size=2>��ǰ��ȣ :</font></td>
		<td><font size=2><%= oip1.flist(i).fitemid %><input type="hidden" name="itemid" value="<%= oip1.flist(i).fitemid %>"></font></td>			
		<td><font size=2>������ �ɼ� : </font></td>
		<td><font size=2><%= oip1.flist(i).fitemoption %><input type="hidden" name="itemoption" value="<%= oip1.flist(i).fitemoption %>"></font></td>			
		</tr>
		<tr bgcolor="#FFFFFF">
		<td><font size=2>��ǰ�� : </font></td>
		<td><font size=2><%= oip1.flist(i).fitemname %></font></td>
		<td><font size=2>�귣�� :</font></td>
		<td><font size=2><%= oip1.flist(i).fmakerid %></font></td>
		</tr>
		<tr bgcolor="#FFFFFF">
		<td><font size=2>����ľǿ���� : </font></td>
		<td><font size=2><%= oip1.flist(i).frealstock %></font><input type="hidden" name="basicstock" value="<%= oip1.flist(i).frealstock %>"></td>		<td><font size=2>��ǰ���� : </font></td> 
		<td><font size=2>
			<% if oip1.flist(i).fitemgubun = 10 then %>
				�¶��λ�ǰ
			<% elseif oip1.flist(i).fitemgubun = 90 then %>
				�������λ�ǰ
			<% elseif oip1.flist(i).fitemgubun = 70 then %>
				�Ҹ�ǰ
			<% else %>
				<%= oip1.flist(i).fitemgubun %>
			<% end if %>
		</font></td>
		</tr>
	</table>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
	<td><font size=2 colspan=4>����ľǼ��� : <input type="text" name="jaego" size="12"> <input type="button" value="����" onclick="javascript:sendit()"> </font> </tr>
	</tr>
	</form>
	</table>
	<!--��ǰ���̺�-->
	
	<% else%>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
	<td align=center>[ �˻������ �����ϴ�. ]</td></tr>
	</table>	
	<% end if %>
	<!-- ǥ �ϴܹ� ����-->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	    <tr valign="top" height="25">
	        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="bottom" align="right">&nbsp;</td>
	        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	    </tr>
	    <tr valign="bottom" height="10">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_08.gif"></td>
	        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	    </tr>
	</table>
	</body>
	<!-- ǥ �ϴܹ� ��-->
	<!--����ľǸ�峡-->	

<% end if %>

<script language='javascript'>
function GetOnLoad(){
    document.form1.jaego.focus();
    document.form1.jaego.select();
}

window.onload = GetOnLoad;
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->