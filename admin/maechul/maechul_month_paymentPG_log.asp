<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����α�_PG�纰
' Hieditor : 2013.12.27 ������ ����
'			 2023.06.26 �ѿ�� ����(���� pg����, pg���̵� ������� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/maechul/incMaechulFunction.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgLogCls.asp"-->
<%
Dim sPGgubun, sPGid, iDateType
Dim intY, intM, dStartYear, dStartMonth,  dEndYear, dEndMonth
Dim clsPG, arrList, intLoop
Dim totPayReq,totrealPay,totCommPay,totJSPay
dim grpByDay

iDateType = requestCheckvar(request("selD"),4)
sPGgubun 	= requestCheckvar(request("selPGC"),60)
sPGid 		= requestCheckvar(request("selPGID"),60)
dStartYear		= requestCheckvar(request("selSY"),4)
dStartMonth		= requestCheckvar(request("selSM"),2)
dEndYear			= requestCheckvar(request("selEY"),4)
dEndMonth			= requestCheckvar(request("selEM"),2)
grpByDay			= requestCheckvar(request("grpByDay"),2)

'�⺻�� ����
IF iDateType = "" THEN iDateType = 1
IF dStartYear ="" THEN dStartYear = year(date())
IF dStartMonth ="" THEN dStartMonth = month(date())
IF dEndYear ="" THEN dEndYear = year(date())
IF dEndMonth ="" THEN dEndMonth = month(date())
'����Ʈ ��������
set clsPG = new CPGLog
 	clsPG.Fdatetype		= iDateType
 	clsPG.Fstartdate	= dStartYear&"-"&format00(2,dStartMonth)
 	clsPG.Fenddate		= dEndYear&"-"&format00(2,dEndMonth)
 	clsPG.Fpggubun		= sPGgubun
 	clsPG.Fpguserid		= sPGid
	clsPG.FRectGroupBy	= grpByDay

	arrList	= clsPG.fnGetPGLogList
set clsPG = nothing
%>
<table width="100%" align="left" cellpadding="5" cellspacing="0" class="a"   border="0">
<tr>
	<td>
		<form name="frm" method="get">
			<input type="hidden" name="menupos" value="<%= menupos %>">
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
			<td align="left">
				<select name="selD"  class="select">
					<option value="1" <%IF iDateType = 1 THEN%>selected<%END IF%>>������(ó����)</option>
					<option value="2" <%IF iDateType = 2 THEN%>selected<%END IF%>>������(������)</option>
					<option value="3" <%IF iDateType = 3 THEN%>selected<%END IF%>>ī��������</option>
					<option value="4" <%IF iDateType = 4 THEN%>selected<%END IF%>>�Աݿ�����</option>
				</select>
				:&nbsp;
				<select name="selSY" class="select">
					<%For intY = year(date()) To 2002 STEP -1 %>
					<option value="<%=intY%>" <%IF  Cstr(dStartYear)  = Cstr(intY) THEN%>selected<%END IF%>><%=intY%></option>
					<%Next%>
				</select>
				��
				<select name="selSM" class="select">
					<%For intM = 1 To 12%>
					<option value="<%=intM%>" <%IF  Cstr(dStartMonth)  = Cstr(intM) THEN%>selected<%END IF%>><%=intM%></option>
					<%Next%>
				</select>
				��
				~
				<select name="selEY" class="select">
					<%For intY = year(date()) To 2002 STEP -1 %>
					<option value="<%=intY%>" <%IF Cstr(dEndYear)  = Cstr(intY) THEN%>selected<%END IF%>><%=intY%></option>
					<%Next%>
				</select>
				��
				<select name="selEM" class="select">
					<%For intM = 1 To 12%>
					<option value="<%=intM%>" <%IF Cstr(dEndMonth)  = Cstr(intM) THEN%>selected<%END IF%>><%=intM%></option>
					<%Next%>
				</select>
				��
			</td>
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();"></td>
		</tr>
		<tr  bgcolor="#FFFFFF">
			<td>
				PG��:&nbsp;
				<select name="selPGC" class="select">
					<option value="">--����--</option>
					<%Call sbGetOptPGgubun(sPGgubun)%>
				</select>
				&nbsp;&nbsp;
					PG��ID:&nbsp;
				<select name="selPGID" class="select">
					<option value="">--����--</option>
					<%Call sbGetOptPGID(sPGid)%>
				</select>
				<% 'Call DrawSelectBoxPGUserid("selPGID", sPGid, "") %>
				&nbsp;&nbsp;
				ǥ�ù�� :
				<select class="select" name="grpByDay">
					<option value="M" <% if (grpByDay = "M") then %>selected<% end if %> >����ǥ��</option>
					<option value="D" <% if (grpByDay = "D") then %>selected<% end if %> >�Ϻ�ǥ��</option>
					<option value="E" <% if (grpByDay = "E") then %>selected<% end if %> >������(�ֹ���ȣ)</option>
				</select>
			</td>
		</tr>
		</table>
	</form>
	</td>
</tr>
<tr>
	<td>
		* �ִ� <font color="red">200��������</font> ǥ�õ˴ϴ�.
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>���</td>
				<td>���ⱸ��</td>
				<td>PG��</td>
				<td>PG��ID</td>
				<% if (grpByDay = "E") then %>
				<td>�ֹ���ȣ</td>
				<% end if %>
				<td>������û��(��޾�)</td>
				<td><font color="red">����</font></td>
				<td>�ǽ��ξ�</td>
				<td>������</td>
				<td>�Աݿ�����</td>
				<td>���</td>
			</tr>
			<%
			totPayReq  = 0
			totrealPay = 0
			totCommPay = 0
			totJSPay	 = 0

			IF isArray(arrList) THEN
					For intLoop = 0 TO UBound(arrList,2)
				%>
			<tr align="center" bgcolor="#ffffff">
				<td><%=arrList(0,intLoop)%></td>
				<td><%=arrList(7,intLoop)%></td>
				<td><%=arrList(1,intLoop)%></td>
				<td><%=arrList(2,intLoop)%></td>
				<% if (grpByDay = "E") then %>
				<td><%=arrList(8,intLoop)%></td>
				<% end if %>
				<td align="right"><%=formatnumber(arrList(3,intLoop),0)%></td>
				<td align="right"><font color="red"><%=formatnumber(arrList(4,intLoop)-arrList(3,intLoop),0)%></font></td>
				<td align="right"><%=formatnumber(arrList(4,intLoop),0)%></td>
				<td align="right"><%=formatnumber(arrList(5,intLoop),0)%></td>
				<td align="right"><%=formatnumber(arrList(6,intLoop),0)%></td>
				<td></td>
			</tr>
			<%	totPayReq = 	totPayReq + arrList(3,intLoop)
					totrealPay = 	totrealPay + arrList(4,intLoop)
					totCommPay = 	totCommPay + arrList(5,intLoop)
					totJSPay	 = 	totJSPay + arrList(6,intLoop)
				Next %>
			<%ELSE%>
			<tr bgcolor="#ffffff">
				<td colspan="11" align="center">��ϵ� ������ �����ϴ�.</td>
			</tr>
			<%END IF%>
			<tr  bgcolor="<%=adminColor("sky")%>" align="center">
				<td colspan="4">�հ�</td>
				<% if (grpByDay = "E") then %>
				<td></td>
				<% end if %>
				<td align="right"><%=formatnumber(totPayReq,0)%></td>
				<td align="right"><%=formatnumber(totrealPay-totPayReq,0)%></td>
				<td align="right"><%=formatnumber(totrealPay,0)%></td>
				<td align="right"><%=formatnumber(totCommPay,0)%></td>
				<td align="right"><%=formatnumber(totJSPay,0)%></td>
				<td></td>
			</tr>
		</table>
	</td>
</tr>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
