<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ��������Ʈ
' History : 2011.02.24 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<%
Dim clsedms, arrList, intLoop
Dim icateidx1, icateidx2
Dim sedmsname,blnUsing
Dim iTotCnt,iPageSize, iTotalPage,page

	iPageSize = 20
	page = requestCheckvar(Request("page"),10)
	if page="" then page=1

	icateidx1 = requestCheckvar(Request("selC1"),10)
	icateidx2 = requestCheckvar(Request("hidC2"),10)

	if icateidx1 = "" then icateidx1 = 0
	if icateidx2 = "" then icateidx2= 0
		
	sedmsname = 	requestCheckvar(Request("sen"),20)
	blnUsing= requestCheckvar(Request("selU"),1)
	
Set clsedms = new Cedms
	clsedms.Fcateidx1 	= icateidx1
	clsedms.Fcateidx2	= icateidx2
	clsedms.Fedmsname		= sedmsname
	clsedms.FisUsing    = blnUsing
	clsedms.FCurrPage 	= page
	clsedms.FPageSize 	= iPageSize 
	arrList = clsedms.fnGetEdmsList
	iTotCnt = clsedms.FTotCnt

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script type="text/javascript" src="/js/ajax.js"></script>
<script language="javascript">
<!--
// ������ �̵�
function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}


//���ε��
function jsNewReg(){
	var winD = window.open("popedmsConts.asp","popD","width=880, height=800, resizable=yes, scrollbars=yes");
	winD.focus();
}
//����
function jsModReg(edmsidx){
	var winD = window.open("popedmsConts.asp?ieidx="+edmsidx,"popD","width=880, height=600, resizable=yes, scrollbars=yes");
	winD.focus();
}

// ī�װ� ajax =========================================================================================================
    initializeReturnFunction("processAjax()");
    initializeErrorFunction("onErrorAjax()");

    var _divName = "CL";

    function processAjax(){
        var reTxt = xmlHttp.responseText;
        eval("document.all.div"+_divName).innerHTML = reTxt;
    }

    function onErrorAjax() {
            alert("ERROR : " + xmlHttp.status);
    }

    //������ ī�װ��� ���� ���� ī�װ� ����Ʈ �������� Ajax
    function jsSetCategory(sMode){
      var ipcidx  = document.frm.selC1.value;
      var icidx   = $("#selC2").val();

        initializeURL('ajaxCategory.asp?sMode='+sMode+'&ipcidx='+ipcidx+'&icidx='+icidx);
    	startRequest();
    }

    //���� �ٿ�ε�
    function jsDownload(ieidx, sRFN, sFN){
    var winFD = window.open("<%=uploadImgUrl%>/linkweb/edms/procDownload.asp?ieidx="+ieidx+"&sRFN="+sRFN+"&sFN="+sFN,"popFD","");
    winFD.focus();
    }

    //����÷��
	function jsAttachFile(ieidx){
	var winAF = window.open("popRegFile.asp?ieidx="+ieidx+"&iML=10&menupos=<%=menupos%>&page=<%=page%>&icateidx1=<%=icateidx1%>&icateidx2=<%=icateidx2%>","popAF","width=450, height=200, resizable=yes, scrollbars=yes");
	winAF.focus();
	}

 	 //���ϻ���
	function jsDeleteFile(ieidx){
	if (confirm("��������� �����Ͻðڽ��ϱ�?")){
		document.frmDel.ieidx.value = ieidx;
		document.frmDel.submit();
	}
	}

	//�˻�
	function jsSearch(){
		document.frm.hidC2.value = $("#selC2").val();
		document.frm.submit();
	}

	//������ ���
	function jsAddForm(ieidx){
		var winAF = window.open("popEdmsForm.asp?ieidx="+ieidx,"popAFo","width=880, height=600, resizable=yes, scrollbars=yes");
	winAF.focus();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<form name="frmDel" method="post" action="procEdms.asp">
<input type="hidden" name="hidM" value="A">
<input type="hidden" name="ieidx" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="icateidx1" value="<%=icateidx1%>">
<input type="hidden" name="icateidx2" value="<%=icateidx2%>">
</form>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="page" value="">
			<input type="hidden" name="hidC2" value="<%=icateidx2%>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					�� ī�װ� :
					<select name="selC1" id="selC1" onChange="jsSetCategory('CL')">
					<option value="0">--�ֻ���--</option>
					<%clsedms.sbGetOptedmsCategory 1,0,icateidx1 %>
					</select>

					�� ī�װ� :
					<span id="divCL">
					<select name="selC2" id="selC2">
					<option value="0">----</option>
				<% 	IF icateidx1 > 0 THEN	'��ī�װ� ���� �� ��ī�װ� ���ð����ϰ�
						clsedms.sbGetOptedmsCategory 2,icateidx1,icateidx2
					END IF
				%>
					</select>
					</span>
					
					������:<input type="text" name="sen"  size="20" maxlength="64" value="<%=sedmsname%>">
					
					�������:
					<select name="selU">
						<option value="">--</option> 
						<option value="1" <%IF blnUsing="1" THEN%>selected<%END IF%>>Y</option>
						<option value="0" <%IF blnUsing="0" THEN%>selected<%END IF%>>N</option>
					</select>
					
				</td>	
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<%Set clsedms = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<tr>
	<td><input type="button" class="button" value="�űԵ��" onClick="jsNewReg();"></td>
</tr>
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="16">
					�˻���� : <b><%=iTotCnt%></b> &nbsp;
					������ : <b><%= page %> / <%=iTotalPage%></b>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>idx</td>
				<td>�����ڵ�</td>
				<td>��ī�װ�</td>
				<td>��ī�װ�</td>
				<td>�Ϸù�ȣ</td>
				<td>������</td>
				<td>ǥ�ü���</td>
				<td>��������</td>
				<td>���ڰ��翩��</td>
				<td>������</td>
				<td>����������</td>
				<!-- td>CFO����</td -->
				<td>���ο��Ῡ��</td>
				<td>������û���������</td>
				<td>�������</td>
				<td>������</td>
				<td>�������</td>
			</tr>
			<% Dim sFileName, sReFileName
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
					IF arrList(9,intLoop) <> "" THEN
					sFileName = split(arrList(9,intLoop),"/")(Ubound(split(arrList(9,intLoop),"/")))
					sReFileName = arrList(7,intLoop)&"_"&arrList(6,intLoop)&"."&split(arrList(9,intLoop),".")(ubound(split(arrList(9,intLoop),".")))
					END IF
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">
				<td><%=arrList(0,intLoop)%></td>
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(7,intLoop)%></a></td>
				<td><%=arrList(2,intLoop)%></td>
				<td><%=arrList(4,intLoop)%></td>
				<td><%=arrList(5,intLoop)%></td>
				<td><%=arrList(6,intLoop)%></td>
				<td><%=arrList(8,intLoop)%></td>
				<td><%IF arrList(10,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><%IF arrList(11,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><%If arrList(21,intLoop) = "Y" THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><%=arrList(16,intLoop)%></td>
				<!-- td><%IF arrList(20,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td -->
				<td><%IF (arrList(13,intLoop) <> "") and (not isNull(arrList(13,intLoop)))  THEN %><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><%IF arrList(17,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><%IF arrList(18,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><input type="button" class="button" value="<%IF isNull(arrList(19,intLoop)) or arrList(19,intLoop)="" THEN %>���<%ELSE%>����<%END IF%>" onClick="jsAddForm('<%=arrList(0,intLoop)%>');"></td>
				<td><%IF arrList(9,intLoop) <> "" THEN%><a href="javascript:jsDownload('<%=arrList(0,intLoop)%>','<%=sReFileName%>','<%=sFileName%>');"><%=sReFileName%></a>&nbsp;&nbsp;<a href="javascript:jsDeleteFile(<%=arrList(0,intLoop)%>);"><img src="/images/icon_minus.gif" border="0" alt="���ϻ���"></a> <%END IF%><A href="javascript:jsAttachFile(<%=arrList(0,intLoop)%>);"><img src="/images/icon_plus.gif" border="0" alt="����÷�� - �������� ���� �� �� ���� �߰�"></a></td>
			</tr>
		<%	Next
			ELSE%>
			<tr height=30 align="center" bgcolor="#FFFFFF">
				<td colspan="17">��ϵ� ������ �����ϴ�.</td>
			</tr>
			<%END IF%>
		</table>
	</td>
</tr>
<!-- ������ ���� -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10

		iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

		If (page mod iPerCnt) = 0 Then
			iEndPage = page
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(page) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>
					    </tr>
					</table>
				</td>
			</tr>
</table>
<!-- ������ �� -->
</body>
</html>




