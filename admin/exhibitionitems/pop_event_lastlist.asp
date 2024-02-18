<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/pop_event_lastlist.asp
' Description :  ���� �̺�Ʈ ���� ��������
' History : 2007.03.20 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls_V3.asp"-->
<%
	'��������
	Dim menupos
	Dim cEvtList	
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
	Dim sEvt,strTxt,sKind, sdate, title, titlestr
	dim blnWeb, blnMobile, blnApp, isWeb, isMobile, isApp, num
	
	menupos = request("menupos")
	sKind 	= request("eventkind")
	sEvt 	= request.form("selEvt")
	strTxt 	= request.form("sEtxt")
	blnWeb		= requestCheckVar(Request("isWeb"),1)
	blnMobile	= requestCheckVar(Request("isMobile"),1)
	blnApp		= requestCheckVar(Request("isApp"),1)

	num		= requestCheckVar(Request("num"),1)
	
	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	iCurrpage = Request("iC")	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
		
	'������ ��������
	set cEvtList = new ClsEvent	
		cEvtList.FCPage = iCurrpage	'����������
		cEvtList.FPSize = iPageSize '���������� ���̴� ���ڵ尹�� 
		
		cEvtList.FIsWeb = blnWeb
		cEvtList.FIsMobile = blnMobile
		cEvtList.FIsApp = blnApp
		cEvtList.FSKind = sKind '�˻� ����
		cEvtList.FSfEvt = sEvt '�˻� �̺�Ʈ�� or �̺�Ʈ�ڵ�
		cEvtList.FSeTxt = strTxt '�˻���	
		
 		arrList = cEvtList.fnGetEventLastList2 	'�����͸�� ��������
 		iTotCnt = cEvtList.FTotCnt	'��ü ������  ��
 	set cEvtList = nothing
 	 	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
	Dim arreventkind, arreventstate
	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	arreventkind = fnSetCommonCodeArr("eventkind",False)
	arreventstate= fnSetCommonCodeArr("eventstate",False)	
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--
	//�˻�
	function jsSearch(){	 
	 if(document.frmLast.selEvt.options[document.frmLast.selEvt.selectedIndex].value == "evt_code" &&  document.frmLast.sEtxt.value !="") {
	   if(!IsDigit(document.frmLast.sEtxt.value)){
	    alert("�̺�Ʈ �ڵ�� ���ڸ� �Է°����մϴ�.");
	    document.frmLast.sEtxt.focus();
	    return false;
	   }
	 }		
	}
	
	//�θ�â�� �� �ѱ��
	function jsSetEvtCont(cd,title,subtitle,sale,csale,sdate,edate){
        $("#infomenu",opener.document).show();
		$("#copyDateButton",opener.document).show();
        opener.document.frmreg.evt_code.value = cd;
        $("#evt_name",opener.document).text(title);
        $("#evt_startdate",opener.document).text(sdate);
        $("#evt_enddate",opener.document).text(edate);
        if (sale == '') {
            $("#evt_saleper",opener.document).text("���� ������ �����ϴ�.");
        } else {
            $("#evt_saleper",opener.document).text(sale);
        }
        
        if (csale == '') {
            $("#evt_salecoupon",opener.document).text("�������� ������ �����ϴ�.");
        } else {
            $("#evt_salecoupon",opener.document).text(csale);
        }
        

		window.close();
	}
	
	//�������̵�
	function jsGoPage(iP){
		document.frmLast.iC.value = iP;
		document.frmLast.submit();
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ���� �̺�Ʈ ����Ʈ </div>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="0" >
<form name="frmLast" method="post" action="pop_event_lastlist.asp" onSubmit="return jsSearch();">
<input type="hidden" name="iC">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="pTarget" value="<%=request("pTarget")%>">
<tr>
	<td>
	     ä��:<input type="checkbox" name="isWeb" value="1" <%if blnWeb="1" then%>checked<%end if%>>PC-Web
			<input type="checkbox" name="isMobile"  value="1" <%if blnMobile="1" then%>checked<%end if%>>Mobile
			<input type="checkbox" name="isApp"  value="1" <%if blnApp="1" then%>checked<%end if%>>App&nbsp;&nbsp;&nbsp;
		����: <%sbGetOptEventCodeValue "eventkind",sKind,True,"onChange=""document.frmLast.submit();"""%>&nbsp;&nbsp;&nbsp;
		�ڵ�/��: <select name="selEvt"> 
        			<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
        			<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>�̺�Ʈ��</option>
        			</select>
        			<input type="text" name="sEtxt" size="15" value="<%=strTxt%>">
        			 <input type="image" src="/images/icon_search.jpg">
    </td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" >
		     <td align="center" width="15%">ä��</td>
			<td align="center" width="10%">�ڵ�</td>	
			<td align="center" width="15%">����</td>
			<td align="center">�̺�Ʈ��</td>	
			<td align="center" width="15%">����</td>	
		</tr>
		<%IF isArray(arrList) THEN 
			For intLoop = 0 To UBound(arrList,2)
			isWeb = False
		    isMobile = False
		    isApp = False
		
		IF isNull(arrList(9,intLoop)) and isNull(arrList(10,intLoop)) and isNull(arrList(11,intLoop)) then
			if arrList(1,intLoop) = "19" THEN
				isWeb = False
				isMobile = True
				isApp = True
			ELSEIF arrList(1,intLoop) = "25"  THEN
				isWeb = False
				isMobile = False
				isApp = True
			ELSEIF arrList(1,intLoop) = "26"  THEN	
				isWeb = False
				isMobile = True
				isApp = False
			ELSE
				isWeb = True
				isMobile = False
				isApp = False	
			END IF
		END IF	
		IF 	 not isNull(arrList(9,intLoop))  THEN	
			isWeb = arrList(9,intLoop)
		END IF	
		IF 	 not isNull(arrList(10,intLoop)) THEN
			 isMobile = arrList(10,intLoop)
		END IF	 
		IF 	 not isNull(arrList(11,intLoop)) THEN
			isApp = arrList(11,intLoop)	
		END If
		
			If arrList(5,intLoop) < now() Then
				sdate = FormatDate(now(),"0000-00-00")
			Else
				sdate = arrList(5,intLoop)
			End If

			If InStr(arrList(4,intLoop),"|")>0 Then
				titlestr = split(db2html(arrList(4,intLoop)),"|")
				title=titlestr(0)
			Else
				title=db2html(arrList(4,intLoop))
			End If

			%>
		<tr bgcolor="#FFFFFF" onClick="jsSetEvtCont(<%=arrList(0,intLoop)%>,'<%=title%>','<%=db2html(arrList(14,intLoop))%>','<%=db2html(arrList(12,intLoop))%>','<%=db2html(arrList(13,intLoop))%>','<%=sdate%>','<%=FormatDate(arrList(6,intLoop),"0000-00-00")%>');" style="cursor:hand;" onMouseOver="this.style.backgroundColor='#FFFFEC'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
		    <td> <%IF isWeb THEN %>  Web <%END IF%><%IF isMobile THEN %>&nbsp; <font color="blue">Mobile</font> <%END IF%><%IF isApp THEN %>&nbsp;<font color="red">App</font><%END IF%></td>
			<td  align="center"><%=arrList(0,intLoop)%></td>
		
			<td  align="center"><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></td>
	
			<td><%=db2html(arrList(4,intLoop))%></td>
	
			<td  align="center"><%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%></td>
		</tr>
		<% Next %>
		<%ELSE%>
		<tr><td colspan="4"  bgcolor="#FFFFFF" align="center">��ϵ� ������ �����ϴ�.</td></tr>
		<%END IF%>
		</table>	
</tr>
<tr>
	<td>
		<!-- ����¡ó�� -->
<%		
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	

If (iCurrpage mod iPerCnt) = 0 Then																
	iEndPage = iCurrpage
Else								
	iEndPage = iStartPage + (iPerCnt-1)
End If	
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="bottom" height="25">       
        <td valign="bottom" align="center">
         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
		<% else %>[pre]<% end if %>
        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(iCurrpage) then
		%>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
		<%		else %>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
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
</form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->  