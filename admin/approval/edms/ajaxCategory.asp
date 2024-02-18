<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<%	
Dim clsedms
 Dim ipcidx, icidx, iedmsidx,smode,iserialNum
 smode	=  requestCheckvar(Request("smode"),2)   
 ipcidx =  requestCheckvar(Request("ipcidx"),10) 
 icidx 	=  requestCheckvar(Request("icidx"),10) 
 iedmsidx=  requestCheckvar(Request("ieidx"),10)  
Set clsedms = new Cedms
IF smode = "CL" THEN 
%> 
<select name="selC2" id="selC2">
<option value="0">전체</option>
<%
 clsedms.sbGetOptedmsCategory 2,ipcidx,icidx 
 %>
</select> 
<%
ELSEIF smode = "C2" THEN 
%>
<select name="selC2"  id="selC2" onChange="jsSetCategory('SN');">
<option value="0">전체</option>
<%
 clsedms.sbGetOptedmsCategory 2,ipcidx,icidx 
 %>
</select>
<%
ELSEIF smode = "CM" THEN   
%>중카테고리:
<select name="selC2"  id="selC2" onChange="jsSetCategory('CD');">
<option value="0">전체</option>
<%
 clsedms.sbGetOptedmsCategory 2,ipcidx,icidx 
 %>
</select> 
<%
ELSEIF smode = "CD" THEN 
%>문서명:
<select name="selC3"  id="selC3">
<option value="0">전체</option>
<% 	IF ipcidx > 0 and icidx>0 THEN	
	clsedms.FCateIdx1 = ipcidx
	clsedms.FCateIdx2 = icidx 
	clsedms.Fedmsidx = iedmsidx
	clsedms.sbOptPayEdmsList 
	END IF
%>
</select> 
<%
ELSEIF	smode="SN" THEN  
	 if ipcidx > 0 and icidx > 0 then
		clsedms.Fcateidx1   =ipcidx   
		clsedms.Fcateidx2   =icidx   
		iserialNum = Format00(3,clsedms.fnGetSerialNum)
	 end if
%>
<input type="text" name="sSN"  id="sSN" size="3" maxlenght="3" value="<%=iserialNum%>" onkeyup="jsSetSDC();">
<% 
 END IF
  
Set clsedms = nothing
%> 
<!-- #include virtual="/lib/db/dbclose.asp" -->
 