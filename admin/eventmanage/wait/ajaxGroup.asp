<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<% 
'###########################################################
' Description : 그룹관리
' History : 2016.08.18 생성
'################################################################## 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/event/eventPartnerWaitCls.asp"-->
<%   
Dim evtCode ,makerid,evtGCode
Dim arrList, intLoop, isort, iNo
dim clsEvt 
 
evtCode = requestCheckvar(Request("eC"),10)   
evtGCode = requestCheckvar(Request("eGC"),10)   

if evtGCode ="" then evtGCode =0
set clsEvt = new CEvent
clsEvt.FevtCode = evtCode
arrList = clsEvt.fnGetEventGroup
set clsEvt = nothing
 
%>	 
 	<ul class="thDataList">
		<li> 
			<p style="width:90px">순서</p>
			<p class="">그룹명 <strong class="cRd1">*</strong></p>
			<p style="width:150px">상품 진열 <strong class="cRd1">*</strong></p>
			<p style="width:150px">관리</p>
		</li>
	</ul>
	<ul id="sortable" class="tbDataList">									 
		<% isort =0
				iNo =1
		if isArray(arrList) then 
			%> 
		<%	for intLoop = 0 To UBound(arrList,2)	 
		if Cstr(evtGCode) = Cstr(arrList(0,intLoop)) then
		%> 
		<li id="G<%=arrList(0,intLoop)%>">									
			<p style="width:90px"><%=iNo%></p>
			<p class="lt"><input type="text" class="formTxt" id="eMGD" name="eMGD" value="<%=arrList(1,intLoop)%>"  style="width:100%" maxlength="64"/></p>
			<p style="width:150px"><input type="button" id="btnItem<%=arrList(0,intLoop)%>" class="btn3 btnIntb" value="상품(<%=arrList(3,intLoop)%>)" onclick="jsSetItem('<%=arrList(0,intLoop)%>');" /></p>
			<p style="width:150px">
				<a href="javascript:jsModGSubmit('<%=arrList(0,intLoop)%>');" class="cRd1 tLine">[저장]</a> 
				<a href="javascript:jsDelGroup('<%=arrList(0,intLoop)%>');" class="cBl1 tLine">[삭제]</a>
			</p><input type="hidden" name="eMGS" value="<%=arrList(2,intLoop)%>">
			<input type="hidden" name="eMGC" value="<%=arrList(0,intLoop)%>">
		</li> 
		<%else%>
		<li id="G<%=arrList(0,intLoop)%>">									
			<p style="width:90px"><%=iNo%></p>
			<p class="lt"><%=arrList(1,intLoop)%></p>
			<p style="width:150px"><input type="button" id="btnItem<%=arrList(0,intLoop)%>" class="btn3 btnIntb" value="상품(<%=arrList(3,intLoop)%>)" onclick="jsSetItem('<%=arrList(0,intLoop)%>');" /></p>
			<p style="width:150px">
				<span id="Gbt<%=arrList(0,intLoop)%>"><a href="javascript:jsSetGList('<%=arrList(0,intLoop)%>');" class="cBl1 tLine">[수정]</a></span>
				<span><a href="javascript:jsDelGroup('<%=arrList(0,intLoop)%>');" class="cBl1 tLine">[삭제]</a></span>
			</p><input type="hidden" name="eMGS" value="<%=arrList(2,intLoop)%>">
			<input type="hidden" name="eMGC" value="<%=arrList(0,intLoop)%>">
		</li> 
	<%	 end if
			iNo = iNo+ 1
		next  
		isort = arrList(2,intLoop-1)+1
		end if%>										  
		<li class="ui-state-disabled" > 
			<p style="width:90px"  ><%=iNo%></p>
			<p class="lt"><input type="text" class="formTxt" id="eGD" name="eGD" value="" placeholder="그룹명을 입력해주세요" style="width:100%" maxlength="64"/></p>
			<p style="width:150px"><input type="button" class="btn3 btnIntb" value="상품(0)" onclick="" disabled="true" /></p>
			<p style="width:150px">
				<a href="javascript:jsAddGroup();" class="cRd1 tLine "><strong>[추가]</strong></a> 
			</p><input type="hidden" name="eGS" id="eGS" value="<%=isort%>">
		</li> 
	</ul> 
<!-- #include virtual="/lib/db/dbclose.asp" -->
 