<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��ǰ���
' History : 
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/partner/incSessionDesigner.asp" --> 

<!-- #include virtual="/partner/lib/adminHead.asp" -->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<!-- #include virtual="/lib/classes/event/eventPartnerCls.asp"-->
<!-- #include virtual="/partner/lib/function/incPageFunction.asp" -->
<%
dim evtCode
evtCode =    requestCheckVar(Request("eC"),10)

if evtCode = "" then
		Call Alert_close ("���԰�ο� ������ ������ϴ�.  ")
end if

dim evtGCode
evtGCode =   requestCheckVar(Request("eGC"),10)
dim menupos 
 
dim ClsEvt, arrList , intLoop,iTotCnt
dim rectestate, realecode, realestate 

set ClsEvt = new CEvent
ClsEvt.FRectECode = evtCode 
ClsEvt.fnGetTotState
rectestate = ClsEvt.FRectevtstate
realecode = ClsEvt.FRealECode
realestate = ClsEvt.FRealEState

if rectestate > 5 then '���λ��� ���Ŀ��� ���ε�񿡼� ��������
	evtCode = realecode
end if

ClsEvt.FMakerid = session("ssBctID")
ClsEvt.FevtCode = evtCode
ClsEvt.FevtGCode = evtGCode
arrList = ClsEvt.fnGetEventGroupItem
iTotCnt = ClsEvt.FTotCnt
set ClsEvt = nothing

%>
 
</head>
<body>	
<div class="popWinV17 scrl">
	<h1>��ǰ ����</h1>
	<h2 style="margin-left:-1px;">PC ���� ����(<%=iTotCnt%>��)</h2>
		<div class="cont pad20">
			
			<div class="rt">
				<% dim eImgSize
				if isArray(arrList) Then
					eImgSize = arrList(3,0)
				Else
					eImgSize =240
				end if
					%> 
				<select class="formSlt" id="eImgSize" name="eImgSize" title="���� ����">
					<option value="240" <%if eImgSize ="240" then%>selected<%end if%>>4���� ��ǰ ����</option>
					<option value="150" <%if eImgSize ="150" then%>selected<%end if%>>5���� ��ǰ ����</option>
				</select>
			</div>
			<div class="tbListWrap tMar10">
				<ul class="thDataList">
					<li> 
						<p class="cell15">��ǰ ID</p>
						<p>��ǰ��</p>
						<p class="cell20">�ǸŰ���</p>
						<p class="cell12">���ļ���</p>
						<p class="cell10">�Ǹſ���</p>
					</li>
				</ul>
				<div id="sitem">
				<ul id="sortable" class="tbDataList">
					<%if isArray(arrList) Then
							for intLoop = 0 To ubound(arrList,2)
						%>
					<li> 
						<p class="cell15"><%=arrList(0,intLoop)%></p>
						<p class="lt"><span><%=arrList(1,intLoop)%></span></p>
						<p class="cell20"><%=FormatNumber(arrList(7,intLoop),0)%>
									<%
										'���ΰ�
												if arrList(4,intLoop)="Y" then
													Response.Write "<br><font color=#F08050>("&CLng((arrList(7,intLoop)-arrList(9,intLoop))/arrList(7,intLoop)*100) & "%��)" & FormatNumber(arrList(9,intLoop),0) & "</font>"
												end if
												'������
												if arrList(11,intLoop)="Y" then
													Select Case arrList(12,intLoop)
														Case "1"
															Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(5,intLoop)-(CLng(arrList(13,intLoop)*arrList(5,intLoop)/100)),0) & "</font>"
														Case "2"
															Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(5,intLoop)-arrList(13,intLoop),0) & "</font>"
													end Select
												end if
									%>
								</p>
						<p class="cell12"><%=arrList(2,intLoop)%> </p>
						<p class="cell10">
								<%if arrList(14,intLoop) ="Y" then%>
									<span class="cBl1">�Ǹ���</span>
								<%elseif arrList(14,intLoop) ="S" then%>
								<span class="cRd1">�Ͻ�ǰ��</span>
								<%elseif arrList(14,intLoop) ="N" then%>
								<span class="cRd1">�Ǹž���</span>
								<%end if%>	
						</p>
					</li>
				<%	next
						end if
				%>	
				 
				</ul>
			</div>

		</div>

	 
</div>
  
</body>
</html> 