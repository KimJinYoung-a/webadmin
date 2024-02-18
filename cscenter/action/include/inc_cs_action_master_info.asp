<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'			2016.06.30 한용민 수정
'###########################################################
%>
<% if (IsDisplayCSMaster = true) then %>
	<%
	dim jupsugubun, jupsudefaulttitle

	jupsugubun = GetCSCommName("Z001", divcd)
	jupsudefaulttitle = GetDefaultTitle(divcd, id, orderserial)

	'CS 접수시에는 반품접수(업체배송)/회수신청(텐바이텐배송) 을 구분하지 않고
	'저장시 브랜드지정이 있는경우 반품접수(업체배송), 없는경우 회수신청(텐바이텐배송) 으로 저장한다.
	if (IsStatusRegister = true) and (divcd = "A004" or divcd = "A010") then
		jupsugubun = "반품접수"
		jupsudefaulttitle = "반품접수"
	end if
	%>
	<tr >
	    <td >
	        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	        <tr>
	            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">접수구분</td>
	            <td bgcolor="#FFFFFF">
				    	<font style='line-height:100%; font-size:15px; color:blue; font-family:돋움; font-weight:bold'><%= jupsugubun %></font>
				    	&nbsp;
	                <% if (Not IsStatusRegister) then %>
				    	<font style='line-height:100%; font-size:15px; color:#CC3333; font-family:돋움; font-weight:bold'>[<%= ocsaslist.FOneItem.GetCurrstateName %>]</font>
				    	<% if ocsaslist.FOneITem.FDeleteyn<>"N" then %>
							<font style='line-height:100%; font-size:15px; color:#FF0000; font-family:돋움; font-weight:bold'>- 삭제된 내역</font>
				    	<% end if %>
			    	<% end if %>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">주문번호</td>
	            <td bgcolor="#FFFFFF" width="200" >
	                <%= orderserial %>
	                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font>]
	                [<font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font>]
	            </td>
	        </tr>
	        <tr height="20">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수자</td>
	            <td bgcolor="#FFFFFF" >
	                <% if (IsStatusRegister) then %>
	                    <%= session("ssbctid") %>
	                <% else %>
	                    <%= ocsaslist.FOneItem.Fwriteuser %>
	                <% end if %>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">주문자ID</td>
	            <td bgcolor="#FFFFFF">
	                <%= oordermaster.FOneItem.FUserID %>
	                (<font color="<%= getUserLevelColorByDate(oordermaster.FOneItem.fUserLevel, left(oordermaster.FOneItem.Fregdate,10)) %>">
					<%= getUserLevelStrByDate(oordermaster.FOneItem.fUserLevel, left(oordermaster.FOneItem.Fregdate,10)) %></font>)
	            </td>
	        </tr>
	        <tr height="20">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수일시</td>
	            <td bgcolor="#FFFFFF" >
	                <% if (IsStatusRegister) then %>
	                	<%= now() %>
	                <% else %>
	                	<%= ocsaslist.FOneItem.Fregdate %>
	                <% end if %>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">주문자정보</td>
	            <td bgcolor="#FFFFFF">
	                <%= oordermaster.FOneItem.FBuyname %>
	                 &nbsp;
	                 [<%= oordermaster.FOneItem.FBuyHp %>]
	            </td>
	        </tr>
	        <tr height="20">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수제목</td>
	            <td bgcolor="#FFFFFF" >
	                <% if (IsStatusRegister) then %>
						<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= jupsudefaulttitle %>" size="56" maxlength="56">
						<% SelectBoxCSTemplateGubunNew "30", "csreg_template", "" %>
						<iframe name="CSTemplateFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
	                <% else %>
	                	<input <% if IsStatusFinishing then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= ocsaslist.FOneItem.Ftitle %>" size="56" maxlength="56">
	                <% end if %>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">수령인정보</td>
	            <td bgcolor="#FFFFFF">
	                 <%= oordermaster.FOneItem.FReqName %>
	                 &nbsp;
	                 [<%= oordermaster.FOneItem.FReqHp %>]
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">사유구분</td>
	            <td bgcolor="#FFFFFF">
	                <input type="hidden" name="gubun01" value="<%= ocsaslist.FOneItem.Fgubun01 %>">
	                <input type="hidden" name="gubun02" value="<%= ocsaslist.FOneItem.Fgubun02 %>">
	                <input class="text_ro" type="text" name="gubun01name" value="<%= ocsaslist.FOneItem.Fgubun01name %>" size="16" Readonly >
	                &gt;
	                <input class="text_ro" type="text" name="gubun02name" value="<%= ocsaslist.FOneItem.Fgubun02name %>" size="16" Readonly >
	                <input class="csbutton" type="button" value="선택" onClick="divCsAsGubunSelect(frmaction.gubun01.value, frmaction.gubun02.value, frmaction.gubun01.name, frmaction.gubun02.name, frmaction.gubun01name.name, frmaction.gubun02name.name,'frmaction','causepop');">
	                <div id="causepop" style="position:absolute;"></div>

	                <!-- 일부 사유 미리 표시 -->
	                <%
	                '참조쿼리
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
		                [<a href="javascript:selectGubun('C004','CD01','공통','고객변심','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">고객변심</a>]
		                [<a href="javascript:selectGubun('C004','CD05','공통','품절','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">품절</a>]
						<!--
						[<a href="javascript:selectGubun('C005','CE02','상품관련','상품불만족','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품불만족</a>]
						-->
						[<a href="javascript:selectGubun('C006','CF06','물류관련','출고지연','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">출고지연</a>]
		                [<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">기타</a>]
		                <% if IsStatusRegister then %>
		                	&nbsp; &nbsp; &nbsp;
		                	<div id="chkmodifyitemstockoutyn" style="display: inline;"><input type="checkbox" name="modifyitemstockoutyn" value="Y" checked> 품절정보 저장(업배상품)</div>
		                <% end if %>

	                <% elseif (ocsaslist.FOneItem.IsReturnProcess) then %>
		                [<a href="javascript:selectGubun('C004','CD01','공통','고객변심','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">고객변심</a>]
		                [<a href="javascript:selectGubun('C005','CE01','상품관련','상품불량','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품불량</a>]
		                [<a href="javascript:selectGubun('C006','CF01','물류관련','오발송','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">오배송</a>]
						<!--
	                    [<a href="javascript:selectGubun('C004','CD04','공통','사이즈교환','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">사이즈교환</a>]
						-->
	                    [<a href="javascript:selectGubun('C004','CD06','공통','사이즈 안맞음','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">사이즈 안맞음(고객변심)</a>]
						[<a href="javascript:selectGubun('C006','CF06','물류관련','출고지연','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">출고지연</a>]
					<% elseif (divcd="A009") or (divcd="A006") or (divcd="A700") or (divcd="A900") then %>
						<% if (divcd="A700") then %>
							[<a href="javascript:selectGubun('C004','CD10','공통','업체반품불가','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">업체반품불가</a>]
						<% end if %>
	                	[<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">기타</a>]
					<% elseif (divcd="A060") then %>
	                	[<a href="javascript:selectGubun('C011','CK01','긴급문의','취소문의','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">취소문의</a>]
	                	[<a href="javascript:selectGubun('C011','CK02','긴급문의','교환반품문의','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">교환반품문의</a>]
	                	[<a href="javascript:selectGubun('C011','CK03','긴급문의','AS문의','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">AS문의</a>]
	                	[<a href="javascript:selectGubun('C011','CK04','긴급문의','배송문의','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">배송문의</a>]
						[<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">기타</a>]

	                <% elseif (divcd="A001") then %>
	                	[<a href="javascript:selectGubun('C006','CF03','물류관련','구매상품누락','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품누락</a>]

	                <% elseif (divcd="A002") then %>
		                [<a href="javascript:selectGubun('C006','CF04','물류관련','사은품누락','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(물류)사은품누락</a>]
		                [<a href="javascript:selectGubun('C005','CE05','상품관련','이벤트오등록','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(MD)이벤트오등록</a>]

	                <% elseif (divcd="A000") then %>
		                [<a href="javascript:selectGubun('C004','CD08','공통','회수후출고','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">회수후출고</a>]
		                [<a href="javascript:selectGubun('C004','CD09','공통','선출고요청','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">선출고요청</a>]

		                [<a href="javascript:selectGubun('C005','CE01','상품관련','상품불량','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품불량</a>]
		                [<a href="javascript:selectGubun('C006','CF01','물류관련','오발송','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">오발송</a>]
		                [<a href="javascript:selectGubun('C006','CF02','물류관련','상품파손','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품파손</a>]
		                <!--
		                [<a href="javascript:selectGubun('C004','CD04','공통','사이즈교환','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">사이즈교환</a>]
		                -->
		                <p>
		                * <font color="red"><b>사이즈교환</b></font>은 "옵션변경 맞교환" 입력 (접수내용에 입력시 오배송할 수 있음)

	                <% elseif (divcd="A100") or (divcd="A111") then %>
	                	<!--
	                	* 고객변심, 사이즈 안맞음(고객변심) 의 경우 회수이후에 맞교환 출고한다.
	                	* 참고 : http://logics.10x10.co.kr/v2/online/m_re_chulgo.asp
	                	-->
	                	[<a href="javascript:selectGubun('C004','CD01','공통','고객변심','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">고객변심</a>]
		                [<a href="javascript:selectGubun('C005','CE01','상품관련','상품불량','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품불량</a>]
		                [<a href="javascript:selectGubun('C006','CF01','물류관련','오발송','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">오발송</a>]
		                [<a href="javascript:selectGubun('C006','CF02','물류관련','상품파손','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품파손</a>]
		                [<a href="javascript:selectGubun('C004','CD06','공통','사이즈 안맞음','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">사이즈 안맞음(고객변심)</a>]
					<% elseif (divcd="A999") then %>
						[<a href="javascript:selectGubun('C012','CL01','추가결제','반품문의','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">반품문의</a>]
	                	[<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">기타</a>]
	                <% end if %>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">결제정보</td>
	            <td bgcolor="#FFFFFF">
	            	<% if oordermaster.FOneItem.IsErrSubtotalPrice then %>
	            		<font color="red"><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice-realSubPaymentSum,0) %>원</font>
	            	<% else %>
	            		<%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice-realSubPaymentSum,0) %>원
					<% end if %>
	            	&nbsp;
	                [<%= oordermaster.FOneItem.JumunMethodName %>]

	                <% if (realdepositsum>0) then %>
	                   /&nbsp; <strong><%= FormatNumber(realdepositsum,0) %></strong>원&nbsp; [예치금]
	                <% end if %>
	                <% if (realgiftcardsum>0) then %>
	                   /&nbsp; <strong><%= FormatNumber(realgiftcardsum,0) %></strong>원&nbsp; [상품권]
	                <% end if %>


	                <% if (oordermaster.FOneItem.Faccountdiv="110") then %>
	                	(OK Cashbag사용 : <strong><%= FormatNumber(oordermaster.FOneItem.FokcashbagSpend,0) %></strong> 원)
	                <% end if %>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="6">
					접수내용<br><br>
	    			<input type="button" class="button" value="시간" onClick="WriteNowDateString(document.frmaction.contents_jupsu)">
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
									response.Write "<br>무료반품 이벤트 주문<br>"
									response.Write "&nbsp; - &nbsp; 브랜드 : " & IsTempEventAvail_Makerid & "<br>"
									if (IsTempEventAvail_Str <> "") then
										response.Write "&nbsp; - &nbsp; 적용불가 : " & IsTempEventAvail_Str & "<br>"
									else
										%>
										&nbsp; - &nbsp; <input type="button" class="button" onClick="jsCheckApplyEvent(frmaction);" value="무료반품적용"><br>
										<%
									end if
								end if
								%>
							</td>
						</tr>
					</table>
	            </td>
	            <td bgcolor="<%= adminColor("topbar") %>" align="center">배송지정보</td>
	            <td bgcolor="#FFFFFF" valign="top">
	            	[<%= oordermaster.FOneItem.FReqZipCode %>]<br>
	                <%= oordermaster.FOneItem.FReqZipAddr %><br>
	                <%= oordermaster.FOneItem.FReqAddress %>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">원송장</td>
	            <td bgcolor="#FFFFFF">
	            	<% if ocsaslist.FOneItem.IsRequireSongjangNO and ocsOrderDetail.FResultCount > 0 and (divcd = "A004" or divcd = "A010") and (Not IsStatusRegister) then %>
					<% Call drawSelectBoxDeliverCompany ("songjangdiv_tmp",ocsOrderDetail.FItemList(ocsOrderDetail.FResultCount - 1).Fsongjangdiv) %>
					<%= ocsOrderDetail.FItemList(ocsOrderDetail.FResultCount - 1).Fsongjangno %>
			        <% end if %>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">택배접수</td>
	            <td bgcolor="#FFFFFF">
	            	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
					<%
					Select Case ocsaslist.FOneItem.FsongjangRegGubun
						Case "U"
							Response.Write("텐바이텐(업체) 접수")
						Case "C"
							Response.Write("고객직접접수")
						Case "T"
							Response.Write("상담사 접수")
						Case Else
							Response.Write ocsaslist.FOneItem.FsongjangRegGubun
					End Select
					%>
			        <% end if %>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">택배접수자</td>
	            <td bgcolor="#FFFFFF">
	            	<%
					if ocsaslist.FOneItem.IsRequireSongjangNO then
						if Not IsNull(ocsaslist.FOneItem.FsongjangRegUserID) and (ocsaslist.FOneItem.FsongjangRegUserID <> "") then
							Response.Write ocsaslist.FOneItem.FsongjangRegUserID
							if (ocsaslist.FOneItem.FsongjangRegUserID = oordermaster.FOneItem.FUserID) then
								Response.Write " (고객)"
							elseif (ocsaslist.FOneItem.Frequireupche = "Y") and (ocsaslist.FOneItem.FsongjangRegUserID = ocsaslist.FOneItem.Fmakerid) then
								Response.Write " (업체)"
							end if
						end if
					end if
					%>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">예약번호</td>
	            <td bgcolor="#FFFFFF">
					<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
					<%= ocsaslist.FOneItem.FsongjangPreNo %>
					<% end if %>
	            </td>
	        </tr>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">택배정보</td>
	            <td bgcolor="#FFFFFF">
	            	<!-- 코딩 확인할것 -->
	            	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
				        <% Call drawSelectBoxDeliverCompany ("songjangdiv",ocsaslist.FOneItem.Fsongjangdiv) %>
				        <input type="text" class="text" name="songjangno" value="<%= ocsaslist.FOneItem.Fsongjangno %>" size="14" maxlength="16">
				        <% dim ifindurl : ifindurl = DeliverDivTrace(ocsaslist.FOneItem.Fsongjangdiv) %>
				        <% if (ocsaslist.FOneItem.Fsongjangdiv="24") then %>
	                		<a href="javascript:popDeliveryTrace('<%= ifindurl %>','<%= ocsaslist.FOneItem.Fsongjangno %>');">추적</a>
	                	<% else %>
				            <a href="<%= ifindurl + ocsaslist.FOneItem.Fsongjangno %>" target="_blank">추적</a>
				        <% end if %>
				        <input type="button" class="button" value="수정" onClick="changeSongjang('<%= id %>');">
			        <% end if %>
	            </td>
	        </tr>

			<% if False and InStr(",A000,A100,A001,A002,A009,A006,A012,", divcd) > 0 and Not IsStatusFinishing and Not IsUpcheConfirmState then %>
	        <tr bgcolor="#F4F4F4">
	            <td bgcolor="<%= adminColor("topbar") %>" align="center" height="25">
					완료구분
				</td>
	            <td bgcolor="#FFFFFF" colspan="3">
					<input type="radio" id="needChkYN_X" name="needChkYN" value="X" <%= CHKIIF(ocsaslist.FOneItem.FneedChkYN="X", "checked", "") %> > 업체처리완료시 즉시완료
					<input type="radio" id="needChkYN_F" name="needChkYN" value="F" <%= CHKIIF(ocsaslist.FOneItem.FneedChkYN="F", "checked", "") %> > 고객센터 확인 필요
	            </td>
	        </tr>
			<% end if %>

	        <% if (IsStatusFinishing) or (IsUpcheConfirmState) or (IsStatusFinished) then %>
		        <tr bgcolor="#F4F4F4">
		            <td bgcolor="<%= adminColor("topbar") %>" align="center">
		            	처리내용
		            	<% if (IsUpcheConfirmState) and (IsRefASExist) and (ocsaslist.FOneItem.Frequireupche = "Y") then %>
		            		<br><br>(업체출고)<br>+<br>(업체회수)
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
										response.write "<font color='red'><b>확인 후 처리(업체 등록)</b></font>"
									Case "N"
										response.write "<b>즉시완료</b>(확인불필요)"
									Case "F"
										response.write "<font color='red'><b>확인 후 처리(CS 등록)</b></font>"
									Case Else
										response.write "-"
								End Select
								%>
							</td></tr></table>
			            <% else %>
			            	<textarea class='textarea' name="contents_finish" cols="68" rows="9"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
			            <% end if %>
		            </td>
		            <td bgcolor="<%= adminColor("pink") %>" align="center">처리관련<br>고객오픈<br>내용입력</td>
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
