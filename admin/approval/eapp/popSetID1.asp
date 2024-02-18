<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결재라인 등록
' History : 2011.03.16 정윤정  생성
'						2013.11.28 정윤정 수정 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
	<tr>
		<td><!--조직도-->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
				<tr>
					<td>검색</td>
				</tr>	
				<tr>
					<td>
						<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
							
							<tr>
								<td></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
		<td><!--결재선-->
		</td>
	</tr>
</table>