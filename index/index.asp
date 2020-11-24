<!--
    DEVELOPMENT CONTACT INFORMATION
	Rebecca Fowler
	rebecca.fowler@redstone.army.mil
	11/4/2003
	
-->
<% 

session("bREQUESTOR") = "N"

dim access()
Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.CursorLocation = 3
			'response.write "conn = " & Session("GLIDE_Conn")
			Conn.Open Session("PWAM_Conn")
		
		
			Set RS= Server.CreateObject("ADODB.Recordset")
			sSQL = "SELECT  MIN(PWAM_AD_GROUP.ACCESS_LEVEL) AS access_level, PWAM_AD_GROUP.PWAM_DEPT_ID as pwam_dept_id "
			sSQL = sSQL & "  FROM  PWAM_AD_GROUP INNER JOIN"
			sSQL = sSQL & "  PWAM_ACCESS_LVL ON PWAM_AD_GROUP.ACCESS_LEVEL = PWAM_ACCESS_LVL.id"
			sSQL = sSQL & "  WHERE  (PWAM_AD_GROUP.AD_GROUP_NAME IN  ('" & session("pwam_ad_group_list") & "')) "
			sSQL = sSQL & " GROUP BY PWAM_AD_GROUP.PWAM_DEPT_ID"
			sSQL = sSQL & " ORDER BY PWAM_AD_GROUP.PWAM_DEPT_ID desc"
		
			'response.write "sSQL = " & sSQL
			'---------------------
  			set session("RS_PWAM_ACCESS")=server.createobject("adodb.recordset")
    		set session("RS_PWAM_ACCESS")=Conn.execute(sSQL)
	
			Set myRS= Server.CreateObject("ADODB.Recordset")
			set myrs =  session("RS_PWAM_ACCESS")
			SESSION("DEPT_ID") = 999
			'response.write "<P> record count = "& myrs.recordcount  & "<P>"
			if not myrs.eof then
								do while not myrs.eof 
												if cint(myrs("pwam_dept_id")) = 0  and  cint(myrs("access_level")) < 7 then
														SESSION("DEPT_ID") = 0
												else
															SESSION("DEPT_ID") = myrs("pwam_dept_id") 
															 SESSION("access_level") = myrs("access_level")
												end if
												'response.write " <P>SESSION RS VALUE  = "& myrs("pwam_dept_id") & "      " & myrs("access_level")
												myrs.movenext
								loop
			end if
			'------------------------
			
			'response.write " ssql = " & ssql 
			
						'response.write "count = " & myrs.recordcount	
						redim access(100)
						'for x = 1 to 100
							'session("access("&x&")") = 10
						'next
						
						if not myrs.bof and not myrs.eof then
						myrs.movefirst
								do while not myrs.eof
												access(myrs("pwam_dept_id")) = myrs("access_level")
												
											'response.write " access"&myrs("pwam_dept_id")&" = "& access(myrs("pwam_dept_id")) & "<P>"
												myrs.movenext
								loop
						end if 
		'SESSION("DEPT_ID") = 4
		'session("pwam_ad_group_list")  =  session("pwam_ad_group_list") & "','Redstone-PWAM-Eng-Admin"
 		
		
		'GET LIST OF REQUESTORS
															Set myRSREQ = Server.CreateObject("ADODB.Recordset")
															sSQL = "SELECT  count(*) as count from pwam_CUST_INDX WHERE (NETWORK_USER_NAME =  '"& trim(session("PWAM_UserID")) &"') " 
															'response.write ssql
															myRSREQ.Open Ssql, Conn
															if not myRSREQ.eof and not myRSREQ.bof then
															'response.write "count = " & myRSREQ("count")
																if cint(myRSREQ("count")) > 0 then
																	session("bREQUESTOR") = "Y"
																	'response.write "req = " & session("bREQUESTOR") 
																end if
															end if
															
															
if  trim(session("PWAM_UserID")) = "PUBLIC"  or  trim(session("PWAM_UserID")) = "Public"  or trim(session("PWAM_UserID")) = "public" then 
	session("DEPT_ID") = 999
	session("bREQUESTOR") = "N"
end if

if instr( session("pwam_ad_group_list") ,"Redstone-PWAM-ENV-Journal") then
	session("bDES_JOURNAL") = "Y"
else
	session("bDES_JOURNAL") = "N"
end if
 'response.write "xxxxxxxxxxxxxxx" &  session("bDES_JOURNAL") 
 %>
<html>
<head>
<title>Directorate of Public Works</title>
	<link rel="STYLESHEET" href="garrisondpw.css" type="text/css">
<!-- 	<style>
	body{
	background: url(dpw_bkgd_600.gif)}
	</style> -->
	<SCRIPT LANGUAGE="JavaScript">
function popUp(URL) {
day = new Date();
id = day.getTime();
eval("page" + id + " = window.open(URL, '" + id + "', 'toolbar=1,scrollbars=1,location=1,status=1,menubar=1,resizable=1,width=800,height=500,left = 112,top = 84');");}
</script>

</head>

<body onload="javascript:popUp('index_frames.asp')" style="background: url(dpw_bkgd_600.gif) ! important background-repeat: none  ! important">
<FONT FACE=Arial SIZE=3><B><%=Session("PWAM_UserID")%></B></FONT><FONT FACE=Arial SIZE=3><B><%'session("pwam_ad_group_list") %></B></FONT>
<table width="600" cellspacing="1" cellpadding="1" border="0">
<tr><td colspan="2" align="center"><i>U.S. Army Garrison - Redstone, 4488 Martin Road, AMSAM-RA-DPW, Redstone Arsenal, AL 35898<br>256-876-1023&nbsp;&nbsp;&nbsp; DSN 746-1023</i></td></tr><tr><td>&nbsp;</td></tr>
<tr valign="top">
 <td width="200"><img src="b_red.gif" width="17" height="17" border="0" align="absmiddle" alt="red arrow bullet">&nbsp;<font size="+1">Organization<br>
<img src="b_red.gif" width="17" height="17" border="0" align="absmiddle" alt="red arrow bullet">&nbsp;Personnel Directory<br>
<img src="b_red.gif" width="17" height="17" border="0" align="absmiddle" alt="red arrow bullet">&nbsp;<a href="policies.asp" target="_self">Regulations</a><br>
<img src="b_red.gif" width="17" height="17" border="0" align="absmiddle" alt="red arrow bullet">&nbsp;Forms<br>
<img src="b_red.gif" width="17" height="17" border="0" align="absmiddle" alt="red arrow bullet">&nbsp;<a href="maps/index.html">Maps</a><br>
<img src="b_red.gif" width="17" height="17" border="0" align="absmiddle" alt="red arrow bullet">&nbsp;<a href="mission.html">Mission Statement</a></font></td>
    <td>
	<table width="400" cellspacing="1" cellpadding="1" border="0">
		<tr valign="top">
			    <td><u>PROJECT INFORMATION</u>
				<br>&nbsp;-M&R Projects
				<br>&nbsp;-MCA Projects
				<br>&nbsp;-Customer Projects
				<br><br><u>REPORTS</u>
				<br>&nbsp;-JOR
				<br>&nbsp;-SO
				<br>&nbsp;-Projects
				<br>&nbsp;-Other</td>
			   <%If cint(session("DEPT_ID")) = 0 or cint(session("DEPT_ID")) = 99 then %>	
    	<td><u><a href="PWAM/PWAM_JOR_MAN_TOP.asp"><u><FONT FACE=Arial SIZE=2><B>JOB ORDER REQUEST</B></FONT></u></a></u>
	<% Else  %>
		<td><u><FONT FACE=Arial SIZE=2><B>JOB ORDER REQUEST</B></FONT></u>
	<% End If %>
	<br>&nbsp;<a href="PWAM/PWAM_SUB_ENTRY_top.asp"><FONT FACE=Arial SIZE=2>- JOR Submission</FONT></a>
	<% If cstr(session("bREQUESTOR")) = "Y"  then  %>
	<br>&nbsp;<a href="PWAM/PWAM_SUB_UPDATE_TOP.asp"><FONT FACE=Arial SIZE=2>- Requestor JOR Review</FONT></a>
	<% else
			If cint(session("DEPT_ID")) => 0  and  cint(session("DEPT_ID")) <= 99 then  %>
				<br>&nbsp;<a href="PWAM/PWAM_SUB_UPDATE_TOP.asp"><FONT FACE=Arial SIZE=2>- Requestor JOR Review</FONT></a>
	<% End If 
	end if
	'response.write "dept id = " &  session("DEPT_ID") & "<p>"
	If cint(session("DEPT_ID")) => 0 and cint(session("DEPT_ID")) <= 99 then %>
			<br>&nbsp;<a href="PWAM/PWAM_JOR_REV_TOP.asp"><FONT FACE=Arial SIZE=2>- DPW JOR  Review</FONT></a>
			<br>&nbsp;<a href="PWAM/PWAM_DEPT_RPT_TOP.asp"><FONT FACE=Arial SIZE=2>- JOR  Status</FONT></a>
	<% End If %>
	<%'response.write "dept id = " &  session("DEPT_ID") & "<p>"
	If cint(session("DEPT_ID")) = 0 or cint(session("DEPT_ID")) = 99 then %>
			<br>&nbsp;<a href="PWAM/PWAM_DEPT_ADMIN_TOP.asp"><FONT FACE=Arial SIZE=2>- JOR Activity</FONT></a>
	<% End If 
	If cint(session("DEPT_ID")) > 0  and  cint(session("DEPT_ID")) < 99 then %>
			
					<br>&nbsp;<a href="PWAM/PWAM_DEPT_TOP.asp"><FONT FACE=Arial SIZE=2>- JOR Activity</FONT></a>
			
	<% End If 
	If cint(session("DEPT_ID")) = 0  and  cint(session("DEPT_ID")) = 99 then %>
			<br>&nbsp;<a href="PWAM/PWAM_JORNAL_TOP.asp"><FONT FACE=Arial SIZE=2>- JOR Journal</FONT></a>
	<% End If
	If cint(session("DEPT_ID")) = 0  or  cint(session("DEPT_ID")) = 99 then %>
			<br>&nbsp;<a href="PWAM/pwam_ADMIN_MAIN.asp"><FONT FACE=Arial SIZE=2>- ADMIN Utilities</FONT></a>
	<% End If
	If cint(session("DEPT_ID")) = 0  or  cint(session("DEPT_ID")) = 99 then %>
			<br>&nbsp;<a href="PWAM/pwam_FIND_top.asp"><FONT FACE=Arial SIZE=2>- FIND Utilities</FONT></a>
	<% End If
	If cint(session("DEPT_ID")) => 0  and  cint(session("DEPT_ID")) <= 99 then %>
			<br>&nbsp;<a href="PWAM/pwam_docreq_top.asp"><FONT FACE=Arial SIZE=2>- Dept. Requirement List</FONT></a>
	<% End If
	if cint(session("DEPT_ID")) = 99 then %>
			<br>&nbsp;<a href="PWAM/pwam_permissions_check.asp"><FONT FACE=Arial SIZE=2>- PERMISSIONS</FONT></a>
	<% End If %>
				<br><br><u>SERVICE ORDER REQUEST</u>
				<br>&nbsp;-SO Submission
				<br>&nbsp;-SO Status
				<br><br><u>SPACE REQUEST</u>
				<br>&nbsp;-SR Submission
				<br>&nbsp;-SR Status
				<br /><FONT FACE=Arial SIZE=2><a target="_blank" href="Space Request Package.pdf">&nbsp;-Space Request Package (Save Separately and attach to Job Order Request)</a></FONT>
				</td>
		</tr>
	</table>
</td>
</tr>
</table>

<p></p>
<!-- INSERT PAGE DATA ABOVE THIS LINE -->

<br><div align="center"><font size="-2"><i>Maintained by U.S. Army Garrison - Redstone<br>
<script language="JavaScript" type="text/javascript">
<!--

	var a;
	a=new Date(document.lastModified);
	lm_year=a.getFullYear();
	if (lm_year<1000){ 				//just in case date is delivered with 4 digits
		if (lm_year<70){
		lm_year=2000+lm_year;
		}
		else lm_year=1900+lm_year;
	}								//end workaround
	lm_month=a.getMonth()+1;
	if (lm_month<10){
		lm_month='0'+lm_month;
	}
	lm_day=a.getDate();
	if (lm_day<10){
		lm_day='0'+lm_day;
	}
	document.write("Last Modified " + lm_year+'-'+lm_month+'-'+lm_day);
// -->
  </script></i></font></div>


</body>
</html>