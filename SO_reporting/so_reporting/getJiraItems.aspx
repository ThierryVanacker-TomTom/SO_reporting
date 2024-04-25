<%@ Page Language="VB" Explicit="True" Debug="true"%>

<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Collections" %>
<%@ Import Namespace="System.Web.Services" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.Linq" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="DocumentFormat.OpenXml" %>
<%@ Import Namespace="DocumentFormat.OpenXml.Packaging" %>
<%@ Import Namespace="DocumentFormat.OpenXml.Spreadsheet" %>
<%@ Import Namespace="DocumentFormat.OpenXml.Wordprocessing" %>
<%@ Import Namespace="System.Web.Script.Serialization" %>

<script language=vb runat=server>
Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12

Dim webClient As New System.Net.WebClient
'webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded")
webClient.Headers.Add("Content-Type", "application/json; charset=utf-8")
'webClient.Headers.Add("Authorization", "Basic " & "c3ZjX3N0c19qaXJhX3VzZXI6RjIzcnQjNDV0eSM0")
webClient.Headers.Add("Authorization", "Bearer " & "NjEwNTg3NTc5NjU0Ot6RDMksA3OJYGutuCZaMVZs6e1i")
webClient.Headers.Add("X-Atlassian-Token", "nocheck") 

Dim result As String 
'result = webClient.DownloadString("https://soreporting.azurewebsites.net/getsuppipeitem.asp?id=1")
Dim url As String
dim prj as string

dim jql as String
jql = request.querystring("jql")
'if jql = "" then jql = " issuekey=""om-11925"" "
if jql = "" then jql = " issuekey=""OM-29043"" "

'	jql = ""
'	jql = jql & " project = ""STS"" "
'	'jql = jql & " AND issuetype in (""Activity - Archive Validation"", ""Activity - Feasibility"", ""Activity - Manual Production"", ""Activity - Other"", ""Activity - Quality Analysis"", ""Activity - SQR Measurement"", ""Activity - Source Acquisition"", ""Activity - Source Acquisition - Field"", ""Activity - Source Analysis"", ""Activity - Source Preparation"")"
'	'jql = jql & " AND ""Assigned Unit"" in (""SO EECA"", ""SO AFR"", ""SO WCE"", ""SO STS"", ""SO SAM"", ""SO PDV"", ""SO OCE"", ""SO NEA"", ""SO NAM"", ""SO LAM"")"
'	jql = jql & " AND ("
'	jql = jql & " status changed to done during (""" & from_ & """, """ & to_ & """)"
'	'jql = jql & " OR status changed to closed during (""" & from_ & """"", """ & to_ & """)"
'	'jql = jql & " OR status changed to planned during (""" & from_ & """"", """ & to_ & """)"
'	jql = jql & " )"

'url = "https://jira.tomtomgroup.com/rest/api/2/search?jql=project=""" & prj & """ and component=""RSO Global Alignment""&fields=summary,issuetype,status,created,updated,duedate,resolutiondate,customfield_11860,fixVersions,components,customfield_19662&maxResults=-1"
'url = "https://jira.tomtomgroup.com/rest/api/2/search?jql=project=""" & prj & """&fields=summary,issuetype,status,created,updated,duedate,resolutiondate,customfield_11860,fixVersions,components,customfield_19662&maxResults=-1"
'url = "https://jira.tomtomgroup.com/rest/api/2/search?jql=" & jql & "&fields=summary,issuetype,status,created,updated,duedate,resolution,resolutiondate,customfield_20162,customfield_20164,customfield_20268,customfield_20363,customfield_11860,fixVersions,components,customfield_19662&maxResults=-1"
url = "https://jira.tomtomgroup.com/rest/api/2/search?jql=" & Server.UrlEncode(jql) & "&fields=summary,issuetype,status,created,updated,duedate,resolution,resolutiondate,customfield_20162,customfield_20164,customfield_20268,customfield_20363,customfield_11860,customfield_23662,assignee,subtasks,parent,customfield_20260,customfield_20261&maxResults=-1"
'url = "https://jira.tomtomgroup.com/rest/api/2/search?jql=" & jql & "&fields=summary,issuetype,status,created,updated&maxResults=-1"
'response.write (url)

result = webClient.DownloadString(url)
dim tmp as string
dim arr
dim a
'response.Write (now)

response.write (result)
responsE.end

Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
'Dim issuesJson As Dictionary(Of String, Object) 

dim jira_key as string
dim xml as string
xml = ""
'xml = xml & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf
xml = xml & "<item_result>" & vbcrlf
For Each pair In parentJson
	if (pair.Key.tostring = "issues" ) then
		jira_key = ""
		'here we have a list of issues - issues is an arraylist
		For each issue in pair.value 
		xml = xml & "<item>"
		for each field_in_issue in issue
			if (field_in_issue.key.tostring = "key") then
				jira_key = field_in_issue.value.tostring
				xml = xml & "<key>" & jira_key & "</key>"
			end if

			if (field_in_issue.Key.tostring = "fields" ) then
				for each field in field_in_issue.value 'loop though the fields within ONE issue
					
					if (field.key.tostring = "summary") then
					if (field.value is nothing) then
						xml = xml & "<summary>" & "" & "</summary>" & vbcrlf
					else
						xml = xml & "<summary>" & xmlsafe("" & field.value.tostring) & "</summary>" & vbcrlf
					end if
					end if

					if (field.key.tostring = "issuetype") then
					if (field.value is nothing) then
						xml = xml & "<issuetype>" & "" & "</issuetype>" & vbcrlf
					else
						'xml = xml & "<issuetype>" & xmlsafe("" & field.value.tostring) & "</issuetype>" & vbcrlf
						for each field_in_issuetype in field.value
						if (field_in_issuetype.key.tostring = "name") then
							xml = xml & "<issuetype>" & xmlsafe("" & field_in_issuetype.value.tostring) & "</issuetype>" & vbcrlf
						end if
						next
					end if
					end if

					if (field.key.tostring = "parent") then
					if (field.value is nothing) then
						xml = xml & "<parent>" & "" & "</parent>" & vbcrlf
					else
						'xml = xml & "<parent>" & xmlsafe("" & field.value.tostring) & "</parent>" & vbcrlf
						for each field_in_issuetype in field.value
						if (field_in_issuetype.key.tostring = "key") then
							xml = xml & "<parent>" & xmlsafe("" & field_in_issuetype.value.tostring) & "</parent>" & vbcrlf
						end if
						next
					end if
					end if

					if (field.key.tostring = "duedate") then
					if (field.value is nothing) then
						xml = xml & "<duedate>" & "" & "</duedate>" & vbcrlf
					else
						xml = xml & "<duedate>" & xmlsafe(toYYYYMMDD("" & field.value.tostring)) & "</duedate>" & vbcrlf
					end if
					end if

					if (field.key.tostring = "customfield_19662") then
					if (field.value is nothing) then
						xml = xml & "<target_duedate>" & "" & "</target_duedate>" & vbcrlf
					else
						xml = xml & "<target_duedate>" & xmlsafe(toYYYYMMDD("" & field.value.tostring)) & "</target_duedate>" & vbcrlf
					end if
					end if

					if (field.key.tostring = "customfield_20164") then
					if (field.value is nothing) then
						xml = xml & "<bl_enddate>" & "" & "</bl_enddate>" & vbcrlf
					else
						xml = xml & "<bl_enddate>" & xmlsafe(toYYYYMMDD("" & field.value.tostring)) & "</bl_enddate>" & vbcrlf
					end if
					end if

					if (field.key.tostring = "customfield_20162") then
					if (field.value is nothing) then
						xml = xml & "<enddate>" & "" & "</enddate>" & vbcrlf
					else
						xml = xml & "<enddate>" & xmlsafe(toYYYYMMDD("" & field.value.tostring)) & "</enddate>" & vbcrlf
					end if
					end if

					if (field.key.tostring = "status") then
					if (field.value is nothing) then
						xml = xml & "<status>" & "" & "</status>" & vbcrlf
					else
						'xml = xml & "<issuetype>" & xmlsafe("" & field.value.tostring) & "</issuetype>" & vbcrlf
						for each field_in_status in field.value
						if (field_in_status.key.tostring = "name") then
							xml = xml & "<status>" & xmlsafe("" & field_in_status.value.tostring) & "</status>" & vbcrlf
						end if
						next
					end if
					end if
					if (field.key.tostring = "customfield_20260") then
					if (field.value is nothing) then
						xml = xml & "<riskconsequence>" & "" & "</riskconsequence>" & vbcrlf
					else
						'xml = xml & "<riskconsequence>" & xmlsafe("" & field.value.tostring) & "</riskconsequence>" & vbcrlf
						for each field_in_status in field.value
						if (field_in_status.key.tostring = "value") then
							xml = xml & "<riskconsequence>" & xmlsafe("" & field_in_status.value.tostring) & "</riskconsequence>" & vbcrlf
						end if
						next
					end if
					end if
					if (field.key.tostring = "customfield_20261") then
					if (field.value is nothing) then
						xml = xml & "<riskprobability>" & "" & "</riskprobability>" & vbcrlf
					else
						'xml = xml & "<riskprobability>" & xmlsafe("" & field.value.tostring) & "</riskprobability>" & vbcrlf
						for each field_in_status in field.value
						if (field_in_status.key.tostring = "value") then
							xml = xml & "<riskprobability>" & xmlsafe("" & field_in_status.value.tostring) & "</riskprobability>" & vbcrlf
						end if
						next
					end if
					end if

					if (field.key.tostring = "assignee") then
					if (field.value is nothing) then
						xml = xml & "<assignee>" & "" & "</assignee>" & vbcrlf
					else
						'xml = xml & "<issuetype>" & xmlsafe("" & field.value.tostring) & "</issuetype>" & vbcrlf
						for each field_in_status in field.value
						if (field_in_status.key.tostring = "name") then
							xml = xml & "<assignee>" & xmlsafe("" & field_in_status.value.tostring) & "</assignee>" & vbcrlf
						end if
						next
					end if
					end if

					if (field.key.tostring = "customfield_23662") then
					if (field.value is nothing) then
						xml = xml & "<decision>" & "" & "</decision>" & vbcrlf
					else
						'xml = xml & "<issuetype>" & xmlsafe("" & field.value.tostring) & "</issuetype>" & vbcrlf
						for each field_in_status in field.value
						if (field_in_status.key.tostring = "value") then
							xml = xml & "<decision>" & xmlsafe("" & field_in_status.value.tostring) & "</decision>" & vbcrlf
						end if
						next
					end if
					end if

					if (field.key.tostring = "resolution") then
					if (field.value is nothing) then
						xml = xml & "<resolution>" & "" & "</resolution>" & vbcrlf
					else
						'xml = xml & "<issuetype>" & xmlsafe("" & field.value.tostring) & "</issuetype>" & vbcrlf
						for each field_in_status in field.value
						if (field_in_status.key.tostring = "name") then
							xml = xml & "<resolution>" & xmlsafe("" & field_in_status.value.tostring) & "</resolution>" & vbcrlf
						end if
						next
					end if
					end if

					if (field.key.tostring = "customfield_20363") then
					if (field.value is nothing) then
						xml = xml & "<assigned_unit>" & "" & "</assigned_unit>" & vbcrlf
					else
						'xml = xml & "<issuetype>" & xmlsafe("" & field.value.tostring) & "</issuetype>" & vbcrlf
						for each field_in_status in field.value
						if (field_in_status.key.tostring = "value") then
							xml = xml & "<assigned_unit>" & xmlsafe("" & field_in_status.value.tostring) & "</assigned_unit>" & vbcrlf
						end if
						next
					end if
					end if

					if (field.key.tostring = "customfield_20268") then
					if (field.value is nothing) then
						xml = xml & "<requesting_unit>" & "" & "</requesting_unit>" & vbcrlf
					else
						'xml = xml & "<issuetype>" & xmlsafe("" & field.value.tostring) & "</issuetype>" & vbcrlf
						for each field_in_status in field.value
						if (field_in_status.key.tostring = "value") then
							xml = xml & "<requesting_unit>" & xmlsafe("" & field_in_status.value.tostring) & "</requesting_unit>" & vbcrlf
						end if
						next
					end if
					end if

					if (field.key.tostring = "created") then
					if (field.value is nothing) then
						xml = xml & "<created>" & "" & "</created>" & vbcrlf
					else
						xml = xml & "<created>" & xmlsafe(toYYYYMMDD("" & field.value.tostring)) & "</created>" & vbcrlf
					end if
					end if

					if (field.key.tostring = "updated") then
					if (field.value is nothing) then
						xml = xml & "<updated>" & "" & "</updated>" & vbcrlf
					else
						xml = xml & "<updated>" & xmlsafe(toYYYYMMDD("" & field.value.tostring)) & "</updated>" & vbcrlf
					end if
					end if

					if (field.key.tostring = "resolutiondate") then
					if (field.value is nothing) then
						xml = xml & "<resolutiondate>" & "" & "</resolutiondate>" & vbcrlf
					else
						xml = xml & "<resolutiondate>" & xmlsafe(toYYYYMMDD("" & field.value.tostring)) & "</resolutiondate>" & vbcrlf
					end if
					end if

					if (field.key.tostring = "subtasks") then
					if (field.value is nothing) then
						xml = xml & "<subtasks>" & "" & "</subtasks>" & vbcrlf
					else
						tmp = "||"
						for each comp_item in field.value
						for each field_in_comp_item in comp_item
						if (field_in_comp_item.key.tostring = "key") then
							tmp = tmp & trim(field_in_comp_item.value.tostring) & "||"
						end if
						next
						next
						if tmp = "||" then tmp = ""
						xml = xml & "<subtasks>" & xmlsafe(tmp) & "</subtasks>" & vbcrlf
					end if
					end if

					if (field.key.tostring = "components") then
					if (field.value is nothing) then
						xml = xml & "<components>" & "" & "</components>" & vbcrlf
					else
						tmp = "||"
						for each comp_item in field.value
						for each field_in_comp_item in comp_item
						if (field_in_comp_item.key.tostring = "name") then
							tmp = tmp & trim(field_in_comp_item.value.tostring) & "||"
						end if
						next
						next
						if tmp = "||" then tmp = ""
						xml = xml & "<components>" & xmlsafe(tmp) & "</components>" & vbcrlf
					end if
					end if
					
					'sprintinfo - customfield_11860
					'sprintenddate - take latest completeddate - customfield_11860
					if (field.key.tostring = "customfield_11860") then
					if (field.value is nothing) then
						xml = xml & "<sprintinfo>" & "" & "</sprintinfo>" & vbcrlf
						xml = xml & "<sprintcompletedate>" & "" & "</sprintcompletedate>" & vbcrlf
					else
						tmp = "||"
						for each sprintinfo_item in field.value
						'for each field_in_sprintinfo_item in sprintinfo_item
						'if (field_in_comp_item.key.tostring = "name") then
							tmp = tmp & trim(sprintinfo_item) & "||"
						'end if
						'next
						next
						if tmp = "||" then tmp = ""
						xml = xml & "<sprintinfo>" & xmlsafe(tmp) & "</sprintinfo>" & vbcrlf
						xml = xml & "<sprintcompletedate>" & xmlsafe(toYYYYMMDD("" & getSprintLatestCompletedDate(tmp))) & "</sprintcompletedate>" & vbcrlf
					end if
					end if
				next 'loop through all the fields
				
				'history
				'first make the call
				dim hist_
				hist_ = getHist_(jira_key)
				
				'second - compose the transition history
				tmp = parseTransition_(hist_)
				
				if tmp = "" then
					xml = xml & "<transitions_history>" & "" & "</transitions_history>" & vbcrlf
				else
					xml = xml & "<transitions_history>" & tmp & "</transitions_history>" & vbcrlf
				end if
				
				'third - compose the assignee history
				tmp = parseAssignee_(hist_)
				if tmp = "" then
					xml = xml & "<assignee_history>" & "" & "</assignee_history>" & vbcrlf
				else
					xml = xml & "<assignee_history>" & tmp & "</assignee_history>" & vbcrlf
				end if

				tmp = parseRiskConsequence_(hist_)
				if tmp = "" then
					xml = xml & "<riskconsequence_history>" & "" & "</riskconsequence_history>" & vbcrlf
				else
					xml = xml & "<riskconsequence_history>" & tmp & "</riskconsequence_history>" & vbcrlf
				end if
				tmp = parseRiskProbability_(hist_)
				if tmp = "" then
					xml = xml & "<riskprobability_history>" & "" & "</riskprobability_history>" & vbcrlf
				else
					xml = xml & "<riskprobability_history>" & tmp & "</riskprobability_history>" & vbcrlf
				end if
				tmp = parseDecision_(hist_)
				if tmp = "" then
					xml = xml & "<decision_history>" & "" & "</decision_history>" & vbcrlf
				else
					xml = xml & "<decision_history>" & tmp & "</decision_history>" & vbcrlf
				end if

				'11JAN2019 - compose Baseline enddate history				
				tmp = parseBaselineEndDate_(hist_)
				if tmp = "" then
					xml = xml & "<bl_enddate_history>" & "" & "</bl_enddate_history>" & vbcrlf
				else
					xml = xml & "<bl_enddate_history>" & tmp & "</bl_enddate_history>" & vbcrlf
				end if
				
'					statedate = "00000000000000"
'					'look for the latest timestamp the item got the current state
'					arr = split(tmp, "@@")
'					for a = lbound(arr) to ubound(arr)
'					if arr(a) <> "" then
'						arr2 = split(arr(a), "||")
'						'Arr(a) = date||from||tostring
'						if arr2(2) = state_ then
'						if arr2(0) > statedate then
'							statedate = arr2(0)
'						end if
'						end if
'					end if
'					next
'					if statedate = "00000000000000" then statedate = ""
'					xml = xml & "<transitions_history>" & xmlsafe("" & statedate) & "</transitions_history>" & vbcrlf

				'8JUN2020 - add comment (do this sep because otherwise we get an error the return string became too big)
				tmp = getComments_(jira_key)
				if tmp = "" then
					xml = xml & "<comments_history>" & "" & "</comments_history>" & vbcrlf
				else
					xml = xml & "<comments_history>" & tmp & "</comments_history>" & vbcrlf
				end if
				
			end if
		next
		xml = xml & "</item>"
		next
	end if	
Next

xml = xml & "</item_result>" & vbcrlf

'Server.ScriptTimeout = 60*60
Response.ContentType = "text/xml"
Response.CharSet = "UTF-8"
response.write (xml)
'response.Write (now)
end sub

function getSprintLatestCompletedDate(s)
'multiple sprint may be attached, take the latest one
'IN : ["com.atlassian.greenhopper.service.sprint.Sprint@790ac54f[id=11467,rapidViewId=3248,state=CLOSED,name=SDP Sprint 1,startDate=2017-11-13T06:12:41.559Z,endDate=2017-11-27T06:12:00.000Z,completeDate=2017-11-29T09:20:45.503Z,sequence=11467]","com.atlassian.greenhopper.service.sprint.Sprint@728aac4a[id=12185,rapidViewId=3248,state=CLOSED,name=SDP Sprint 2,startDate=2018-02-12T12:02:27.361Z,endDate=2018-02-26T12:02:00.000Z,completeDate=2018-02-23T11:26:35.645Z,sequence=12185]"]
Dim arr
Dim arr2
Dim arr3
Dim a
Dim a2
Dim tmp
getSprintLatestCompletedDate = "00000000"
arr = Split(s, "||")
For a = LBound(arr) To UBound(arr)
If arr(a) <> "" Then
    arr2 = Split(arr(a), ",")
    For a2 = LBound(arr2) To UBound(arr2)
    If arr2(a2) <> "" Then
        arr3 = Split(arr2(a2), "=")
        If arr3(0) = "completeDate" Then
            If "" & arr3(1) <> "" And "" & arr3(1) <> "<null>" Then
            tmp = Mid(arr3(1), 1, 4) & Mid(arr3(1), 6, 2) & Mid(arr3(1), 9, 2)
            'response.write s & "##" & tmp & "<br>"
    
            'since we might have multiple sprints check if this one is the latest
            If tmp > getSprintLatestCompletedDate Then getSprintLatestCompletedDate = tmp
        End If
        End If
    End If
    Next
End If
Next
'if nothing found then return empty
If getSprintLatestCompletedDate = "00000000" Then getSprintLatestCompletedDate = ""
end function

function xmlsafe(s)
xmlsafe = s
xmlsafe = replace(xmlsafe, "&", "&amp;")
xmlsafe = replace(xmlsafe, "<", "&lt;")
xmlsafe = replace(xmlsafe, ">", "&gt;")
xmlsafe = replace(xmlsafe, chr(11), "")
end function

function getHist_(jira_key)
dim url as string
Dim webClient As New System.Net.WebClient
Dim result As String 

url = "https://jira.tomtomgroup.com/rest/api/2/issue/" & jira_key & "?expand=changelog"
webClient.Headers.Add("Content-Type", "application/json; charset=utf-8")
'webClient.Headers.Add("Authorization", "Basic " & "c3ZjX3N0c19qaXJhX3VzZXI6RjIzcnQjNDV0eSM0")
webClient.Headers.Add("Authorization", "Bearer " & "NjEwNTg3NTc5NjU0Ot6RDMksA3OJYGutuCZaMVZs6e1i")
webClient.Headers.Add("X-Atlassian-Token", "nocheck") 

'try
	result = webClient.DownloadString(url)
	getHist_ = result
'catch e as exception
'response.write ("HelpLink : " &	e.HelpLink 		& "<br>")	'Link to the help file associated with this exception.
''response.write ("InnerException : " & e.InnerException 	& "<br>")	'A reference to the inner exceptionthe exception that originally occurred, if this exception is based on a previous exception. Exceptions can be nested. That is, when a procedure throws an exception, it can nest another exception inside the exception it's raising, passing both exceptions out to the caller. The InnerException property gives access to the inner exception.
'response.write ("Message : " &	e.Message 		& "<br>")	'Error message text.
'response.write ("StackTrace : " &	e.StackTrace 	& "<br>")	'The stack trace, as a single string, at the point the error occurred.
''response.write ("TargetSite : " &	e.TargetSite 	& "<br>")	'The name of the method that raised the exception.
'response.write ("ToString : " &	e.ToString 		& "<br>")	'Converts the exception name, description, and the current stack dump into a single string.
'response.write ("Message : " &	e.Message 		& "<br>")	'Returns a description of the error that occurred.
'	'give it another try
''	result = webClient.DownloadString(url)
''	getHist_ = result
'end try
end function

function getComments_(jira_key)

'return is datetime||comment@@...
getComments_ = ""
dim url as string
Dim webClient As New System.Net.WebClient
Dim result As String 

url = "https://jira.tomtomgroup.com/rest/api/2/issue/" & jira_key & "/comment"
webClient.Headers.Add("Content-Type", "application/json; charset=utf-8")
'webClient.Headers.Add("Authorization", "Basic " & "c3ZjX3N0c19qaXJhX3VzZXI6RjIzcnQjNDV0eSM0")
webClient.Headers.Add("Authorization", "Bearer " & "NjEwNTg3NTc5NjU0Ot6RDMksA3OJYGutuCZaMVZs6e1i")
webClient.Headers.Add("X-Atlassian-Token", "nocheck") 

'try
result = webClient.DownloadString(url)

Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string
dim body as string

For Each pair In parentJson
	'if (pair.Key.tostring = "comments" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
	'	for each field in pair.value 'dictionary loop
			if (pair.key.tostring = "comments") then 'comments is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in pair.value 'array loop
				created = ""
				body = ""
				
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.Key.tostring = "body") then
						body = xmlsafe("" & p.Value.tostring)
					end if
					
				next 'dictionary loop
				
				getComments_ = getComments_ & created & "||"
				getComments_ = getComments_ & body & "@@"
				
				next 'array loop
			end if
	'	next
	'end if
next
'catch e as exception
'response.write ("HelpLink : " &	e.HelpLink 		& "<br>")	'Link to the help file associated with this exception.
''response.write ("InnerException : " & e.InnerException 	& "<br>")	'A reference to the inner exceptionthe exception that originally occurred, if this exception is based on a previous exception. Exceptions can be nested. That is, when a procedure throws an exception, it can nest another exception inside the exception it's raising, passing both exceptions out to the caller. The InnerException property gives access to the inner exception.
'response.write ("Message : " &	e.Message 		& "<br>")	'Error message text.
'response.write ("StackTrace : " &	e.StackTrace 	& "<br>")	'The stack trace, as a single string, at the point the error occurred.
''response.write ("TargetSite : " &	e.TargetSite 	& "<br>")	'The name of the method that raised the exception.
'response.write ("ToString : " &	e.ToString 		& "<br>")	'Converts the exception name, description, and the current stack dump into a single string.
'response.write ("Message : " &	e.Message 		& "<br>")	'Returns a description of the error that occurred.
'	'give it another try
''	result = webClient.DownloadString(url)
''	getHist_ = result
'end try
end function

function parseTransition_(result)
parseTransition_ = ""
if (result is nothing) then exit function

Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer() 
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string

For Each pair In parentJson
	if (pair.Key.tostring = "changelog" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
		for each field in pair.value 'dictionary loop
			if (field.key.tostring = "histories") then 'histories is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in field.value 'array loop
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.key.tostring = "items") then 'items is an array, loop through the items, within the items loop through as a dictionary
						for each itemitem in p.value 'array loop
						bln = false 'use a toggle to start capturing status info
						for each p_item in itemitem 'dictionary loop							
							if (p_item.Value is nothing) then
							else
								if (p_item.key.tostring = "field" and p_item.value.tostring = "status") then
									bln = true
									parseTransition_ = parseTransition_ & created & "||"
								end if
								if (p_item.key.tostring = "fromString" and bln) then
									parseTransition_ = parseTransition_ & p_item.value.tostring & "||"	
								end if
								if (p_item.key.tostring = "toString" and bln) then
									parseTransition_ = parseTransition_ & p_item.value.tostring & "@@"	
								end if
							end if
					'		
						next
						next
					end if
				next 'dictionary loop
				next 'array loop
			end if
		next
	end if
next
end function

function parseAssignee_(result)
parseAssignee_ = ""
if (result is nothing) then exit function

Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string

For Each pair In parentJson
	if (pair.Key.tostring = "changelog" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
		for each field in pair.value 'dictionary loop
			if (field.key.tostring = "histories") then 'histories is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in field.value 'array loop
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.key.tostring = "items") then 'items is an array, loop through the items, within the items loop through as a dictionary
						for each itemitem in p.value 'array loop
						bln = false 'use a toggle to start capturing status info
						for each p_item in itemitem 'dictionary loop							
							'if (p_item.Value is nothing) then
							'else
								if (p_item.Value is nothing) then
								else
								if (p_item.key.tostring = "field" and p_item.value.tostring = "assignee") then
									bln = true
									parseAssignee_ = parseAssignee_ & created & "||"
								end if
								end if
								if (p_item.key.tostring = "fromString" and bln) then
								if (p_item.Value is nothing) then
									parseAssignee_ = parseAssignee_ & "" & "||"	
								else
									parseAssignee_ = parseAssignee_ & p_item.value.tostring & "||"	
								end if
								end if
								if (p_item.key.tostring = "toString" and bln) then
								if (p_item.Value is nothing) then
									parseAssignee_ = parseAssignee_ & "" & "@@"
								else
									parseAssignee_ = parseAssignee_ & p_item.value.tostring & "@@"
								end if
								end if
							'end if
					'		
						next
						next
					end if
				next 'dictionary loop
				next 'array loop
			end if
		next
	end if
next
end function

function parseRiskConsequence_(result)
parseRiskConsequence_ = ""
if (result is nothing) then exit function

Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string

For Each pair In parentJson
	if (pair.Key.tostring = "changelog" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
		for each field in pair.value 'dictionary loop
			if (field.key.tostring = "histories") then 'histories is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in field.value 'array loop
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.key.tostring = "items") then 'items is an array, loop through the items, within the items loop through as a dictionary
						for each itemitem in p.value 'array loop
						bln = false 'use a toggle to start capturing status info
						for each p_item in itemitem 'dictionary loop							
							'if (p_item.Value is nothing) then
							'else
								if (p_item.Value is nothing) then
								else
								if (p_item.key.tostring = "field" and p_item.value.tostring = "Risk consequence") then
									bln = true
									parseRiskConsequence_ = parseRiskConsequence_ & created & "||"
								end if
								end if
								if (p_item.key.tostring = "fromString" and bln) then
								if (p_item.Value is nothing) then
									parseRiskConsequence_ = parseRiskConsequence_ & "" & "||"	
								else
									parseRiskConsequence_ = parseRiskConsequence_ & p_item.value.tostring & "||"	
								end if
								end if
								if (p_item.key.tostring = "toString" and bln) then
								if (p_item.Value is nothing) then
									parseRiskConsequence_ = parseRiskConsequence_ & "" & "@@"
								else
									parseRiskConsequence_ = parseRiskConsequence_ & p_item.value.tostring & "@@"
								end if
								end if
							'end if
					'		
						next
						next
					end if
				next 'dictionary loop
				next 'array loop
			end if
		next
	end if
next
end function

function parseDecision_(result)
parseDecision_ = ""
if (result is nothing) then exit function

Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string

For Each pair In parentJson
	if (pair.Key.tostring = "changelog" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
		for each field in pair.value 'dictionary loop
			if (field.key.tostring = "histories") then 'histories is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in field.value 'array loop
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.key.tostring = "items") then 'items is an array, loop through the items, within the items loop through as a dictionary
						for each itemitem in p.value 'array loop
						bln = false 'use a toggle to start capturing status info
						for each p_item in itemitem 'dictionary loop							
							'if (p_item.Value is nothing) then
							'else
								if (p_item.Value is nothing) then
								else
								if (p_item.key.tostring = "field" and p_item.value.tostring = "Decision") then
									bln = true
									parseDecision_ = parseDecision_ & created & "||"
								end if
								end if
								if (p_item.key.tostring = "fromString" and bln) then
								if (p_item.Value is nothing) then
									parseDecision_ = parseDecision_ & "" & "||"	
								else
									parseDecision_ = parseDecision_ & p_item.value.tostring & "||"	
								end if
								end if
								if (p_item.key.tostring = "toString" and bln) then
								if (p_item.Value is nothing) then
									parseDecision_ = parseDecision_ & "" & "@@"
								else
									parseDecision_ = parseDecision_ & p_item.value.tostring & "@@"
								end if
								end if
							'end if
					'		
						next
						next
					end if
				next 'dictionary loop
				next 'array loop
			end if
		next
	end if
next
end function

function parseRiskProbability_(result)
parseRiskProbability_ = ""
if (result is nothing) then exit function

Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string

For Each pair In parentJson
	if (pair.Key.tostring = "changelog" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
		for each field in pair.value 'dictionary loop
			if (field.key.tostring = "histories") then 'histories is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in field.value 'array loop
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.key.tostring = "items") then 'items is an array, loop through the items, within the items loop through as a dictionary
						for each itemitem in p.value 'array loop
						bln = false 'use a toggle to start capturing status info
						for each p_item in itemitem 'dictionary loop							
							'if (p_item.Value is nothing) then
							'else
								if (p_item.Value is nothing) then
								else
								if (p_item.key.tostring = "field" and p_item.value.tostring = "Risk probability") then
									bln = true
									parseRiskProbability_ = parseRiskProbability_ & created & "||"
								end if
								end if
								if (p_item.key.tostring = "fromString" and bln) then
								if (p_item.Value is nothing) then
									parseRiskProbability_ = parseRiskProbability_ & "" & "||"	
								else
									parseRiskProbability_ = parseRiskProbability_ & p_item.value.tostring & "||"	
								end if
								end if
								if (p_item.key.tostring = "toString" and bln) then
								if (p_item.Value is nothing) then
									parseRiskProbability_ = parseRiskProbability_ & "" & "@@"
								else
									parseRiskProbability_ = parseRiskProbability_ & p_item.value.tostring & "@@"
								end if
								end if
							'end if
					'		
						next
						next
					end if
				next 'dictionary loop
				next 'array loop
			end if
		next
	end if
next
end function

function parseBaselineEndDate_(result)
parseBaselineEndDate_ = ""
if (result is nothing) then exit function

Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string

For Each pair In parentJson
	if (pair.Key.tostring = "changelog" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
		for each field in pair.value 'dictionary loop
			if (field.key.tostring = "histories") then 'histories is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in field.value 'array loop
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.key.tostring = "items") then 'items is an array, loop through the items, within the items loop through as a dictionary
						for each itemitem in p.value 'array loop
						bln = false 'use a toggle to start capturing status info
						for each p_item in itemitem 'dictionary loop							
							'if (p_item.Value is nothing) then
							'else
								if (p_item.Value is nothing) then
								else
								if (p_item.key.tostring = "field" and p_item.value.tostring = "Baseline end date") then
									bln = true
									parseBaselineEndDate_ = parseBaselineEndDate_ & created & "||"
								end if
								end if
								if (p_item.key.tostring = "from" and bln) then
								if (p_item.Value is nothing) then
									parseBaselineEndDate_ = parseBaselineEndDate_ & "" & "||"	
								else
									parseBaselineEndDate_ = parseBaselineEndDate_ & p_item.value.tostring & "||"	
								end if
								end if
								if (p_item.key.tostring = "to" and bln) then
								if (p_item.Value is nothing) then
									parseBaselineEndDate_ = parseBaselineEndDate_ & "" & "@@"
								else
									parseBaselineEndDate_ = parseBaselineEndDate_ & p_item.value.tostring & "@@"
								end if
								end if
							'end if
					'		
						next
						next
					end if
				next 'dictionary loop
				next 'array loop
			end if
		next
	end if
next
end function

function toYYYYMMDD(s)
'2018-05-29T12:27:48.000+0000 comes in
'20180529 goes out
toYYYYMMDD = mid(s, 1, 4) & mid(s, 6, 2) & mid(s, 9, 2)
end function
</script>
