<%option explicit%>
<%server.scripttimeout=10*60 '10minutes%>
<!--#include file="jsonObject.class.asp" -->
<html>
<head>
<script>
function downloadCSV(tbl,filename)
{
	//loop through table rows & cols
	var obj = document.getElementById(tbl);
	var row;
	var col;
	var csv='';
	for (var i = 0; row = obj.rows[i]; i++)
	{
	   //iterate through rows
	   //rows would be accessed using the "row" variable assigned in the for loop
	   for (var j = 0; col = row.cells[j]; j++)
	   {
	     //iterate through columns
	     //columns would be accessed using the "col" variable assigned in the for loop
	     var str = col.innerText;
	     str = str.replace("#", encodeURIComponent("#"));
	     csv+='"'+str+'";';
	   }
	   csv+='\n';
	}
	filename = filename;
	if (!csv.match(/^data:text\/csv/i))
	{
	csv = 'data:text/csv;charset=utf-8,' + encodeURI('\uFEFF'+csv); //force the file to be UTF8
	}
	var data = csv;
	var link = document.createElement('a');
	link.setAttribute('href', data);
	link.setAttribute('download', filename);
	document.body.appendChild(link); //needed for FF
	link.click();
}
</script>
</head>
<body style='font-family:verdana;font-size:8pt;'>
<%
Dim tmp
Dim tmp2
Dim url
dim id
dim prj
'dim arr
dim issuetype
dim event_
dim sum_
dim desc_
dim i
dim j

dim days
dim tot_days
dim vali
dim invali

dim arr
dim a
dim arr2
dim a2
dim arrP
dim aP

dim wi_set_to_done
dim wi_rejected
dim wi_rejected_keys
dim wi_missed_deadline
dim wi_missed_deadline_keys
dim wi_committed
dim wi_committed_keys

dim xml
dim objXML
set objXML = createobject("MSXML.DOMDocument")
Dim item_
Dim field_in_item_
dim nodelist

dim label
dim hist_
dim comp
dim bln
dim bln2

dim duedate
dim target_duedate
dim sprint
dim sprintenddate

dim yearmonth
yearmonth = request.querystring("yearmonth")
dim settodone

dim fso
dim f

arrP = split("STSPES,STSCNT,MOEBIUS,RES", ",")
'arr = split("STSPES", ",")
'arr = split("STSCNT,STSPES", ",")
arrP = split(request.querystring("project"), ",")
'response.write request.querystring("project") & "<br>"
'response.write ubound(arr) & "<br>"

for ap = lbound(arrp) to ubound(arrp)
response.write arrp(ap) & "<br>"
if arrp(ap) <> "" then
	response.write "<br>" & now & "<br>"

	prj = arrP(aP)

	'getJiraItems gets all the items from the project (no additional filtering) and turns it into an XML
	xml = getJiraItems(prj,yearmonth)

'load the XML we generated before (to speed  up things)
'set fso = createobject("scripting.filesystemobject")
'set f = fso.opentextfile("c:\inetpub\wwwroot\src_reporting\src_kpi_detail.xml")
'xml = f.readall
'f.close
'set f = nothing
'set fso = nothing


	'response.write (replace(xml, "<" , "["))
	'response.write ("<hr noshade>")
	objXML.LoadXML xml

''save the the output for now (just to speed up devs)
'set fso = createobject("scripting.filesystemobject")
'set f = fso.createtextfile("c:\inetpub\wwwroot\src_reporting\src_kpi_detail.xml")
'f.writeline xml
'f.close
'set f = nothing
'set fso = nothing


	xml = ""
	xml = xml & "<item_result>" & vbcrlf
	'now rebuild the XML using addiotnal filters
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	i = 1
	For Each item_ In nodelist
		bln = true
		For Each field_in_item_ In item_.ChildNodes
			'check if correct issuetype
			if field_in_item_.BaseName = "issuetype" then
			if instr("||Epic||Initiative||", "||" & field_in_item_.Text & "||") > 0 then
				bln = false
			end if
			end if
			'check if correct components
			if field_in_item_.BaseName = "components" then
			arr = split(field_in_item_.Text, "||")
			for a = lbound(arr) to ubound(arr)
			if arr(a) <> "" then
				if instr(lcase("||Events organization||Hosting Sharepoint sites||R&D||STS internal processes||STS Newsletter||TD platform||"),lcase("||" & arr(a) & "||")) > 0 then
					bln = false
				end if
			end if
			next
			end if
			'check if done in this month
			if field_in_item_.BaseName = "transitions_history" then
			    bln2 = False
			    arr = Split(field_in_item_.Text, "@@")
			    For a = LBound(arr) To UBound(arr)
			    If arr(a) <> "" Then
				arr2 = Split(arr(a), "||")
				'For a2 = LBound(arr2) To UBound(arr2)
				'If arr2(a2) <> "" Then
				    If arr2(2) = "Done" Then
				    If Mid(arr2(0), 1, 6) = yearmonth Then
					bln2 = True
				    End If
				    End If
				'End If
				'Next
			    End If
			    Next
			    If bln2 = True Then ' it is set to done in this month
				'do nothing
			    Else
				'not a good item
				bln = False
			    End If
			end if

		Next
		if bln = true then
		xml = xml & "<item>" & vbcrlf
		For Each field_in_item_ In item_.ChildNodes
			xml = xml & "<" & field_in_item_.BaseName & ">" & xmlsafe("" & field_in_item_.Text) & "</" & field_in_item_.BaseName & ">" & vbcrlf
		Next
		xml = xml & "</item>" & vbcrlf
		end if

		i = i + 1
	Next
	xml = xml & "</item_result>" & vbcrlf
	'response.write (replace(xml, "<" , "["))
	objXML.LoadXML xml

	'ok now we have a good set of xml to go with

	'first metric Completion missrate
	'get all the items set to done in this month
	'AND they have an intended enddate (via target due or via enddate sprint)

	'COMPARE via target due or via completedate in Sprint
	response.write "Overview for " & prj & "<br>"
	response.write "<hr noshade>"

	response.write "Completion Missrate<br>"
	response.write "<i>Show all the items that have an intended enddate in this month (via target due or via enddate sprint) and compare with date set to Done</i><br>"
	i = 0
	j = 0
	response.write "<table style='font-family:verdana;font-size:8pt;border-collapse:collapse;' border='1' id='table_" & prj & "_1'>"
	response.write "<tr>"
	response.write "<td>" & "pkey" & "-" & "issuenum" & "</td>"
	response.write "<td>" & "issuetype" & "</td>"
	response.write "<td>" & "component" & "</td>"
	'response.write "<td>" & "id"      & "</td>"
	response.write "<td>" & "summary" & "</td>"
	response.write "<td>" & "state"   & "</td>"
	response.write "<td>" & "created"   & "</td>"
	response.write "<td>" & "updated"   & "</td>"
	response.write "<td>" & "resolutiondate"   & "</td>"
	response.write "<td>" & "state change overview"   & "</td>"
	response.write "<td>" & "date set to done"   & "</td>"
	'response.write "<td>" & "state change sum of days"   & "</td>"
	'response.write "<td>" & "validation tracking to in validation occurences"   & "</td>"
	'response.write "<td>" & "in validation to anything but done/Validation Tracking occurences"   & "</td>"
	response.write "<td>" & "item duedate"   & "</td>"
	response.write "<td>" & "item target duedate"   & "</td>"
	response.write "<td>" & "duedate via linked sprint(s)"   & "</td>"
	'response.write "<td>" & "duedate via linked release (fixVersions)"   & "</td>"
	'response.write "<td>" & "compiled duedate"   & "</td>"
	response.write "</tr>"

	Set nodelist = objXML.getElementsByTagName("item_result/*")
	For Each item_ In nodelist
		bln = true
		'determine enddate
		target_duedate = getFieldValue(item_,"target_duedate")
		if target_duedate = "" then
			target_duedate = getFieldValue(item_,"sprintcompletedate")
		end if
		settodone = getSettoDone(getFieldValue(item_,"transitions_history"))

		if target_duedate = "" then
			bln = false
		end if

		if bln = true then
		if settodone > target_duedate then
			j = j + 1
		response.write "<tr style='background-color:red'>"
		else
		response.write "<tr>"
		end if
		response.write "<td>" & getFieldValue(item_,"key") & "</td>"
		response.write "<td>" & getFieldValue(item_,"issuetype") & "</td>"
		response.write "<td>" & getFieldValue(item_,"components") & "</td>"
		response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
		response.write "<td>" & getFieldValue(item_,"status") & "</td>"
		'response.write "<td>" & getFieldValue(item_,"components") & "</td>"
		response.write "<td>" & getFieldValue(item_,"created") & "</td>"
		response.write "<td>" & getFieldValue(item_,"updated") & "</td>"
		response.write "<td>" & getFieldValue(item_,"resolutiondate") & "</td>"
		response.write "<td>" & replace(getFieldValue(item_,"transitions_history"), "@@", "<br>") & "</td>"
		response.write "<td>" & getSettoDone(getFieldValue(item_,"transitions_history")) & "</td>"
		'response.write "<td>" & "" & "</td>" 'validation tracking to in validation occurences
		'response.write "<td>" & "" & "</td>" 'in validation to anything but done/Validation Tracking occurences
		response.write "<td>" & getFieldValue(item_,"duedate") & "</td>"
		response.write "<td>" & getFieldValue(item_,"target_duedate") & "</td>"
		response.write "<td>" & getFieldValue(item_,"sprintcompletedate") & "</td>"

		'For Each field_in_item_ In item_.ChildNodes
		'response.write "<td>" & field_in_item_.text & "</td>"
		'next
		response.write "</tr>"
		i = i + 1
		end if
	next
	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table_" & prj & "_1"",""output_1.csv"");'>CSV</a><br>"
	response.write i & " items targetted to end in " & yearmonth & " however " & j & " items did not make the deadline<br>"
	response.write "<br>"

	response.write "Rejection Rate<br>"
	response.write "<i>When items go from Validation to a state different from Done we count this as a Rejection</i><br>"
	i = 0
	j = 0
	response.write "<table style='font-family:verdana;font-size:8pt;border-collapse:collapse;' border='1' id='table_" & prj & "_2'>"
	response.write "<tr>"
	response.write "<td>" & "pkey" & "-" & "issuenum" & "</td>"
	response.write "<td>" & "issuetype" & "</td>"
	response.write "<td>" & "component" & "</td>"
	response.write "<td>" & "summary" & "</td>"
	response.write "<td>" & "state"   & "</td>"
	response.write "<td>" & "created"   & "</td>"
	response.write "<td>" & "updated"   & "</td>"
	response.write "<td>" & "resolutiondate"   & "</td>"
	response.write "<td>" & "state change overview"   & "</td>"
	response.write "<td>" & "date set to done"   & "</td>"
	response.write "<td>" & "item duedate"   & "</td>"
	response.write "<td>" & "item target duedate"   & "</td>"
	response.write "<td>" & "duedate via linked sprint(s)"   & "</td>"
	response.write "</tr>"
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	For Each item_ In nodelist
		bln = true
		if bln = true then
			if itemRejected(getFieldValue(item_,"transitions_history")) then
				j = j + 1
				response.write "<tr style='background-color:red'>"
			else
				response.write "<tr>"
			end if
			response.write "<td>" & getFieldValue(item_,"key") & "</td>"
			response.write "<td>" & getFieldValue(item_,"issuetype") & "</td>"
			response.write "<td>" & getFieldValue(item_,"components") & "</td>"
			response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
			response.write "<td>" & getFieldValue(item_,"status") & "</td>"
			'response.write "<td>" & getFieldValue(item_,"components") & "</td>"
			response.write "<td>" & getFieldValue(item_,"created") & "</td>"
			response.write "<td>" & getFieldValue(item_,"updated") & "</td>"
			response.write "<td>" & getFieldValue(item_,"resolutiondate") & "</td>"
			response.write "<td>" & replace(getFieldValue(item_,"transitions_history"), "@@", "<br>") & "</td>"
			response.write "<td>" & getSettoDone(getFieldValue(item_,"transitions_history")) & "</td>"
			'response.write "<td>" & "" & "</td>" 'validation tracking to in validation occurences
			'response.write "<td>" & "" & "</td>" 'in validation to anything but done/Validation Tracking occurences
			response.write "<td>" & getFieldValue(item_,"duedate") & "</td>"
			response.write "<td>" & getFieldValue(item_,"target_duedate") & "</td>"
			response.write "<td>" & getFieldValue(item_,"sprintcompletedate") & "</td>"

			'For Each field_in_item_ In item_.ChildNodes
			'response.write "<td>" & field_in_item_.text & "</td>"
			'next
			response.write "</tr>"
			i = i + 1
		end if
	next

	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table_" & prj & "_2"",""output_2.csv"");'>CSV</a><br>"
	response.write i & " items were set to done in " & yearmonth & " however " & j & " items have been rejected<br>"
	response.write "<br>"

	response.write "Cycle Time<br>"
	response.write "<i>Count the # of days an item is In Progress (all states except Cancelled/Done/Backlog/)</i><br>"
	i = 0
	j = 0
	response.write "<table style='font-family:verdana;font-size:8pt;border-collapse:collapse;' border='1' id='table_" & prj & "_3'>"
	response.write "<tr>"
	response.write "<td>" & "pkey" & "-" & "issuenum" & "</td>"
	response.write "<td>" & "issuetype" & "</td>"
	response.write "<td>" & "component" & "</td>"
	response.write "<td>" & "summary" & "</td>"
	response.write "<td>" & "state"   & "</td>"
	response.write "<td>" & "created"   & "</td>"
	response.write "<td>" & "updated"   & "</td>"
	response.write "<td>" & "resolutiondate"   & "</td>"
	response.write "<td>" & "state change overview"   & "</td>"
	response.write "<td>" & "date set to done"   & "</td>"
	response.write "<td>" & "item duedate"   & "</td>"
	response.write "<td>" & "item target duedate"   & "</td>"
	response.write "<td>" & "duedate via linked sprint(s)"   & "</td>"
	response.write "<td>" & "#days in progress"   & "</td>"
	response.write "</tr>"
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	For Each item_ In nodelist
		bln = true
		if bln = true then
			tmp = countInProgress(getFieldValue(item_,"created"), getFieldValue(item_,"transitions_history"))
			j = j + tmp
			response.write "<tr>"
			response.write "<td>" & getFieldValue(item_,"key") & "</td>"
			response.write "<td>" & getFieldValue(item_,"issuetype") & "</td>"
			response.write "<td>" & getFieldValue(item_,"components") & "</td>"
			response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
			response.write "<td>" & getFieldValue(item_,"status") & "</td>"
			'response.write "<td>" & getFieldValue(item_,"components") & "</td>"
			response.write "<td>" & getFieldValue(item_,"created") & "</td>"
			response.write "<td>" & getFieldValue(item_,"updated") & "</td>"
			response.write "<td>" & getFieldValue(item_,"resolutiondate") & "</td>"
			response.write "<td>" & replace(getFieldValue(item_,"transitions_history"), "@@", "<br>") & "</td>"
			response.write "<td>" & getSettoDone(getFieldValue(item_,"transitions_history")) & "</td>"
			'response.write "<td>" & "" & "</td>" 'validation tracking to in validation occurences
			'response.write "<td>" & "" & "</td>" 'in validation to anything but done/Validation Tracking occurences
			response.write "<td>" & getFieldValue(item_,"duedate") & "</td>"
			response.write "<td>" & getFieldValue(item_,"target_duedate") & "</td>"
			response.write "<td>" & getFieldValue(item_,"sprintcompletedate") & "</td>"
			response.write "<td>" & tmp & "</td>"

			'For Each field_in_item_ In item_.ChildNodes
			'response.write "<td>" & field_in_item_.text & "</td>"
			'next
			response.write "</tr>"
			i = i + 1
		end if
	next

	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table_" & prj & "_3"",""output_3.csv"");'>CSV</a><br>"
	response.write i & " items that were set to done in " & yearmonth & " were all together " & j & " days in progress<br>"
	response.write "<br>"

	'first get the items , take the info we need and store in XML for later processing
	'we already exclude Epic/Initiave and a series of components vie the JQL
	'if instr(lcase(",Epic,Initiative,"),"," & lcase(trim(item.Value("fields").Value("issuetype").Value("name"))) & ",") > 0 then
	'if instr(lcase(",,"),"," & lcase(trim(comp)) & ",") > 0 then
	'was is set to done in this month?

	'hist_ = getTransitions(strunpw, item.Value("key"))
	'settodone = getSettoDone(hist_) 'function returns 00000000 if no set to done date was found
	'bln = true
	'if mid(settodone,1,6) <> yearmonth then
	'	bln = false
	'end if

if false then
	wi_set_to_done = 0
	wi_rejected = 0
	wi_rejected_keys = ""
	wi_committed = 0
	wi_committed_keys = ""
	wi_missed_deadline = 0
	wi_missed_deadline_keys = ""




	response.write "Overview for " & prj & "<br>"
	response.write "<hr noshade>"

	response.write "Completion Missrate<br>"
	response.write "<i>Show all the items that have an intended enddate in this month (via target due or via enddate sprint) and compare with date set to Done</i><br>"

	response.write "<table style='font-family:verdana;font-size:8pt;border-collapse:collapse;' border='1' id='table_" & prj & "'>"
	response.write "<tr>"
	response.write "<td>" & "pkey" & "-" & "issuenum" & "</td>"
	response.write "<td>" & "issuetype" & "</td>"
	response.write "<td>" & "component" & "</td>"
	'response.write "<td>" & "id"      & "</td>"
	response.write "<td>" & "summary" & "</td>"
	response.write "<td>" & "state"   & "</td>"
	response.write "<td>" & "created"   & "</td>"
	response.write "<td>" & "updated"   & "</td>"
	response.write "<td>" & "resolutiondate"   & "</td>"
	response.write "<td>" & "state change overview"   & "</td>"
	response.write "<td>" & "state change sum of days"   & "</td>"
	response.write "<td>" & "validation tracking to in validation occurences"   & "</td>"
	response.write "<td>" & "in validation to anything but done/Validation Tracking occurences"   & "</td>"
	response.write "<td>" & "item duedate"   & "</td>"
	response.write "<td>" & "item target duedate"   & "</td>"
	response.write "<td>" & "duedate via linked sprint(s)"   & "</td>"
	'response.write "<td>" & "duedate via linked release (fixVersions)"   & "</td>"
	'response.write "<td>" & "compiled duedate"   & "</td>"
	response.write "<td>" & "date set to done"   & "</td>"
	response.write "</tr>"

	'jsonList = getJiraItems(strunpw, prj)
	'response.write jsonList
	'if false then
	set outputObj = jsonObj.parse(jsonList)
	for each item in outputObj.Value("issues").items
		bln = true

		if false then
		'if issuetype = epic or Initiative the do NOT show
		if instr(lcase(",Epic,Initiative,"),"," & lcase(trim(item.Value("fields").Value("issuetype").Value("name"))) & ",") > 0 then
			bln = false
		else
			'get component info
			comp = ""
			if not isnull(item.Value("fields").Value("components")) then
				comp = item.Value("fields").Value("components").Serialize
				if comp <> "[]" then comp = getJson(comp, "name")
			end if
			'if component one of below then do NOT show
			if instr(lcase(",Events organization,Hosting Sharepoint sites,R&D,STS internal processes,STS Newsletter,TD platform,"),"," & lcase(trim(comp)) & ",") > 0 then
				bln = false
			else
				'get state history
				hist_ = getTransitions(strunpw, item.Value("key"))

				'was is set to Done this month?
				settodone = getSettoDone(hist_) 'function returns 00000000 if no set to done date was found
				if mid(settodone,1,6) <> yearmonth then
					bln = false
				end if

				'NEW 23 APR 2018 :
				'was the item committed for this month?
				'set to done in this month (latest linked sprint was set to this month)
				'if not isnull(item.Value("fields").Value("customfield_11860")) then
				'	sprintenddate = getSprintEndDate(item.Value("fields").Value("customfield_11860").Serialize)
				'end if

				'if not donethismonth(hist_, yearmonth) then
				'	bln = false
				'end if
			end if
		end if
		end if

		'get state history
		bln = true
		hist_ = getTransitions(strunpw, item.Value("key"))
		'was is set to Done this month?
		settodone = getSettoDone(hist_) 'function returns 00000000 if no set to done date was found
		if mid(settodone,1,6) <> yearmonth then
			bln = false
		end if

		if bln then

		wi_set_to_done = wi_set_to_done + 1

		duedate = "" & toYYYYMMDD(item.Value("fields").Value("duedate"))
		target_duedate = "" & toYYYYMMDD(item.Value("fields").Value("customfield_19662"))
		'if duedate = "" then
		'	if not isnull(item.Value("fields").Value("customfield_11860")) then
		'		duedate = getSprintEndDate(item.Value("fields").Value("customfield_11860").Serialize)
		'	end if
		'end if
		'Due/Committed Date: ‘Target end date’ associated directly to item or alternatively via Sprint Linkage (‘Planned End Date’)

		'missed deadline
		'if the date set to done is passed duedate
		if settodone <> "" then
		if target_duedate <> "" then
		if settodone > target_duedate then
			wi_missed_deadline = wi_missed_deadline + 1
			wi_missed_deadline_keys = wi_missed_deadline_keys & item.Value("key") & ","
		end if
		end if
		end if

		'rejection (check the history of the state changes)
		'if state goes from validation to <>done then rejected
		if itemRejected(hist_) then
			wi_rejected = wi_rejected + 1
			wi_rejected_keys = wi_rejected_keys & item.Value("key") & ","
		end if
		'ask rogrido

		'boxplot


		response.write "<tr>"
		response.write "<td>" & item.Value("key") & "</td>"
		response.write "<td>" & item.Value("fields").Value("issuetype").Value("name") & "</td>"
		response.write "<td>" & comp & "</td>"
		response.write "<td>" & item.Value("fields").Value("summary") & "</td>"
		response.write "<td>" & item.Value("fields").Value("status").Value("name") & "</td>"
		response.write "<td>" & toYYYYMMDD(item.Value("fields").Value("created")) & "</td>"
		response.write "<td>" & toYYYYMMDD(item.Value("fields").Value("updated")) & "</td>"
		response.write "<td>" & toYYYYMMDD(item.Value("fields").Value("resolutiondate")) & "</td>"

		'response.write "<td>" & "state change overview"   & "</td>" 'calc
		response.write "<td>" & replace(hist_, "@@", "<br>") & "</td>"

		response.write "<td>" & countSumOfDays(hist_) & "</td>" 'calc
		response.write "<td>" & countVali(hist_) & "</td>" 'calc 'validation tracking to in validation occurences
		response.write "<td>" & countInvali(hist_) & "</td>" 'calc '"in validation to anything but done/Validation Tracking occurences"
		response.write "<td>" & duedate & "</td>" 'TODO???
		response.write "<td>" & target_duedate & "</td>" 'TODO???


		'response.write "<td>" & "duedate via linked sprint(s)"   & "</td>"
		'response.write "<td>" & item.Value("fields").Value("customfield_11860") & "</td>"
		'response.write "<td>" & getSprints(strunpw, item.Value("key")) & "</td>"
		'response.write "<td>" & getSprints(strunpw, item.Value("key")) & "</td>"
		tmp = ""
		'for each label in item.Value("fields").Value("customfield_11860").items
		'	tmp = tmp & label(0)
		'next
		tmp = ""
		if not isnull(item.Value("fields").Value("customfield_11860")) then
			'fetch the related sprint info in a seperate call
			tmp = item.Value("fields").Value("customfield_11860").Serialize
	''data looks like this:
	''customfield_11860":["com.atlassian.greenhopper.service.sprint.Sprint@5d417ded[id=12782,rapidViewId=3520,state=CLOSED,name=QMS Sprint Test,startDate=2018-03-08T13:57:33.570Z,endDate=2018-03-09T13:57:00.000Z,completeDate=2018-03-09T12:05:36.443Z,sequence=12782]"],"updated":"2018-03-08T13:56:42.000+0000"
	''customfield_11860":["com.atlassian.greenhopper.service.sprint.Sprint@5d417ded[id=12782,rapidViewId=3520,state=CLOSED,name=QMS Sprint Test,startDate=2018-03-08T13:57:33.570Z,endDate=2018-03-09T13:57:00.000Z,completeDate=2018-03-09T12:05:36.443Z,sequence=12782]"],"updated":"2018-03-15T10:47:07.000+0000"
	''customfield_11860":["com.atlassian.greenhopper.service.sprint.Sprint@5d417ded[id=12782,rapidViewId=3520,state=CLOSED,name=QMS Sprint Test,startDate=2018-03-08T13:57:33.570Z,endDate=2018-03-09T13:57:00.000Z,completeDate=2018-03-09T12:05:36.443Z,sequence=12782]"],"updated":"2018-03-09T12:04:56.000+0000"
	'		for each item2 in item.Value("fields").Value("customfield_11860").items
	'			tmp = tmp & "x" & "<br>" 'Value("fields").Value("customfield_11860")
	'			tmp = tmp & item2.Value("0").text & "<br>"
	'		next
			tmp = tmp & "<br>" & getSprintEndDate(item.Value("fields").Value("customfield_11860").Serialize)

		end if
		response.write "<td>" & tmp & "</td>"

		'tmp = ""
		'if not isnull(item.Value("fields").Value("fixVersions")) then
		'	'fetch the related sprint info in a seperate call
		'	tmp = item.Value("fields").Value("fixVersions").Serialize
		'end if
		'response.write "<td>" & tmp & "</td>"

		'response.write "<td>" & duedate & "</td>"
		response.write "<td>" & settodone & "</td>"

		response.write "</tr>"

		end if 'done this month
	next

	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table_" & prj & """,""output.csv"");'>CSV</a>"
	response.write "<br>"
	'end if

	response.write "<hr noshade>"

	response.Write "Rejection Rate<br>"
	response.write "<i>Show all the items that have an intended enddate in this month (via target due or via enddate sprint) and compare with date set to Done</i><br>"

	'Rodrigo vroeg om deze templates door te sturen zodat je een zicht kreeg op de data die we nodig hebben voor zijn KPIs.
	response.write "<b>Input for PI Metric Sheet Template Box Plot Percentage 2014.xlsx</b>" & "<br>"
	response.write "Under construction" & "<br>"
	'Voor de boxplot heb ik de data nodig zoals in de ‘data-raw’ sheet en daarmee ga ik dan aan de slag voor de berekeningen in de ‘data-calculations’
	'sheet.
	'Vraag hier is dus of je al onmiddellijk de berekeningen zou kunnen maken vanuit JIRA.
	response.write "<b>Input for PI Work Item On Time Delivery Miss Rate 2018.xlsx</b>" & "<br>"
	responsE.write     "WI set to done for project " & prj & " for month " & yearmonth & " : " & wi_set_to_done & "<br>"
	responsE.write "WI missed deadline for project " & prj & " for month " & yearmonth & " : " & wi_missed_deadline
	if wi_missed_deadline_keys <> "" then
		response.write " (" & mid(wi_missed_deadline_keys, 1, len(wi_missed_deadline_keys)-1) & ")"
	end if
	response.write "<br>"
	'Voor de OTD en Rejection Rates heb ik enkel de cijfers nodig voor de groene kolommen in de data sheet.
	response.write "<b>Input for PI Work Item Rejection Rate 2018.xlsx</b>" & "<br>"
	responsE.write "WI set to done for project " & prj & " for month " & yearmonth & " : " & wi_set_to_done & "<br>"
	responsE.write    "WI rejected for project " & prj & " for month " & yearmonth & " : " & wi_rejected
	if wi_rejected_keys <> "" then
		response.write " (" & mid(wi_rejected_keys, 1, len(wi_rejected_keys)-1) & ")"
	end if
	response.write "<br>"
	'Voor de OTD en Rejection Rates heb ik enkel de cijfers nodig voor de groene kolommen in de data sheet.
	response.write "<hr noshade>"
end if

	response.write "<br>" & now & "<br>"
	response.write "<hr noshade>"

end if 'prj loop
next 'prj loop

'response.write tmp
%>
</body>
</html>
<%
function toYYYYMMDD(s)
toYYYYMMDD = s
if toYYYYMMDD <> "" then
	toYYYYMMDD = mid(s,1,4) & mid(s,6,2) & mid(s,9,2)
end if
end function

Function Lng2Dt(n)
if n = "" then
Lng2Dt = ""
else
n=n/1000
Lng2Dt = dateserial(1900,1,1) + (n + 2 ^ 31) / 86400
'Lng2Dt = right("0000" & year(lng2dt),4) & right("00" & month(lng2dt),2) & right("00" & day(lng2dt),2) & right("00" & hour(lng2dt),2) & right("00" & minute(lng2dt),2) & right("00" & second(lng2dt),2)
end if
End Function

function toDate(s)
'toDate = timeserial(mid(s,1,4),mid(s,5,2),mid(s,7,2),mid(s,9,2),mid(s,11,2),mid(s,13,2))
toDate = dateserial(mid(s,1,4),mid(s,5,2),mid(s,7,2))
end function

Function URLEncode(ByVal str)
 Dim strTemp, strChar
 Dim intPos, intASCII
 strTemp = ""
 strChar = ""
 For intPos = 1 To Len(str)
  intASCII = Asc(Mid(str, intPos, 1))
  If intASCII = 32 Then
   strTemp = strTemp & "+"
  ElseIf ((intASCII < 123) And (intASCII > 96)) Then
   strTemp = strTemp & Chr(intASCII)
  ElseIf ((intASCII < 91) And (intASCII > 64)) Then
   strTemp = strTemp & Chr(intASCII)
  ElseIf ((intASCII < 58) And (intASCII > 47)) Then
   strTemp = strTemp & Chr(intASCII)
  Else
   strChar = Trim(Hex(intASCII))
   If intASCII < 16 Then
    strTemp = strTemp & "%0" & strChar
   Else
    strTemp = strTemp & "%" & strChar
   End If
  End If
 Next
 URLEncode = strTemp
End Function

function getCustomFieldValue(id, prj, fld, cn)
getCustomFieldValue = ""
'todo depeding on the type we need to select the correct column in the customfieldvalue table (stringvalue, numbervalue, textvalue, datevalue
dim strsql
dim objRec
set objrec = createobject("adodb.recordset")

strsql = ""
strsql = strsql & " select cfo.customvalue "
strsql = strsql & " from jiraissue i, project p, customfield cf, customfieldoption cfo, customfieldvalue cfv "
strsql = strsql & " where i.issuenum = " & id & " "
strsql = strsql & " and i.project = p.id "
strsql = strsql & " and p.pkey = '" & prj & "' "
strsql = strsql & " and cf.cfname = '" & fld & "' "
strsql = strsql & " and cf.id = cfv.customfield "
strsql = strsql & " and cfv.issue = i.id "
strsql = strsql & " and cfv.stringvalue = cfo.id"
strsql = strsql & " and cfo.customfield = cf.id "
strsql = strsql & " "
strsql = strsql & " "
strsql = strsql & " "
strsql = strsql & " "

objrec.open strsql, cn, 1, 3
if not objrec.eof then
	getCustomFieldValue = "" & objrec("customvalue")
end if
objrec.close

set objrec = nothing
end function

function getAllLinks(id, prj, cn)
'function receives eg IM-2 and returns IM-3 Countainmet, IM-4 Correction
getAllLinks = ""
dim strsql
dim objRec
set objrec = createobject("adodb.recordset")

strsql = ""
strsql = strsql & " select p2.pkey, i2.issuenum, it2.pname, ilt.linkname "
strsql = strsql & " from jiraissue i, project p, issuelink il, issuelinktype ilt, jiraissue i2, project p2, issuetype it2 "
strsql = strsql & " where i.issuenum = " & id & ""
strsql = strsql & " and i.project = p.id "
strsql = strsql & " and p.pkey = '" & prj & "'"
strsql = strsql & " and il.source = i.id "
strsql = strsql & " and il.destination = i2.id "
strsql = strsql & " and il.linktype = ilt.id "
strsql = strsql & " and i2.project = p2.id "
strsql = strsql & " and i2.issuetype = it2.id "
strsql = strsql & "  "
strsql = strsql & "  "
strsql = strsql & "  "

strsql = ""
strsql = strsql & " select * from issuelink "
strsql = strsql & " where source = (select i.id from jiraissue i, project p where i.issuenum = " & id & " and i.project = p.id and p.pkey = '" & prj & "') "
strsql = strsql & " union "
strsql = strsql & " select * from issuelink "
strsql = strsql & " where destination = (select i.id from jiraissue i, project p where i.issuenum = " & id & " and i.project = p.id and p.pkey = '" & prj & "') "

objrec.open strsql, cn, 1, 3
while not objrec.eof
	getAllLinks = getAllLinks & objrec("pkey") & "-" & objrec("issuenum") & "||" & objrec("pname") & "||" & objrec("linkname") & "@@"
	objrec.movenext
wend
objrec.close

set objrec = nothing
end function

function getIssueType(id, prj, cn)
getIssueType = ""
dim strsql
dim objRec
set objrec = createobject("adodb.recordset")

	strsql = ""
	strsql = strsql & " select it.pname "
	strsql = strsql & " from jiraissue i, project p, issuetype it "
	strsql = strsql & " where i.issuenum = " & id & ""
	strsql = strsql & " and i.project = p.id "
	strsql = strsql & " and p.pkey = '" & prj & "'"
	strsql = strsql & " and i.issuetype = it.id"

objrec.open strsql, cn, 1, 3
if not objrec.eof then
	getIssueType = "" & objrec("pname")
end if
objrec.close

set objrec = nothing
end function

function getEvent(id, prj, cn)
getEvent = ""
dim strsql
dim objRec
set objrec = createobject("adodb.recordset")

	strsql = ""
	strsql = strsql & " select i.created, i.updated "
	strsql = strsql & " from jiraissue i, project p "
	strsql = strsql & " where i.issuenum = " & id & ""
	strsql = strsql & " and i.project = p.id "
	strsql = strsql & " and p.pkey = '" & prj & "'"

objrec.open strsql, cn, 1, 3
if not objrec.eof then
	if objrec("created") = objrec("updated") then
		getEvent = "create"
	else
		getEvent = "update"
	end if
end if
objrec.close

set objrec = nothing
end function

function getOldFieldValue(id, prj, fld, cn)
getOldFieldValue = ""
dim strsql
dim objRec
set objrec = createobject("adodb.recordset")

If fld = "description" Then
fld = "CAST(i.description AS CHAR(10000) CHARACTER SET utf8)"
Else
fld = "i." & fld
End If

	strsql = ""
	strsql = strsql & " select cg.*, ci.*  "
	strsql = strsql & " from jiraissue i, project p, changegroup cg , changeitem ci "
	strsql = strsql & " where i.project = p.id and i.issuenum = " & id & " "
	strsql = strsql & " and p.pkey = '" & prj & "' "
	strsql = strsql & " and i.id = cg.issueid "
	strsql = strsql & " and cg.id = ci.groupid "
	strsql = strsql & " and field = '" & fld & "' "
	strsql = strsql & " order by cg.created desc "

'we're intrested in the first row
'row looks like this: blah blah , oldvalue, oldstring, newvalue, newstring

objrec.open strsql, cn, 1, 3
if not objrec.eof then
	getOldFieldValue = objrec("oldstring")
end if
objrec.close

set objrec = nothing
end function

function getFieldValue(id, prj, fld, cn)
getFieldValue = ""
dim strsql
dim objRec
set objrec = createobject("adodb.recordset")

If fld = "description" Then
fld = "CAST(i.description AS CHAR(10000) CHARACTER SET utf8)"
Else
fld = "i." & fld
End If

	strsql = ""
	strsql = strsql & " select " & fld & " "
	strsql = strsql & " from jiraissue i, project p "
	strsql = strsql & " where i.issuenum = " & id & ""
	strsql = strsql & " and i.project = p.id "
	strsql = strsql & " and p.pkey = '" & prj & "'"

objrec.open strsql, cn, 1, 3
if not objrec.eof then
	getFieldValue = objrec(0)
end if
objrec.close

set objrec = nothing
end function

function getfieldValue(xml, fld)
dim field_in_item_
For Each field_in_item_ In xml.ChildNodes
if field_in_item_.basename = fld then
getfieldValue = field_in_item_.text
end if
next
end function

sub log_(s)
dim fso
dim f
set fso = createobject("scripting.filesystemobject")
set f = fso.opentextfile(LOGFILE, 8)
f.writeline now & " - " & s
f.close
set f = nothing
set fso = nothing
end sub

function getJson(json, key)
Dim arr
Dim a
Dim arr2
Dim p1
Dim p2
'{"id":"1886722","key":"IM-35","self":"https://preprodjira.tomtomgroup.com/rest/api/2/issue/1886722"}
p1 = InStr(json, """" & key & """:""")
p2 = InStr(p1, json, """,""")
If p2 = 0 Then
    p2 = InStr(p1, json, """}")
End If
If p1 <> 0 And p2 <> 0 Then
    p1 = p1 + Len("""" & key & """:""")
    getJson = Mid(json, p1, p2 - p1)
End If
end function

' Function to decode string from Base64
Public Function base64_decode(ByVal strIn)
Dim w1, w2, w3, w4, n, strOut
For n = 1 To Len(strIn) Step 4
    w1 = mimedecode(Mid(strIn, n, 1))
    w2 = mimedecode(Mid(strIn, n + 1, 1))
    w3 = mimedecode(Mid(strIn, n + 2, 1))
    w4 = mimedecode(Mid(strIn, n + 3, 1))
    If w2 >= 0 Then _
    strOut = strOut + _
        Chr(((w1 * 4 + Int(w2 / 16)) And 255))
    If w3 >= 0 Then _
    strOut = strOut + _
        Chr(((w2 * 16 + Int(w3 / 4)) And 255))
    If w4 >= 0 Then _
    strOut = strOut + _
        Chr(((w3 * 64 + w4) And 255))
Next
base64_decode = strOut
End Function

Private Function mimedecode(ByVal strIn)
If Len(strIn) = 0 Then
    mimedecode = -1: Exit Function
Else
    mimedecode = InStr(Base64Chars, strIn) - 1
End If
End Function

' Functions for encoding string to Base64
Public Function base64_encode(ByVal strIn)
Dim c1, c2, c3, w1, w2, w3, w4, n, strOut
For n = 1 To Len(strIn) Step 3
    c1 = Asc(Mid(strIn, n, 1))
    c2 = Asc(Mid(strIn, n + 1, 1) + Chr(0))
    c3 = Asc(Mid(strIn, n + 2, 1) + Chr(0))
    w1 = Int(c1 / 4): w2 = (c1 And 3) * 16 + Int(c2 / 16)
    If Len(strIn) >= n + 1 Then
    w3 = (c2 And 15) * 4 + Int(c3 / 64)
    Else
    w3 = -1
    End If
    If Len(strIn) >= n + 2 Then
    w4 = c3 And 63
    Else
    w4 = -1
    End If
    strOut = strOut + mimeencode(w1) + mimeencode(w2) + _
          mimeencode(w3) + mimeencode(w4)
Next
base64_encode = strOut
End Function

Private Function mimeencode(ByVal intIn)
Dim Base64Chars
Base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
    "abcdefghijklmnopqrstuvwxyz" & _
    "0123456789" & _
    "+/"
If intIn >= 0 Then
    mimeencode = Mid(Base64Chars, intIn + 1, 1)
Else
    mimeencode = ""
End If
End Function

Function getJiraItems(project,yearmonth)
'response.write "https://soreporting.azurewebsites.net/src_reporting/getJiraitems.aspx?project=" & project & "&yearmonth=" & yearmonth & "&rnd=" & rnd & ""
randomize timer
'On Error Resume Next
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
xmlhttp.Open "GET", "https://soreporting.azurewebsites.net/src_reporting/getJiraitems.aspx?project=" & project & "&yearmonth=" & yearmonth & "&rnd=" & rnd & "", False
xmlhttp.send
getJiraItems = xmlhttp.responseText
Set xmlhttp = Nothing
End Function

Function getSprints(unpw, key)
'On Error Resume Next
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
'xmlhttp.Open "GET", "https://jira.tomtomgroup.com/rest/api/2/issue/OTS-169561", False
xmlhttp.Open "GET", "https://jira.tomtomgroup.com/rest/agile/1.0/issue/" & key & "", False 'or &maxResults=10&startAt=11
'xmlhttp.Open "GET", "https://jira.tomtomgroup.com/rest/api/2/search?jql=project=""" & project & """&maxResults=-1", False 'or &maxResults=10&startAt=11
xmlhttp.setRequestHeader "Content-Type", "application/json"
xmlhttp.setRequestHeader "Authorization", "Basic " & unpw
'If Err.Number <> 0 Then Debug.Print Err.Description
xmlhttp.send
'If Err.Number <> 0 Then Debug.Print Err.Description
getSprints = xmlhttp.responseText
'response.write getJiraItems
Set xmlhttp = Nothing
End Function

Function getTransitions(unpw, key)
'On Error Resume Next
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
'xmlhttp.Open "GET", "https://jira.tomtomgroup.com/rest/api/2/issue/OTS-169561", False
'xmlhttp.Open "GET", "https://jira.tomtomgroup.com/rest/agile/1.0/issue/" & key & "", False 'or &maxResults=10&startAt=11
xmlhttp.Open "GET", "https://jira.tomtomgroup.com/rest/api/2/issue/" & key & "?expand=changelog&fields=status", False 'or &maxResults=10&startAt=11
'xmlhttp.Open "GET", "https://jira.tomtomgroup.com/rest/api/2/search?jql=project=""" & project & """&maxResults=-1", False 'or &maxResults=10&startAt=11
xmlhttp.setRequestHeader "Content-Type", "application/json"
xmlhttp.setRequestHeader "Authorization", "Basic " & unpw
'If Err.Number <> 0 Then Debug.Print Err.Description
xmlhttp.send
'If Err.Number <> 0 Then Debug.Print Err.Description
getTransitions = xmlhttp.responseText
'response.write getJiraItems
Set xmlhttp = Nothing

'now get the transisitons in a
dim jsonObj
dim outputObj
dim item
dim item2
dim createddate
set jsonObj = new JSONobject
set outputObj = jsonObj.parse(getTransitions)
getTransitions = ""
for each item in outputObj.Value("changelog").Value("histories").items
	createddate = toYYYYMMDD(item.Value("created"))
	'now loop within the items field
	for each item2 in item.Value("items").items
		if item2.Value("field") = "status" then
			getTransitions = getTransitions & createddate & "||" & item2.Value("fromString") & "||" & item2.Value("toString") & "@@"
		end if
	next
	'getTransitions = getTransitions & item.Value("id") & "-"
	'getTransitions = getTransitions & toYYYYMMDD(item.Value("created"))
	'getTransitions = getTransitions & "<br>"
next
set jsonObj = nothing
set outputObj = nothing
End Function

Function getJiraItem(unpw, key)
'On Error Resume Next
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
'xmlhttp.Open "GET", "https://jira.tomtomgroup.com/rest/api/2/issue/OTS-169561", False
xmlhttp.Open "GET", "https://jira.tomtomgroup.com/rest/api/2/issue/" & key, False 'or &maxResults=10&startAt=11
xmlhttp.setRequestHeader "Content-Type", "application/json"
xmlhttp.setRequestHeader "Authorization", "Basic " & unpw
'If Err.Number <> 0 Then Debug.Print Err.Description
xmlhttp.send
'If Err.Number <> 0 Then Debug.Print Err.Description
getJiraItem = xmlhttp.responseText
Set xmlhttp = Nothing
End Function

'google for : ; taken from https://stackoverflow.com/questions/1019223/any-good-libraries-for-parsing-json-in-classic-asp
Function JSONtoXML(jsonText)
  Dim idx, max, ch, mode, xmldom, xmlelem, xmlchild, name, value

  Set xmldom = CreateObject("Microsoft.XMLDOM")
  xmldom.loadXML "<xml/>"
  Set xmlelem = xmldom.documentElement

  max = Len(jsonText)
  mode = 0
  name = ""
  value = ""
  While idx < max
    idx = idx + 1
    ch = Mid(jsonText, idx, 1)
    Select Case mode
    Case 0 ' Wait for Tag Root
      Select Case ch
      Case "{"
        mode = 1
      End Select
    Case 1 ' Wait for Attribute/Tag Name
      Select Case ch
      Case """"
        name = ""
        mode = 2
      Case "{"
        Set xmlchild = xmldom.createElement("tag")
        xmlelem.appendChild xmlchild
        xmlelem.appendchild xmldom.createTextNode(vbCrLf)
       xmlelem.insertBefore xmldom.createTextNode(vbCrLf), xmlchild
        Set xmlelem = xmlchild
      Case "["
        Set xmlchild = xmldom.createElement("tag")
        xmlelem.appendChild xmlchild
        xmlelem.appendchild xmldom.createTextNode(vbCrLf)
        xmlelem.insertBefore xmldom.createTextNode(vbCrLf), xmlchild
        Set xmlelem = xmlchild
      Case "}"
        Set xmlelem = xmlelem.parentNode
      Case "]"
        Set xmlelem = xmlelem.parentNode
      End Select
    Case 2 ' Get Attribute/Tag Name
      Select Case ch
      Case """"
        mode = 3
      Case Else
        name = name + ch
      End Select
    Case 3 ' Wait for colon
      Select Case ch
      Case ":"
        mode = 4
      End Select
    Case 4 ' Wait for Attribute value or Tag contents
      Select Case ch
      Case "["
        Set xmlchild = xmldom.createElement(name)
        xmlelem.appendChild xmlchild
        xmlelem.appendchild xmldom.createTextNode(vbCrLf)
        xmlelem.insertBefore xmldom.createTextNode(vbCrLf), xmlchild
        Set xmlelem = xmlchild
        name = ""
        mode = 1
      Case "{"
        Set xmlchild = xmldom.createElement(name)
        xmlelem.appendChild xmlchild
        xmlelem.appendchild xmldom.createTextNode(vbCrLf)
        xmlelem.insertBefore xmldom.createTextNode(vbCrLf), xmlchild
        Set xmlelem = xmlchild
        name = ""
        mode = 1
      Case """"
        value = ""
        mode = 5
      Case " "
      Case Chr(9)
      Case Chr(10)
      Case Chr(13)
      Case Else
        value = ch
        mode = 7
      End Select
    Case 5
      Select Case ch
      Case """"
        xmlelem.setAttribute name, value
        mode = 1
      Case "\"
        mode = 6
      Case Else
        value = value + ch
      End Select
   Case 6
      value = value + ch
      mode = 5
    Case 7
      If Instr("}], " & Chr(9) & vbCr & vbLf, ch) = 0 Then
        value = value + ch
      Else
        xmlelem.setAttribute name, value
        mode = 1
        Select Case ch
        Case "}"
          Set xmlelem = xmlelem.parentNode
        Case "]"
          Set xmlelem = xmlelem.parentNode
        End Select
      End If
    End Select
  Wend

  Set JSONtoXML = xmlDom
End Function

function getSettoDone(allstates)
getSettoDone = "00000000"
dim arr
dim arr2
dim a

arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(2) = "Done" then
	getSettoDone = arr2(0)
	end if
end if
next
if getSettoDone = "00000000" then getSettoDone = ""
end function

function donethismonth(allstates, ym)
donethismonth = false
'allstates = date||from||to@@...
dim arr
dim arr2
dim a

arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	'response.write ym & " - " & arr2(0) & "-" & arr2(1) & "-" & arr2(2) & "-" & "<br>"
	if arr2(2) = "Done" then
	if mid(arr2(0), 1, 6) = ym then
		donethismonth = true
	end if
	end if
end if
next
end function

function countVali(allstates)
'take first date, take last date, subtract, if 0 then result is 1
dim arr
dim arr2
dim a
dim dt1
dim dt2
countVali = 0
arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(1) = "Validation Tracking" then
	if arr2(2) = "In Validation" then
		countVali = countVali + 1
	end if
	end if
end if
next
end function

function countInvali(allstates)
'take first date, take last date, subtract, if 0 then result is 1
dim arr
dim arr2
dim a
dim dt1
dim dt2
countInvali = 0
arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(1) = "Validation Tracking" and arr2(2) <> "Done" then
	if arr2(1) = "Validation Tracking" and arr2(2) <> "Validation Tracking" then
		countInvali = countInvali + 1
	end if
	end if
end if
next

'if objrec2("oldstring") = "In Validation" and objrec2("newstring") <> "Done" then
'if objrec2("oldstring") = "In Validation" and objrec2("newstring") <> "Validation Tracking" then
'	invali = invali + 1
'end if
'end if

end function

Function CountinProgress(created, allstates)
'loop through all transitions, if an item is 'in progress (all but Done, Cancelled, Backlog)' then count the # of days
Dim arr
Dim arr2
Dim a
Dim dt
Dim prevdate

'allstates = created & "||" & "<new>" & "||" & "Open" & "@@"
CountinProgress = 0
arr = Split(allstates, "@@")
prevdate = DateSerial(Mid(created, 1, 4), Mid(created, 5, 2), Mid(created, 7, 2))
For a = LBound(arr) To UBound(arr)
If arr(a) <> "" Then
    arr2 = Split(arr(a), "||")
    'did we move to an not in progres state?
    If arr2(2) <> "Done" And arr2(2) <> "Cancelled" And arr2(2) <> "Backlog" Then
        dt = DateSerial(Mid(arr2(0), 1, 4), Mid(arr2(0), 5, 2), Mid(arr2(0), 7, 2))
        CountinProgress = CountinProgress + DateDiff("d", prevdate, dt)
    End If
    prevdate = DateSerial(Mid(arr2(0), 1, 4), Mid(arr2(0), 5, 2), Mid(arr2(0), 7, 2))
    'If a = LBound(arr) Then dt1 = DateSerial(Mid(arr2(0), 1, 4), Mid(arr2(0), 5, 2), Mid(arr2(0), 7, 2))
    'If a = UBound(arr) - 1 Then dt2 = DateSerial(Mid(arr2(0), 1, 4), Mid(arr2(0), 5, 2), Mid(arr2(0), 7, 2))
    'End If

End If
Next
'CountinProgress = DateDiff("d", dt1, dt2)
'If CountinProgress = 0 Then CountinProgress = 1
End Function

function countSumOfDays(allstates)
'take first date, take last date, subtract, if 0 then result is 1
dim arr
dim arr2
dim a
dim dt1
dim dt2

arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if a = lbound(arr) then dt1 = dateserial(mid(arr2(0),1,4),mid(arr2(0),5,2),mid(arr2(0),7,2))
	if a = ubound(arr)-1 then dt2 = dateserial(mid(arr2(0),1,4),mid(arr2(0),5,2),mid(arr2(0),7,2))
end if
next
countSumOfDays = datediff("d", dt1, dt2)
if countSumOfDays = 0 then countSumOfDays = 1
end function

function getSprintEndDate(s)
'multiple sprint may be attached, take the latest one
'IN : ["com.atlassian.greenhopper.service.sprint.Sprint@790ac54f[id=11467,rapidViewId=3248,state=CLOSED,name=SDP Sprint 1,startDate=2017-11-13T06:12:41.559Z,endDate=2017-11-27T06:12:00.000Z,completeDate=2017-11-29T09:20:45.503Z,sequence=11467]","com.atlassian.greenhopper.service.sprint.Sprint@728aac4a[id=12185,rapidViewId=3248,state=CLOSED,name=SDP Sprint 2,startDate=2018-02-12T12:02:27.361Z,endDate=2018-02-26T12:02:00.000Z,completeDate=2018-02-23T11:26:35.645Z,sequence=12185]"]
dim arr
dim arr2
dim a
dim tmp
getSprintEndDate = "00000000"
arr = split(s, ",")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "=")
	if arr2(0) = "endDate" then
		tmp = mid(arr2(1),1,4)&mid(arr2(1),6,2)&mid(arr2(1),9,2)
		'response.write s & "##" & tmp & "<br>"

		'since we might have multiple sprints check if this one is the latest
		if tmp > getSprintEndDate then getSprintEndDate = tmp
	end if
end if
next
'if nothing found then return empty
if getSprintEndDate = "00000000" then getSprintEndDate = ""
end function

function itemRejected(allstates)
'in = 20180226||Open||Breakdown@@20180226||Breakdown||Breakdown Done@@20180302||Breakdown Done||Development@@20180302||Development||Development Done@@20180305||Development Done||Validation@@20180305||Validation||Done
dim arr
dim arr2
dim a
itemRejected = false
arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if lcase(arr2(1)) = lcase("Validation") and lcase(arr2(2)) <> lcase("Done") then
	'17DEC2019 - from Validation to On hold is not a rejection (cfr mail Ryno Botes)
	if lcase(arr2(1)) = lcase("Validation") and lcase(arr2(2)) = lcase("On hold") then
	else
	itemRejected = true
	end if
	end if
end if
next
end function

function xmlsafe(s)
xmlsafe = s
xmlsafe = replace(xmlsafe, "&", "&amp;")
xmlsafe = replace(xmlsafe, "<", "&lt;")
xmlsafe = replace(xmlsafe, ">", "&gt;")
xmlsafe = replace(xmlsafe, chr(11), "")
end function

%>