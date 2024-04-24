<%option explicit%>
<%server.scripttimeout=60*60 '60minutes%>
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
	//alert(data);
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
Dim tmp1
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

dim arr
dim a
dim arr2
dim a2
dim arrP
dim aP

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

dim created

dim riskprobability
dim riskconsequence
dim riskprobability_history_
dim riskconsequence_history_

dim fso
dim f

dim dt
dim from_
dim to_

dim current_impact
dim previous_impact

dim decision
dim decision_history

yearmonth = request.querystring("yearmonth")
if (yearmonth = "") then yearmonth = right("0000" & year(date),4) & right("00" & Month(date),2)

tmp = dateserial(mid(yearmonth,1,4), mid(yearmonth,5,2), 1)
'format into 2018/06/30
dt = dateadd("w", -5, tmp)
from_ = right("0000" & year(dt),4) & "/" & right("00" & month(dt),2) & "/" & right("00" & day(dt),2)
dt = dateadd("m", 1, tmp)
dt = dateadd("w", -5, dt)
to_ = right("0000" & year(dt),4) & "/" & right("00" & month(dt),2) & "/" & right("00" & day(dt),2)

dim jql
jql = ""
jql = jql & " project = ""SOR"" "
jql = jql & " AND issuetype in (""Risk"")"
'jql = jql & " AND createdDate >= """ & from_ & """ "
'jql = jql & " AND createdDate <= """ & to_ & """ "

'risk where decision was changed in the given month - done below in looping through XML - reason is that Jira doesn't support this

'jql = jql & " AND ""Assigned Unit"" in (""SO MOMA"",""SO SSO"",""SO APA"",""SO AME"",""SO PMO"",""SO GDT"",""SO EAP"", ""SO ECA"", ""SO SAMEA"",""SO EECA"", ""SO AFR"", ""SO WCE"", ""SO STS"", ""SO SAM"", ""SO PDV"", ""SO OCE"", ""SO NEA"", ""SO NAM"", ""SO LAM"")"
'jql = jql & " AND "
'jql = jql & " status changed to closed during (""" & from_ & """, """ & to_ & """)"
'' jql = jql & " and issuekey = 'OM-19434' "
'jql = jql & " AND ""End date"" >= 2018-01-01 "
'' jql = jql & " AND cf[20162] >= ""2018/01/01"" "

'jql = jql & " OR status changed to closed during (""" & from_ & """"", """ & to_ & """)"
'jql = jql & " OR status changed to planned during (""" & from_ & """"", """ & to_ & """)"
' jql = jql & " )"

'response.write jql
'response.end

'for testing
'jql = " issuekey = ""OM-27842"" "

xml = getJiraItems(jql)

'Dim fso
'Dim f
'Set fso = CreateObject("scripting.filesystemobject")
'Set f = fso.createtextfile("c:\inetpub\wwwroot\so_reporting\kpi_test.txt")
'xml = f.readall
'f.Close
'Set f = Nothing
'Set fso = Nothing

'response.write xml
'response.end

'response.write replace(xml, "<", "[")
'response.write jql
'response.write xml
'for ap = lbound(arrp) to ubound(arrp)
'response.write arrp(ap) & "<br>"
'if arrp(ap) <> "" then
'	response.write "<br>" & now & "<br>"
'
'	prj = arrP(aP)

	'getJiraItems gets all the items from the project (no additional filtering) and turns it into an XML


'load the XML we generated before (to speed  up things)
'set fso = createobject("scripting.filesystemobject")
'set f = fso.opentextfile("c:\inetpub\wwwroot\src_reporting\src_kpi_detail.xml")
'xml = f.readall
'f.close
'set f = nothing
'set fso = nothing


'	response.write (replace(xml, "<" , "["))
'	response.write ("<hr noshade>")
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
		bln = false

		decision = getFieldValue(item_,"decision")
		decision_history = getFieldValue(item_,"decision_history")

		'if there is no history then go with the creation date
		if decision_history = "" then
			tmp = getFieldValue(item_,"created")
		else
			arr = split(decision_history, "@@")
			for a = ubound(arr) to lbound(arr) step -1
			if arr(a) <> "" then
				arr2 = split(arr(a), "||")
				tmp = arr2(0)
				a = -1
			end if
			next
		end if

'response.write getFieldValue(item_,"key") & "<br>"
'response.write tmp & "<br>"
'response.write replace(from_,"/","") & "<br>"
'response.write replace(to_,"/","") & "<br>"
'response.write "<hr noshade>"

		if tmp >= replace(from_,"/","") and tmp <= replace(to_,"/","") then
			bln = true
		end if

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
	'response.write ("<br>after check : " & replace(xml, "<" , "["))
	objXML.LoadXML xml

	'ok now we have a good set of xml to go with
	i = 0
	j = 0
	response.write "<span style='font-family:verdana;font-size:10pt;'><b>SOR Risk Impact Reduction</b></span><br>"
	response.write "<span style='font-family:verdana;font-size:8pt;'>"
	response.write "All issues in project SOR and of type Risk which have a decision which was last updated between " & from_ & " and " & to_ & "<br>"
	response.write "</span>"

	response.write "<table style='font-family:verdana;font-size:8pt;border-collapse:collapse;' border='1' id='table'>"
	response.write "<tr>"
	response.write "<td>" & "Pkey" & "-" & "issuenum" & "</td>"
	response.write "<td>" & "Summary" & "</td>"
	response.write "<td>" & "Assignee"   & "</td>"
	response.write "<td>" & "Requesting unit"   & "</td>"
	response.write "<td>" & "Created"   & "</td>"
	response.write "<td>" & "Calculated Impact before update" & "</td>"
	response.write "<td>" & "Calculated Impact after update" & "</td>"
	response.write "<td>" & "Reduction level" & "</td>"
	response.write "<td>" & "" & "</td>"
	response.write "<td>" & "Decision<br>Decision history" & "</td>"
	response.write "</tr>"
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	For Each item_ In nodelist
		bln = true
		'8APR2020 - quickfix if itemid = OM-65464 then skip
		'14APR2020 : not needed anymore : if getFieldValue(item_,"key") = "OM-65464" then bln = false
		if bln = true then
			'if (getFieldValue(item_,"resolution")) = "Declined"  then
			'	j = j + 1
			'	response.write "<tr style='background-color:red'>"
			'else
				response.write "<tr>"
			'end if
			response.write "<td>" & getFieldValue(item_,"key") & "</td>"
			response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
			response.write "<td>" & getFieldValue(item_,"assignee") & "</td>"
			response.write "<td>" & getFieldValue(item_,"requesting_unit") & "</td>"
			created = getFieldValue(item_,"created")
			response.write "<td>" & toDD_MM_YYYY(created) & "</td>"

			riskprobability = getFieldValue(item_,"riskprobability")
			riskconsequence = getFieldValue(item_,"riskconsequence")
			riskprobability_history_ = getFieldValue(item_,"riskprobability_history")
			riskconsequence_history_ = getFieldValue(item_,"riskconsequence_history")

			'if nothing is in
			if riskprobability = "" and riskconsequence = "" and riskprobability_history_ = "" and riskconsequence_history_ = "" then
				response.write "<td>" & "" & "</td>" 'before update
				response.write "<td>" & "" & "</td>" 'after update
				response.write "<td>" & "" & "</td>" 'Reduction level N/A in this case
			end if

			'if there is no history
			if riskprobability <> "" and riskconsequence <> "" and riskprobability_history_ = "" and riskconsequence_history_ = "" then
				'take the current impact level
				current_impact = calcRiskImpactLevel_(riskprobability,riskconsequence)
				'the previous is also the current
				previous_impact = calcRiskImpactLevel_(riskprobability,riskconsequence)

				response.write "<td>" & previous_impact & "</td>" 'before update
				response.write "<td>" & current_impact & "</td>" 'after update 'most recent - take current values
				response.write "<td>" & ImpactDelta(current_impact, previous_impact) & "</td>" 'Reduction level N/A in this case
			end if

			'if there is  history (last one is most recent) for prob but not for cons
			if riskprobability <> "" and riskconsequence <> "" and riskprobability_history_ <> "" and riskconsequence_history_ = "" then
				'take the current impact level
				current_impact = calcRiskImpactLevel_(riskprobability,riskconsequence)
				arr = split(riskprobability_history_, "@@")
				for a = ubound(arr) to lbound(arr) step -1
				if arr(a) <> "" then
					arr2 = split(arr(a), "||")
					previous_impact = calcRiskImpactLevel_(arr2(1),riskconsequence)
					if previous_impact <> current_impact then
						a = -1
					end if
				end if
				next
				response.write "<td>" & previous_impact & "</td>" 'before update
				response.write "<td>" & current_impact & "</td>" 'after update 'most recent - take current values
				response.write "<td>" & ImpactDelta(current_impact, previous_impact) & "</td>" 'Reduction level N/A in this case
			end if
			'if there is  history for cons but not for prob
			if riskprobability <> "" and riskconsequence <> "" and riskprobability_history_ = "" and riskconsequence_history_ <> "" then
				'take the current impact level
				current_impact = calcRiskImpactLevel_(riskprobability,riskconsequence)
				arr = split(riskconsequence_history_, "@@")
				for a = ubound(arr) to lbound(arr) step -1
				if arr(a) <> "" then
					arr2 = split(arr(a), "||")
					previous_impact = calcRiskImpactLevel_(riskprobability,arr2(1))
					if previous_impact <> current_impact then
						a = -1
					end if
				end if
				next
				response.write "<td>" & previous_impact & "</td>" 'before update
				response.write "<td>" & current_impact & "</td>" 'after update 'most recent - take current values
				response.write "<td>" & ImpactDelta(current_impact, previous_impact) & "</td>" 'Reduction level N/A in this case
			end if

			'if there is  history for cons and for prob
			if riskprobability <> "" and riskconsequence <> "" and riskprobability_history_ <> "" and riskconsequence_history_ <> "" then
				'take the current impact level
				current_impact = calcRiskImpactLevel_(riskprobability,riskconsequence)
				'in this case go by the dates
				'first get the dates when update were done (both in cons as in prob)
				tmp = ","
                arr = Split(riskprobability_history_, "@@")
                For a = UBound(arr) To LBound(arr) Step -1
                If arr(a) <> "" Then
                    arr2 = Split(arr(a), "||")
                    If InStr(tmp, "," & arr2(0) & ",") = 0 Then
                        tmp = tmp & arr2(0) & ","
                    End If
                End If
                Next
				arr = split(riskconsequence_history_, "@@")
				for a = ubound(arr) to lbound(arr) step -1
				if arr(a) <> "" then
					arr2 = split(arr(a), "||")
					if instr(tmp, "," & arr2(0) & ",") = 0 then
						tmp = tmp & arr2(0) & ","
					end if
				end if
				next

				'now play the dates backwards
				arr = split(tmp, ",")
				for a = ubound(arr) to lbound(arr) step -1
				if arr(a) <> "" then
					'what were the values a this time
					tmp1 = getValueatDate(riskprobability_history_, arr(a))
					tmp2 = getValueatDate(riskconsequence_history_, arr(a))
					if tmp1 = "" then tmp1 = riskprobability
					if tmp2 = "" then tmp2 = riskconsequence
					previous_impact = calcRiskImpactLevel_(tmp1,tmp2)
					if previous_impact <> current_impact then
						a = -1
					end if
				end if
				next

				response.write "<td>" & previous_impact & "</td>" 'before update
				response.write "<td>" & current_impact & "</td>" 'after update 'most recent - take current values
				response.write "<td>" & ImpactDelta(current_impact, previous_impact) & "</td>" 'Reduction level
			end if

			'for now(=debug) do show the values
			response.write "<td>Prob:" & riskprobability & "--" & replace(riskprobability_history_, "@@", "@@<br>") & "<hr>"
			response.write "Cons:" & riskconsequence & "--" & replace(riskconsequence_history_, "@@", "@@<br>") & "</td>"

			decision = getFieldValue(item_,"decision")
			decision_history = getFieldValue(item_,"decision_history")
			response.write "<td>" & decision & "<br>" & replace(decision_history, "@@", "@@<br>") & "</td>"

			response.write "</tr>"
			i = i + 1
		end if
	next

	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table"",""output.csv"");'>CSV</a><br>"
	response.write i & " items<br>" ' were set to Closed in " & yearmonth & " however " & j & " items have Quality Declined<br>"
	response.write "<br>"
%>
</body>
</html>
<%
function ImpactDelta(curr_, prev_)
ImpactDelta = ""
dim tmp
dim arr
dim a
dim arr2
tmp = ""
tmp = tmp & "Curr:Low||Prev:Low||0@@"
tmp = tmp & "Curr:Low||Prev:Medium||-1@@"
tmp = tmp & "Curr:Low||Prev:High||-2@@"
tmp = tmp & "Curr:Low||Prev:Extreme||-3@@"
tmp = tmp & "Curr:Medium||Prev:Low||1@@"
tmp = tmp & "Curr:Medium||Prev:Medium||0@@"
tmp = tmp & "Curr:Medium||Prev:High||-1@@"
tmp = tmp & "Curr:Medium||Prev:Extreme||-2@@"
tmp = tmp & "Curr:High||Prev:Low||2@@"
tmp = tmp & "Curr:High||Prev:Medium||1@@"
tmp = tmp & "Curr:High||Prev:High||0@@"
tmp = tmp & "Curr:High||Prev:Extreme||-1@@"
tmp = tmp & "Curr:Extreme||Prev:Low||3@@"
tmp = tmp & "Curr:Extreme||Prev:Medium||2@@"
tmp = tmp & "Curr:Extreme||Prev:High||1@@"
tmp = tmp & "Curr:Extreme||Prev:Extreme||0@@"
arr = split(tmp, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if lcase(trim("curr:" & curr_)) = lcase(trim(arr2(0))) then
	if lcase(trim("prev:" & prev_)) = lcase(trim(arr2(1))) then
		ImpactDelta = arr2(2)
		a = ubound(arr)
	end if
	end if
end if
next
end function

function calcRiskImpactLevel_(prob_, cons_)
calcRiskImpactLevel_ = ""
dim tmp
dim arr
dim a
dim arr2
tmp = ""
tmp = tmp & "Prob:Improbable||Cons:Acceptable||Low@@"
tmp = tmp & "Prob:Improbable||Cons:Tolerable||Medium@@"
tmp = tmp & "Prob:Improbable||Cons:Undesirable||Medium@@"
tmp = tmp & "Prob:Improbable||Cons:Intolerable||High@@"
tmp = tmp & "Prob:Possible||Cons:Acceptable||Low@@"
tmp = tmp & "Prob:Possible||Cons:Tolerable||Medium@@"
tmp = tmp & "Prob:Possible||Cons:Undesirable||High@@"
tmp = tmp & "Prob:Possible||Cons:Intolerable||Extreme@@"
tmp = tmp & "Prob:Probable||Cons:Acceptable||Medium@@"
tmp = tmp & "Prob:Probable||Cons:Tolerable||Medium@@"
tmp = tmp & "Prob:Probable||Cons:Undesirable||High@@"
tmp = tmp & "Prob:Probable||Cons:Intolerable||Extreme@@"
arr = split(tmp, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if lcase(trim("prob:" & prob_)) = lcase(trim(arr2(0))) then
	if lcase(trim("cons:" & cons_)) = lcase(trim(arr2(1))) then
		calcRiskImpactLevel_ = arr2(2)
		a = ubound(arr)
	end if
	end if
end if
next

'update 16MAR2022 - since new values have been used we need to redefine the matrix, apply this code when the legacy code doesn't return a value
if calcRiskImpactLevel_ = "" then
	tmp = ""
	tmp = tmp & "Prob:Low||Cons:Low||Low@@"
	tmp = tmp & "Prob:Medium||Cons:Low||Low@@"
	tmp = tmp & "Prob:High||Cons:Low||Medium@@"

	tmp = tmp & "Prob:Low||Cons:Medium||Low@@"
	tmp = tmp & "Prob:Medium||Cons:Medium||Medium@@"
	tmp = tmp & "Prob:High||Cons:Medium||High@@"

	tmp = tmp & "Prob:Low||Cons:High||Medium@@"
	tmp = tmp & "Prob:Medium||Cons:High||High@@"
	tmp = tmp & "Prob:High||Cons:High||Extreme@@"
	arr = split(tmp, "@@")
	for a = lbound(arr) to ubound(arr)
	if arr(a) <> "" then
		arr2 = split(arr(a), "||")
		if lcase(trim("prob:" & prob_)) = lcase(trim(arr2(0))) then
		if lcase(trim("cons:" & cons_)) = lcase(trim(arr2(1))) then
			calcRiskImpactLevel_ = arr2(2)
			a = ubound(arr)
		end if
		end if
	end if
	next
end if
end function

function showFirst(s) '20181227||11/Jan/19||2/Jan/19@@20181227||2/Jan/19||11/Jan/19 comes in
showFirst = ""
dim arr
dim arr2
dim a
dim tmp
if s = "" then
else
	arr = split(s, "@@")
	arr2 = split(arr(0), "||")
	showFirst = arr2(1)
end if
end function

function showLast(s) '20181227||11/Jan/19||2/Jan/19@@20181227||2/Jan/19||11/Jan/19 comes in
showLast = ""
dim arr
dim arr2
dim a
dim tmp
if s = "" then
else
	arr = split(s, "@@")
	arr2 = split(arr(ubound(arr)-1), "||")
	showLast = arr2(2)
end if
end function

function countChanges(s) '20181227||11/Jan/19||2/Jan/19@@20181227||2/Jan/19||11/Jan/19 comes in
countChanges = ""
dim arr
dim arr2
dim a
dim tmp
if s = "" then
else
	arr = split(s, "@@")
	countChanges = ubound(arr)
end if
end function

function mmmTOmm(s)
mmmTOmm = s
mmmTOmm = replace(mmmTOmm, "/Jan/", "01")
mmmTOmm = replace(mmmTOmm, "/Feb/", "02")
mmmTOmm = replace(mmmTOmm, "/Mar/", "03")
mmmTOmm = replace(mmmTOmm, "/Apr/", "04")
mmmTOmm = replace(mmmTOmm, "/May/", "05")
mmmTOmm = replace(mmmTOmm, "/Jun/", "06")
mmmTOmm = replace(mmmTOmm, "/Jul/", "07")
mmmTOmm = replace(mmmTOmm, "/Aug/", "08")
mmmTOmm = replace(mmmTOmm, "/Sep/", "09")
mmmTOmm = replace(mmmTOmm, "/Oct/", "10")
mmmTOmm = replace(mmmTOmm, "/Nov/", "11")
mmmTOmm = replace(mmmTOmm, "/Dec/", "12")
end function

function countDaysbetween(s) '20181227||11/Jan/19||2/Jan/19@@20181227||2/Jan/19||11/Jan/19 comes in
'update 14APR2020 : dateformat changed what comes in : yyyy-mm-dd
countDaysbetween = ""
Dim dt1
Dim dt2
dt1 = showFirst(s)
dt2 = showLast(s)
If dt1 = "" Or dt2 = "" Then
Else
	'dt1 = right("000000" & mmmTOmm(dt1),6)
	'dt2 = right("000000" & mmmTOmm(dt2),6)
	'countDaysbetween = dt1 & "#" & dt2 & "#"
	'dt1 = "20" & Mid(dt1, 5, 2) & Mid(dt1, 3, 2) & Mid(dt1, 1, 2)
	'dt2 = "20" & Mid(dt2, 5, 2) & Mid(dt2, 3, 2) & Mid(dt2, 1, 2)
	'14APR - just replace the - with space and that 'll give YYYYMMDD
	dt1 = replace(dt1, "-", "")
	dt2 = replace(dt2, "-", "")

	'countDaysbetween = countDaysbetween & dt1 & "#" & dt2 & "#"
	countDaysbetween = DateDiff("d", toDate(dt1), toDate(dt2))
End If
end function

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

function getSettoClosed(allstates)
getSettoClosed = "00000000"
dim arr
dim arr2
dim a

arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(2) = "Closed" then
	getSettoClosed = arr2(0)
	end if
end if
next
if getSettoClosed = "00000000" then getSettoClosed = ""
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

function xmlsafe(s)
xmlsafe = s
xmlsafe = replace(xmlsafe, "&", "&amp;")
xmlsafe = replace(xmlsafe, "<", "&lt;")
xmlsafe = replace(xmlsafe, ">", "&gt;")
xmlsafe = replace(xmlsafe, chr(11), "")
end function

Function getJiraItems(jql)
randomize timer
'On Error Resume Next
Dim xmlhttp
dim url
url = "https://soreporting.azurewebsites.net/so_reporting/getJiraitems.aspx?jql=" & jql & "&rnd=" & rnd & ""
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
xmlhttp.Open "GET", url, False
xmlhttp.send
getJiraItems = xmlhttp.responseText
Set xmlhttp = Nothing
End Function

function getfieldValue(xml, fld)
dim field_in_item_
For Each field_in_item_ In xml.ChildNodes
if field_in_item_.basename = fld then
getfieldValue = field_in_item_.text
end if
next
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
	if arr2(1) = "Validation" and arr2(2) <> "Done" then
	itemRejected = true
	end if
end if
next
end function

function getValueatDate(hist_, dt_)
getValueatDate = ""
dim arr
dim a
dim arr2
arr = split(hist_, "@@")
for a = ubound(arr) to lbound(arr) step -1
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(0) = dt_ then
		getValueatDate = arr2(1)
		a = -1
	end if
end if
next
end function

function toDD_MM_YYYY(s)
toDD_MM_YYYY = s
if toDD_MM_YYYY = "" then exit function
'20180526 comes in
'26/05/2018 goes out
toDD_MM_YYYY = mid(s, 7, 2) & "/" & mid(s, 5, 2) &"/" & mid(s, 1, 4)
end function
%>