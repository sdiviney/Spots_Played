#Static Variables

$htmlLoc = "\\server2\Spots.html"         #location of the user accessible static web page
$numSpots = 15                            #number of spots to list on the web page
$timeStamp = "C:\server2\timestamp.log"   #location of the time stamp

<#Variables to hold HTML content
    - For simplicity the CSS is included directly within the head tag
    - The META HTTP-EQUIV=`"refresh`" CONTENT=`"3`" tag causes the page to refresh itself every 3 seconds #>

$head = "<!DOCTYPE html>
<html>
<head>
<title>WideOrbit Spots</title>
<META HTTP-EQUIV=`"refresh`" CONTENT=`"3`">
<style>
body {
    background-color: black;
    }

h1 {
    text-align:center;
    font-family: arial, sans-serif;
    color: white;
}

table {
  font-family: arial, sans-serif;
  border-collapse: collapse;
  width: 100%;
  font-size: 270%;
}

th {
    background-color: #00008B;
    color: #FFF6BA;
    }


td, th {
  border: 1px solid #dddddd;
  text-align: center;
  padding: 8px;
}

tr:nth-child(odd) {
  background-color: #4169E1;
}

tr {
   color: white;
}

.short {
   background-color: #FF9900 !important;
}
</style>
</head>
<body>
<table>
  <tr>
    <th>Time Played</th>
    <th>Media ID</th>
    <th>Title</th>
    <th>Status</th>
  </tr>"

$foot = "</table>

</body>
</html>"


#Helper Functions

#get the current airlog -- a new airlog is created each day when the first spot is played after midnight
function get-file(){
    $date = get-date -format yyMMdd
    $log = "\\server1\logs\station1\AirLogs\"+$date+".air"
    #check to see if the first log file of the day has been created
    if(test-path $log){
    return $log
    } else {
        $date = (get-date).addDays(-1)
        $date = $date.ToString("yyMMdd")
        return "\\server1\logs\station1\AirLogs\"+$date+".air"
        }
}

<#takes information for one spot and returns an array with values for 1) Time Played, 2) Media ID, 3) Media Title, 4) Play Status
There is also string manipulation done on several of these to format the raw data coming out of WideOrbit's log (ex uppercase, titlecase, remove extra spaces, etc)#>
function parse-line($spot){
$time = $spot.split(",")[1]
$title = $spot.split(",")[5] -split "\s{2,}" -replace '"', ''
$title = $title[0]
$cart = $spot.split(",")[3]
$cartnum = $cart.substring(2,4)
$cat = $spot.split(",")[4]
$media = $cat+"-"+$cartnum
$status = $spot.split(",")[2] -replace ':', ''
if($status -eq "on-air"){
    $status = "Played"
    } else {$status = "Short"}
#$textInfo = (Get-Culture).TextInfo
#$status = $textInfo.ToTitleCase($status.toLower())
return $time, $media, $title.toUpper(), $status
}

<#Takes records from the airfile records and returns an array list where each array record is an array containing a spot's data
The first half of the created array list contains integers from 0-x -- rather than figuring out how to correct this, I adjusted for it in the create-tableHTML function#>
function create-tableData($records){
[System.Collections.ArrayList]$tab = @()
forEach ($record in $records){$tab.Add(@(parse-line($record)))}
return $tab
}

<#Takes the spot array list and return a variable containing the HTML for the table rows
The array list contains twice as many items as needed -- the last half are the ones with the data, hence ($tableData.Length/2)-1#>
function create-tableHtml($spotArray){
    for($i = $tableData.Length-1; $i -gt ($tableData.Length/2)-1; $i--){
        #check to see if spot was not successfully played -- if not successfully played assign the 'short' class to the row
        if($spotArray[$i][3] -eq "Played"){
        $tableBody = $tableBody+"<tr>"
        } else {$tableBody = $tableBody+"<tr class=`"short`">"}
        for($j = 0; $j -lt 4; $j++){
            $tableBody = $tableBody+"<td>"+$spotArray[$i][$j]+"</td>"
        }
        $tableBody = $tableBody+"</tr>"
    }
    return $tableBody
}

#creates the html file using the static $head and $foot variables and the dynamic $body variable
function build-html($head, $body, $foot, $htmlLoc){
out-file -FilePath $htmlLoc -InputObject $head
out-file -FilePath $htmlLoc -InputObject $body -Append
out-file -FilePath $htmlLoc -InputObject $foot -Append
}

#Controller

<#infinite loop that checks to see if a spot has been played since the last time the HTML file was created
    - If no, wait for 1 second and check again
    - If yes, update the time stamp and create a new HTML file#>
while(1){
    $airFile = get-file                               #gets the current day's "as aired" log from WideOrbit
    $lines = get-content $airFile -Tail $numSpots     #get the last numspots lines from the airfile log
    $lastTime = get-content $timeStamp                #stores the time the last spot was played
    $tableData = create-tableData($lines)             #parses the data from the airfile into an array of arrays

    if($lastTime -eq $tableData[$tableData.Length-1][0]){
        start-sleep -s 1
    } else {
            out-file $timeStamp -InputObject $tableData[$tableData.Length-1][0]
            $body = create-tableHtml($tableData)
            build-html $head $body $foot $htmlLoc
    }
}