<# FrequencyLog.ps1
# Version 2.0 (Tested on windows 8.1/10 pro and Win 7 SP1 Ultimate)
# Reduced Minimum repetitions from 5 to 2
Reduced Minimun repetitions var on line 25 from 5 to 2 (you can move it to 1).
Corrected the minimun repetitions variables relation instead of "greather than" to be "greather equal to"
Corrected name of the Html file from Graphics.htm to "Graphics-YYYYMMdd.html"
Either way remember that the data will be refreshed every time you run the script.
#Pending add clean up variables section
Added Security events into the graphics
Requires Powershell Version 5
#version 3.0
	Added log file 
	Corrected bugs in version 2.0 in date
	Added Clean up section
	Jose Gabriel Ortega (j0rt3g4@j0rt3g4.com / https://www.j0rt3g4.com)
	CEO J0rt3g4 Consulting Services
#>

param(
	[Parameter(mandatory=$false,position=0)][string]$computerp
)
########  VARIABLES  ##########

###Global:
$TimeStart=Get-Date
$global:ScriptLocation = "C:\Windows\LTSVC\Packages\FrequencyLog"
$global:R=4.0
$global:Ite=3000
$global:trans=1000
$global:Version=$PSVersionTable.PSVersion.Major
$MinimunRepetitions = 2
$global:DefaultLog = "$global:ScriptLocation\Frequency.log"

###Local:
$html="$global:ScriptLocation\FrequencyLog.htm"
$jsonsystem="$global:ScriptLocation\system.json"
$jsonapplication="$global:ScriptLocation\application.json"
$jsonsecurity="$global:ScriptLocation\security.json"

###CleanupVariable
$CleanUpGlobal=@()
$cleanUpLocal=@()
##Local & Global Cleanup
$CleanUpGlobal+="ScriptLocation"
$CleanUpGlobal+="R"
$CleanUpGlobal+="Ite"
$CleanUpGlobal+="trans"
$cleanUpLocal+="html"
$cleanUpLocal+="jsonsystem"
$cleanUpLocal+="jsonapplication"
$cleanUpLocal+="jsonsecurity"
$cleanUpLocal+="MinimunRepetitions"


########  FUNCTIONS  ##########

function Write-Log{
        [CmdletBinding()]
        #[Alias('wl')]
        [OutputType([int])]
        Param
        (
            # The string to be written to the log.
            [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0)] [ValidateNotNullOrEmpty()] [Alias("LogContent")] [string]$Message,
            # The path to the log file.
            [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true,Position=1)] [Alias('LogPath')] [string]$Path=$DefaultLog,
            [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true,Position=2)] [ValidateSet("Error","Warn","Info","Load","Execute")] [string]$Level="Info",
            [Parameter(Mandatory=$false)] [switch]$NoClobber
        )

     Process{
        
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Warning "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
            }

        # If attempting to write to a log file in a folder/path that doesn't exist
        # to create the file include path.
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            $NewLogFile = New-Item $Path -Force -ItemType File
            }

        else {
            # Nothing to see here yet.
            }

        # Now do the logging and additional output based on $Level
        switch ($Level) {
            'Error' {
                Write-Warning $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ERROR: `t $Message" | Out-File -FilePath $Path -Append
                }
            'Warn' {
                Write-Warning $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") WARNING: `t $Message" | Out-File -FilePath $Path -Append
                }
            'Info' {
                Write-Host $Message -ForegroundColor Green
                Write-Verbose $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") INFO: `t $Message" | Out-File -FilePath $Path -Append
                }
            'Load' {
                Write-Host $Message -ForegroundColor Magenta
                Write-Verbose $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") LOAD: `t $Message" | Out-File -FilePath $Path -Append
                }
            'Execute' {
                Write-Host $Message -ForegroundColor Green
                Write-Verbose $Message
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") EXEC: `t $Message" | Out-File -FilePath $Path -Append
                }
            }
    }
}
function CheckExists{
	param(
		[Parameter(mandatory=$true,position=0)]$itemtocheck,
		[Parameter(mandatory=$true,position=1)]$colection
	)
	BEGIN{
		$item=$null
		$exist=$false
	}
	PROCESS{
		foreach($item in $colection){
			if($item.EventID -eq $itemtocheck){
				$exist=$true
				break;
			}
		}

	}
	END{
		return $exist
	}

}
function CheckCount{
	param(
		[Parameter(mandatory=$true,position=0)]$itemtocheck,
		[Parameter(mandatory=$true,position=1)]$colection
	)
	BEGIN{
		$item=$null
		$count=0
	}
	PROCESS{
		foreach($item in $colection){
			
			if($item.EventID -eq $itemtocheck){
				$count++
			}
		}

	}
	END{
		return $count
	}

}
function LastWrittenTime{
	param(
		[Parameter(mandatory=$true,position=0)]$colection,
		[Parameter(mandatory=$true,position=1)]$EventID

	)
	BEGIN{
		$filterCollection= $colection | Where-Object{ $_.EventID -eq $EventID}
	}
	PROCESS{
		$previous = $filterCollection[0].TimeWritten
		$last = $filterCollection[0].TimeWritten
		foreach($item in $filterCollection){
			if($item.TimeWritten -lt $previous){
				$previous =$item.TimeWritten
			}
			if($item.TimeWritten -gt $last){
				$last = $item.TimeWritten
			}

		}

	}
	END{
		$output = New-Object psobject -Property @{
			first= $previous
			last= $last
		}
		return $output
	}

}
function CreateJS{
	param(
		[Parameter(mandatory=$true,position=0)]$SystemData,
		[Parameter(mandatory=$true,position=1)]$AppData,
		[Parameter(mandatory=$true,position=2)]$SecData,
		[Parameter(mandatory=$false,position=3)]$JSFile="$global:ScriptLocation\graph.js"
	)
	$JSContent+='
	var chart;
	var SystemData ='
	$JSContent+=$SystemData
	$JSContent+=';
	var appData ='
	$JSContent+=$AppData
	$JSContent+=';
	var secData ='
	$JSContent+=$SecData
	$JSContent+=';
        AmCharts.ready(function () {
		var chart = AmCharts.makeChart("chartdiv",{
			"type": "serial",
			"dataProvider": SystemData,
			"categoryField": "EventID",
			"startDuration": 1,
			//axes
			"valueAxes": [ {
				"dashLength": 5,
				"title": "Frequency of the event",
				"axisAlpha": 0,
			}],
			"gridAboveGraphs": false,
			
			"graphs": [ {
				"balloonText": "EventID [[category]]</br>Repeated: <b>[[value]]</b> times</br>Source: [[Source]]</br>[[Message]]</br>First on:<b>[[FirstTimeWritten]]</b></br>Last on:<b>[[LastTimeWritten]]</b> </br> <b class=' + "Yellow" +'>[[EntryType]]</b>",
				"fillAlphas": 0.8,
				"lineAlpha": 0.2,
				"type": "column",
				"valueField": "Count",
				"colorField": "color"
			}],
			"chartCursor": {
				"categoryBalloonEnabled": false,
				"cursorAlpha": 0,
				"zoomable": false
			},
			
			"categoryAxis": {
				"gridPosition": "start",
				"gridAlpha": 0,
				"fillAlpha": 1,
				"labelRotation" : 60,
				"fillColor": "#EEEEEE",
				"gridPosition": "start"
			},
			"creditsPosition" : "top-right",
			"export": {
				"enabled": true
			}
    });

		var chart2 = AmCharts.makeChart("chart2div",{
			"type": "serial",
			"dataProvider":appData,
			"categoryField": "EventID",
			"startDuration": 1,
			//axes
			"valueAxes": [ {
				"dashLength": 5,
				"title": "Frequency of the event",
				"axisAlpha": 0,
			}],
			"gridAboveGraphs": false,
			
			"graphs": [ {
				"balloonText": "EventID [[category]]</br>Repeated: <b>[[value]]</b> times</br>Source: [[Source]]</br>[[Message]]</br>First on:<b>[[FirstTimeWritten]]</b></br>Last on:<b>[[LastTimeWritten]]</b> </br> <b class=' + "Yellow" +'>[[EntryType]]</b>",
				"fillAlphas": 0.8,
				"lineAlpha": 0.2,
				"type": "column",
				"valueField": "Count",
				"colorField": "color"
			}],
			"chartCursor": {
				"categoryBalloonEnabled": false,
				"cursorAlpha": 0,
				"zoomable": false
			},
			
			"categoryAxis": {
				"gridPosition": "start",
				"gridAlpha": 0,
				"fillAlpha": 1,
				"labelRotation" : 60,
				"fillColor": "#EEEEEE",
				"gridPosition": "start"
			},
			"creditsPosition" : "top-right",
			"export": {
				"enabled": true
			}
    });

		var chart3 = AmCharts.makeChart("chart3div",{
			"type": "serial",
			"dataProvider":secData,
			"categoryField": "EventID",
			"startDuration": 1,
			//axes
			"valueAxes": [ {
				"dashLength": 5,
				"title": "Frequency of the event",
				"axisAlpha": 0,
			}],
			"gridAboveGraphs": false,
			
			"graphs": [ {
				"balloonText": "EventID [[category]]</br>Repeated: <b>[[value]]</b> times</br>Source: [[Source]]</br>[[Message]]</br>First on:<b>[[FirstTimeWritten]]</b></br>Last on:<b>[[LastTimeWritten]]</b> </br> <b class=' + "Yellow" +'>[[EntryType]]</b>",
				"fillAlphas": 0.8,
				"lineAlpha": 0.2,
				"type": "column",
				"valueField": "Count",
				"colorField": "color"
			}],
			"chartCursor": {
				"categoryBalloonEnabled": false,
				"cursorAlpha": 0,
				"zoomable": false
			},
			
			"categoryAxis": {
				"gridPosition": "start",
				"gridAlpha": 0,
				"fillAlpha": 1,
				"labelRotation" : 60,
				"fillColor": "#EEEEEE",
				"gridPosition": "start"
			},
			"creditsPosition" : "top-right",
			"export": {
				"enabled": true
			}
    });

			//Original
		/*
        // SERIAL CHART
        chart = new AmCharts.AmSerialChart();
        chart.dataProvider = SystemData;
        chart.categoryField = "EventID";
        chart.startDuration = 1;


        // AXES
        // category
        var categoryAxis = chart.categoryAxis;
        categoryAxis.labelRotation = 60; // this line makes category values to be rotated
        categoryAxis.gridAlpha = 0;
        categoryAxis.fillAlpha = 1;
        categoryAxis.fillColor = "#EEEEEE";
        categoryAxis.gridPosition = "start";

        // value
        var valueAxis = new AmCharts.ValueAxis();
        valueAxis.dashLength = 5;
        valueAxis.title = "Frequency of the event";
        valueAxis.axisAlpha = 0;
        chart.addValueAxis(valueAxis);

        // GRAPH
        var graph = new AmCharts.AmGraph();
        graph.valueField = "Count";
        graph.colorField = "color";
        graph.balloonText = "<b>[[category]]: [[value]]</b>";
        graph.type = "column";
        graph.lineAlpha = 0;
        graph.fillAlphas = 1;
		
        chart.addGraph(graph);

        // CURSOR
        var chartCursor = new AmCharts.ChartCursor();
        chartCursor.cursorAlpha = 0;
        chartCursor.zoomable = false;
        chartCursor.categoryBalloonEnabled = false;
        chart.addChartCursor(chartCursor);

        chart.creditsPosition = "top-right";

        // WRITE
        chart.write("chartdiv");
		*/
});'

	$JSContent | Out-File $JSFile -Encoding utf8
}
function CreateHtml{
	param(
		[Parameter(mandatory=$true,position=0)][string]$filename
	)
	BEGIN{}
	PROCESS{
	$htmlFile=$null
	$htmlFile+=@"
<!DOCTYPE HTML>
<html lang="en-US">
<head>
	<meta charset="UTF-8">
	<title>Frecuency on Events</title>
	<link rel="stylesheet" href="style.css" type="text/css">
    <script src="amcharts.js" type="text/javascript"></script>
    <script src="serial.js" type="text/javascript"></script>
	<!-- scripts for exporting chart as an image -->
        <!-- Exporting to image works on all modern browsers except IE9 (IE10 works fine) -->
        <!-- Note, the exporting will work only if you view the file from web server -->
        <!--[if (!IE) | (gte IE 10)]> -->
        <script type="text/javascript" src="export.min.js"></script>
        <link  type="text/css" href="export.css" rel="stylesheet">
        <!-- <![endif]-->
	<script type="text/javascript" src="graph.js"></script> <!-- This file is created by a function in the script -->
</head>
<body>
<h4>Report ran on: $(Get-Date -Format "F") </h4>
	<center>
		<h2>Frecuency of events in Event Viewer: System</h2>
	</center>
	<div id="chartdiv" style="width: 100%; height: 415px;"></div>
	<div>&nbsp;</div>
	<center>
		<h2>Frecuency of events in Event Viewer: Application</h2>
	</center>
	<div id="chart2div" style="width: 100%; height: 415px;"></div>
	<div>&nbsp;</div>
	<center>
		<h2>Frecuency of events in Event Viewer: Security</h2>
	</center>
		<div id="chart3div" style="width: 100%; height: 415px;"></div>
<h4>Report ran on: $(Get-Date -Format "F") </h4>
</body>
</html>
"@
		}
	END{
	$htmlFile | Out-File $filename -Encoding utf8
		}
}
function RdnNumber{
	BEGIN{
		#Logistic Map Random Number Generator
		[int]$i=0;
	}
	PROCESS{
		Start-Sleep -m 100
		
		[double]$x1=0.0

		#LogisticMap
		[Random]$Rdn = New-Object System.Random
		$x0 = $Rdn.NextDouble()

		$total = $global:Ite + $global:trans
		for($c=0 ; $c -lt $total ; $c++){
			$x1 = $global:R * $x0 * (1.0 - $x0)
			$x0=$x1
		}
	}
	END{
		return [convert]::ToInt32([Math]::Floor([decimal]($x1*21.0)),10)
	}
}
function AssignColor{
	
	BEGIN{
		[int]$i = RdnNumber
		[string]$out=$null
	}
	PROCESS{
		switch($i){
			0 {
				$out="#AE3B3B"
				break;
			}
			1 {
				$out="#FFABAB"
				break;
			}
			2 {
				$out="#D86D6D"
				break;
			}
			3{
				$out="#851818"
				break;
			}
			4{
				$out="#5A0101"
				break;
			}
			5{
				$out="#AE743B"
				break;
			}
			6{
				$out="#FFD5AB"
				break;
			}
			7{
				$out="#D8A26D"
				break;
			}
			8{
				$out="#854E18"
				break;
			}
			9{
				$out="#5A2D01"
				break;
			}
			10{
				$out="#275E6C"
				break;
			}
			11{
				$out="#6E98A1"
				break;
			}
			12{
				$out="#457986"
				break;
			}
			13{
				$out="#114652"
				break;
			}
			14{
				$out="#022E38"
				break;
			}
			15{
				$out="#2F8B2F"
				break;
			}
			16{
				$out="#8ACE8A"
				break;
			}
			17{
				$out="#57AD57"
				break;
			}
			18{
				$out="#136B13"
				break;
			}
			19{
				$out="#014801"
				break;
			}
			default {
				$out="#275F6C"
				break;
			}
		}
	}
	END{
		return $out
	}


}
function FrequencyData{
	param(
		[Parameter(mandatory=$true,position=0)][ValidateSet("System","Application","security")][string]$log,
		[Parameter(mandatory=$true,position=1)][int]$minimun,
		[Parameter(mandatory=$false,position=2)]$computer="localhost"
	)
	BEGIN{
		$events=@()
		if($log -ne "security"){
			if($computer -eq "localhost"){
				$Allinfo = Get-EventLog -LogName "$log" -EntryType Error,Warning |  select *  # | Where-Object{ $_.EventID -eq 7045}
			}
			else{
				$Allinfo = Get-EventLog -LogName "$log" -EntryType Error,Warning  -ComputerName $computer | select * # | Where-Object{ $_.EventID -eq 7045}
			}
		}
		else{
			if($computer -eq "localhost"){
				$Allinfo = Get-EventLog -LogName $log -EntryType Successaudit,Failureaudit 
				}
			else{
				$Allinfo = Get-EventLog -LogName $log -EntryType Successaudit,Failureaudit -ComputerName $computer |  select * 
			}
		}
	}
	PROCESS{
		foreach($event in $Allinfo){
			if(! (checkExists $event.EventID $events)){
				$TimeObject = LastWrittenTime -colection $Allinfo  -EventID $event.EventID
				$TempObj = New-Object PSObject -Property @{
					EventID = $event.EventID
					Count = CheckCount $event.EventID $Allinfo
					color = AssignColor
					EntryType = $event.EntryType
					Source =$event.Source
					Message = $event.Message
					FirstTimeWritten=[convert]::ToString($TimeObject.first)
					LastTimeWritten=[convert]::ToString( $TimeObject.last)
				}
				if($TempObj.Count -ge $minimun){
					$events+=$TempObj
				}
			}
		}
	}
	END{
		return $events
	}
}
function ShowTimeMS{
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$True,position=0,mandatory=$true)]	[datetime]$timeStart,
        [parameter(ValueFromPipeline=$True,position=1,mandatory=$true)]	[datetime]$timeEnd
    )
    BEGIN {}
    PROCESS {
    write-Log -Level Info -Message  "Stamping time"
    
    $diff = New-TimeSpan $TimeStart $TimeEnd
    #Write-Verbose "Timediff= $diff"
    $miliseconds = $diff.TotalMilliseconds
    }
    END{
        Write-Log -Level Info -Message  "Total Time in miliseconds is: $miliseconds ms"
    }
}


########  SCRIPT  ##########

	Write-Log -level Info -Message "Starting Script"
	Write-Progress -Id 1 -Activity MainScript -PercentComplete 0
	Write-Log -Level Load -Message "Getting System log Information"
	#system
	if($computerp -ne $null){
		$SystemData = FrequencyData -log System -minimun $MinimunRepetitions
	}
	else{
		$SystemData = FrequencyData -log System -minimun $MinimunRepetitions -computer $computerp
	}

	Write-Progress -Id 1 -Activity MainScript -PercentComplete 14
	Write-Log -Level Load -Message "Getting Application log Information"
	#application
	if($computerp  -ne $null){
		$Appdata = FrequencyData -log Application -minimun $MinimunRepetitions
	}
	else{
		$Appdata = FrequencyData -log Application -minimun $MinimunRepetitions -computer $computerp
	}

	Write-Progress -Id 1 -Activity MainScript -PercentComplete 28
	Write-Log -Level Load -Message "Getting Security log Information, this can take a while"
	#security
	if($computerp  -ne $null){
		$SecData = FrequencyData -log Security -minimun $MinimunRepetitions
	}
	else{
		$SecData = FrequencyData -log Security -minimun $MinimunRepetitions -computer $computerp
	}

	#cleanup
	$cleanUpLocal+="SystemData"
	$cleanUpLocal+="Appdata"
	$cleanUpLocal+="SecData"

	Write-Progress -Id 1 -Activity MainScript -PercentComplete 42
	Write-Log -Level Info -Message "Converting Data to JSON"
	#convert to json
	$SystemDataJSON = $SystemData | Sort-Object Count -Descending | ConvertTo-Json
	$AppDataJSON = $Appdata | Sort-Object Count -Descending | ConvertTo-Json
	$SecDataJSON = $SecData | Sort-Object Count -Descending | ConvertTo-Json

	#cleanup
	$cleanUpLocal+="SystemDataJSON"
	$cleanUpLocal+="AppDataJSON"
	$cleanUpLocal+="SecDataJSON"

	Write-Progress -Id 1 -Activity MainScript -PercentComplete 50
	#Export JsonData
	Write-Log -Level Info -Message "Exporting JSON data"
	$SystemData | Sort-Object Count | ConvertTo-Json | Out-File $jsonsystem
	$Appdata  | Sort-Object Count | ConvertTo-Json | Out-File $jsonapplication
	$SecData | Sort-Object Count | ConvertTo-Json | Out-File $jsonsecurity

	Write-Progress -Id 1 -Activity MainScript -PercentComplete 60
	Write-Log -Level Info -Message "Creating graph.js file"

		Try
		{
			CreateJS -SystemData $SystemDataJSON -AppData $AppDataJSON -SecData $SecDataJSON -ErrorAction Stop 
		}
		Catch
		{
			$ErrorMessage = $_.Exception.Message
			$FailedItem = $_.Exception.ItemName
			Write-Output "System Count Objects: $SystemDataJSON.Count, App Count Objects: $AppDataJSON.Count and Security Objects Count:$SecDataJSON.Count, must be all three greathenr than 0 (not empty)"
			Write-Log -Level Error -Message "System Count Objects: $SystemDataJSON.Count, App Count Objects: $AppDataJSON.Count and Security Objects Count:$SecDataJSON.Count, must be all three greathenr than 0 (not empty)"
			$TimeEnd=Get-Date
			showTimeMS $TimeStart $TimeEnd
			Break
		}

	Write-Progress -Id 1 -Activity MainScript -PercentComplete 80
	Write-Log -Level Info -Message "Creating $html file"
	CreateHtml -filename "$html"

	Write-Progress -Id 1 -Activity MainScript -PercentComplete 95
	Write-Log -Level Info -Message "Clean Up"
	
	Write-Log -Level Info "Cleaning up variables"
	$cleanUpLocal | ForEach-Object{
		Remove-Variable $_
	}
	$CleanUpGlobal | ForEach-Object{
		Remove-Variable -Scope global $_
	}


	Write-Progress -Id 1 -Activity MainScript -PercentComplete 100 -Completed
	$TimeEnd=Get-Date
	showTimeMS $TimeStart $TimeEnd

	Write-Log -Level Info -Message "Finished"

	#cleanup
	Remove-Variable -Scope global -Name DefaultLog
	Remove-Variable CleanUpGlobal,cleanUpLocal,TimeEnd,TimeStart