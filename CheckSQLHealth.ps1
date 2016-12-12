# ===================================================================================
# Func: Set-AlternatingRows
# Desc: Simple function to alternate the row colors in an HTML table
# ===================================================================================
Function Set-AlternatingRows {

    [CmdletBinding ()]
       Param(
            [Parameter(Mandatory ,ValueFromPipeline)][string] $Line,
            [Parameter(Mandatory)][string] $CSSEvenClass,
            [Parameter(Mandatory)][string] $CSSOddClass
       )
        Begin 
        {
               $ClassName = $CSSEvenClass
        }
        Process 
        {
            If ($Line.Contains("<tr><td>"))
            {       
                $Line = $Line.Replace("<tr>", "<tr class=""$ClassName"">" )
                If ($CSSAlertClass -ne $null)
                { 
                    $ClassName = $CSSAlertClass
                }
                ElseIf ($ClassName -eq $CSSEvenClass)
                {       
                    $ClassName = $CSSOddClass
                }
                Else
                {       
                    $ClassName = $CSSEvenClass
                };
            };
        
            Return $Line;
        }
}

# ===================================================================================
# Func: Send-SPSLogs
# Desc: Send Email with log file in attachment
# ===================================================================================
Function Send-Mail {
	Param 
	(
		[Parameter(Mandatory=$true,Position=0)]$MailFromAddress,
		[Parameter(Mandatory=$true,Position=1)]$MailAddress,
		[Parameter(Mandatory=$true,Position=3)]$MailSubject,
		[Parameter(Mandatory=$true,Position=4)]$MailBody,
		[Parameter(Mandatory=$false,Position=5)]$MailAttachment
	)
	
	$SMTPServer = "smtprelay" #$inputFile.Configuration.EmailNotification.SMTPServer
    $MailFromAddress = "$env:COMPUTERNAME@email.com"

	#Write-Host -ForegroundColor Green "--------------------------------------------------------------"
	Write-Host -ForegroundColor Green " - Sending Email with Log file to $MailAddress ...$EmailSubject"
	Try
	{
		If ($MailAttachment -ne $null)
		{
			$SendLogs = Send-MailMessage -To $MailAddress -From $MailFromAddress -Subject $MailSubject -Body $MailBody -BodyAsHtml -SmtpServer $SMTPServer -Attachments $MailAttachment -ea stop
		}
		else
		{
			$SendLogs = Send-MailMessage -To $MailAddress -From $MailFromAddress -Subject $MailSubject -Body $MailBody -BodyAsHtml -SmtpServer $SMTPServer -ea stop
		}
		#Write-Host -ForegroundColor Green " - Email sent successfully to $MailAddress"
	}
	Catch 
	{
		Write-Host -ForegroundColor Yellow "   * Exception Message: $($_.Exception.Message)"
	}

}

# ===================================================================================
# Func: RunSQLCommand
# Desc: Run T-SQL query
# ===================================================================================
Function RunSQLCommand([String]$SQLServer, [String]$TSQL, [String]$ReportHeading){
	[String]$HTML = ""

	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "server=$SQLServer ; database=msdb; Integrated Security=true"
	#Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
	#Write-Host $SqlConnection.ConnectionString
	
	try{
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
		$SqlCmd.Connection =  
		$SqlCmd.CommandTimeout = 0
		$SqlCmd.Connection.Open()
		$SqlCmd.CommandText = $TSQL

		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $SqlCmd
		
		$DataSet = New-Object System.Data.DataSet
		$SqlAdapter.Fill($DataSet)
		
        if($DataSet.Tables[0].Rows.Count -ne 0 )
		{
            #write-host -ForegroundColor Yellow "## Format results HTML ## TableHeadding: $ReportHeading"			
            $HTML = "  <H2>$ReportHeading</H2><table>"

            #Loop through results then columns the row
            foreach ($Table in $DataSet.Tables)
            { 
                ## build table headings
                $HTML += "<tr>"
                foreach ($field in $Table.Columns)
                {
                    if ($field.ToString().ToLower() -ne "alert") 
                    {
                        $HTML += "<th>$field</th>"
                    }
                }
                $HTML += "</tr>"
                
                [string]$ClassName = "even"
                ## build results
                foreach ($Row in $Table.Rows)
                {
                    If ($ClassName -eq "even") {$ClassName = "odd"} Else {$ClassName = "even"}
                    
                    $TD = ""
                    foreach ($field in $Table.Columns)
                    {
                        #write-host $field #$row[$field]
                        if (($field.ToString().ToLower() -eq "alert") -and ($row[$field].ToString().ToLower() -eq "alert"))
                        {
                            $ClassName = "alert"
                        }
                        if ($field.ToString().ToLower() -ne "alert") 
                        {
                            $TD += "<td>$($Row[$field].Tostring())</td>"
                        }
                    }
            
                    $HTML += "<tr class=""$ClassName""> $TD</tr>";
                }
            } 
			$HTML += "</table>"
		}
	
		return $HTML

	} catch [system.exception]{
		$sqlEx = "<b><font color=red>Caught exception, possibly can't connect to $svr. </font></b>"
		Write-Host "$svr - $sqlEx"
		Write-Host $_
		return $sqlEx
	} finally {
		$SqlConnection.Close()
	}
	
}

# ===================================================================================
# Func: Get-HostUptime
# Desc: get Server Uptime
# ===================================================================================
Function Get-HostUptime {
	param ([string]$ComputerName)
	$Uptime = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName
	$LastBootUpTime = $Uptime.ConvertToDateTime($Uptime.LastBootUpTime)
	$Time = (Get-Date) - $LastBootUpTime
	Return '{0:00} Days, {1:00} Hours, {2:00} Minutes, {3:00} Seconds' -f $Time.Days, $Time.Hours, $Time.Minutes, $Time.Seconds
}

# ===================================================================================
# Func: Get-DiskInfo
# Desc: Get server Get Drive Space
# ===================================================================================
function Get-DiskInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [string[]]$Server, 
        
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [decimal]$DiskThreshold = 100
    )
    
    BEGIN {}
    PROCESS {
        
		$disks =Get-WMIObject -ComputerName $Server Win32_LogicalDisk | Where-Object {$_.DriveType -eq 3}

		foreach ($disk in $disks ) {
		   if ($disks.count -ge 1) {           
			   $Used= $disk.freespace / $disk.size * 100
			   $result =  @{'Server'=$computer.server;
						  'Server Description'=$computer.description;
						  'Volume'=$disk.VolumeName;
						  'Drive'=$disk.name;
						  'Size (gb)'="{0:n2}" -f ($disk.size / 1gb);
						  'Used (gb)'="{0:n2}" -f (($disk.size - $disk.freespace) / 1gb);
						  'Free (gb)'="{0:n2}" -f ($disk.freespace / 1gb);
						  '% free'="{0:n2}" -f ($disk.freespace / $disk.size * 100)}                         
						 

			   $obj = New-Object -TypeName PSObject -Property $result
			   #if ($Used -lt $Diskthreshold){  
					Write-Output $obj 
			   #}
		   }
		}
        
    }
    END {}
} # end of function

function BuildDriveSpace([String]$Server){
    #define an array for html fragments
    $fragments=@()

    # Set free disk space threshold below in percent (default at 10%)
    [decimal]$thresholdspace = 10

    #this is the graph character
	<#
	    ?  &#9644;
        ¦  &#9617; {dither light}	
        ¦  &#9618; {dither medium}	
        ¦  &#9619; {dither heavy}	
        ¦  &#9608; {full box}
	#>
    [string]$g=[char]9619 
	

    # call the main function
    $Disks = Get-DiskInfo `
                -ErrorAction SilentlyContinue `
                -Server ($Server) `
                -DiskThreshold $thresholdspace

    #create an html fragment
    $html= $Disks|select @{name="Drive";expression={$_.Drive}},
                  @{name="Volume";expression={$_.Volume}},
                  @{name="Size (gb)" ;expression={($_."size (gb)")}},
                  @{name="Used (gb)";expression={$_."used (gb)"}},
                  @{name="Free (gb)";expression ={$_."free (gb)"}},
                  @{name="% free";expression ={$_."% free"}},           
                  @{name="Disk usage";expression={
                        $UsedPer = (($_."Size (gb)" - $_."Free (gb)") / $_."Size (gb)") * 100
                        $UsedGraph = $g * ($UsedPer / 4)
                        $FreeGraph = $g * ((100 - $UsedPer) / 4)
                        #using place holders for the < and > characters
                         "xopenFont color=Redxclose{0}xopen/FontxclosexopenFont Color=Greenxclose{1}xopen/fontxclose" -f $usedGraph, $FreeGraph }}`
        | sort-Object {[string]$_."Drive"} `
        | ConvertTo-HTML -fragment
	
	$html=$html -replace $g, "&#9608;"
	
    #replace the tag place holders.
    $html=$html -replace "xopen","<"
    $html=$html -replace "xclose",">"
    

    #add to fragments
    $Fragments+=$html         

    return $Fragments

}

# ===================================================================================
# Get system info and disk used free space
# ===================================================================================
Function System-And-DiskSpace {
	Param 
	(
		[Parameter(Mandatory=$true,Position=0)]$computer
	)

    # Set free disk space threshold below in percent (default at 10%)
    $thresholdspace = 100
    [int]$EventNum = 3

    $ListOfAttachments = @()
    $Report = @()
    $CurrentTime = Get-Date

	# Build disk report
    $DiskInfo = BuildDriveSpace -Server $computer

#region System Info
	$OS = (Get-WmiObject Win32_OperatingSystem -computername $computer).caption
	$SystemInfo = Get-WmiObject -Class Win32_OperatingSystem -computername $computer | Select-Object Name, TotalVisibleMemorySize, FreePhysicalMemory
	$TotalRAM = $SystemInfo.TotalVisibleMemorySize/1MB
	$FreeRAM = $SystemInfo.FreePhysicalMemory/1MB
	$UsedRAM = $TotalRAM - $FreeRAM
	$RAMPercentFree = ($FreeRAM / $TotalRAM) * 100
	$TotalRAM = [Math]::Round($TotalRAM, 2)
	$FreeRAM = [Math]::Round($FreeRAM, 2)
	$UsedRAM = [Math]::Round($UsedRAM, 2)
	$RAMPercentFree = [Math]::Round($RAMPercentFree, 2)

#region Uptime

# Fetch the Uptime of the current system using our Get-HostUptime Function.
	$SystemUptime = Get-HostUptime -ComputerName $computer
#endregion


    [String]$CurrentSystemHTML = @"
	    <table class="list" border=1>
            <tr><th colspan=2>System Info: $computer </th></tr>
            <tr>
	            <td>System Uptime</td>
	            <td>$SystemUptime</td>
	        </tr>
	        <tr>
	            <td>OS</td>
	            <td>$OS</td>
	        </tr>
	        <tr>
	            <td>Total RAM (GB)</td>
	            <td>$TotalRAM</td>
	        </tr>
	        <tr>
	            <td>Free RAM (GB)</td>
	            <td>$FreeRAM</td>
	        </tr>
	        <tr>
	            <td>Percent free RAM</td>
	            <td>$RAMPercentFree</td>
	        </tr>
	    </table>


	    <font size=3><b>Disk Info</b></font>
        $DiskInfo
"@

	If ($ServicesReport -ne $null){
		$CurrentSystemHTML += @"    
			<font size=3><b>System Services - Automatic Startup but not Running</b></font><br>
			<i>The following services are those which are set to Automatic startup type, yet are currently not running on $computer</i>
			$ServicesReport
"@
	}
	If ($SystemEventsReport -ne $null){
		$CurrentSystemHTML += @" 
			<font size=3><b>Events Report - The last $EventNum System/Application Log Events that were Warnings or Errors</b></font><br>
			<i>The following is a list of the last $EventNum <b>System log</b> events that had an Event Type of either Warning or Error on $computer</i>
			$SystemEventsReport
"@
	}
	If ($ApplicationEventsReport -ne $null){
	$CurrentSystemHTML += 
@" 
			<i>The following is a list of the last $EventNum <b>Application log</b> events that had an Event Type of either Warning or Error on $computer</i><br>
			$ApplicationEventsReport
"@
	}
	    
    return $CurrentSystemHTML
}

# ======================================================================================================================================================================
# Public var
# ======================================================================================================================================================================
#region Variables and Arguments
try{

[string]$Results = "";
[string]$Heading = "";
[string]$sql = "";
[string]$MailHeader = @"
<style type="text/css">
	body {font-family: Arial, Helvetica, sans-serif;}
	TABLE {border-width: 0px; border-style: solid; border-color: black; border-collapse: collapse;}
	TH {border-width: 1px; border-style: solid; border-color: #000000; background-color: #778899; color: #FFFFFF; text-align:left;}
	TD {font-size:12px; border-width: 1px; padding: 8px;border-style: solid; border-color: black;}
	.odd  { background-color:#F0F0F0; }
	.even { background-color:#dddddd; }
	.alert { background-color:#FFCCCC; }
	h2{ clear: both; font-size: 110%; }
</style>
"@;
[string]$MissingBackups = $null;
[String]$FailedAgentJobResults = $null;
[string]$SQLErrorsResults = $null;
[string]$emailTo = "${MailReportsTo}";

[xml]$XmlDocument = @"
<?xml version="1.0" encoding="utf-8"?>
<Checks>
	<Server Name="PRD3002SUPSQL">
		<CheckSQLErrors>true</CheckSQLErrors>
		<CheckAgentJob>false</CheckAgentJob>
		<CheckMissingBackups>false</CheckMissingBackups>
	</Server>
 </Checks>
"@;
}
Catch 
{
	Write-Host -ForegroundColor Yellow "   * Exception Message: $($_.Exception.Message)"
    Exit
}
#endregion

# ===================================================================================
# Run through servers and check them
# ===================================================================================
foreach($ServerName in $XmlDocument.Checks.Server){
	[string]$svr = $ServerName.Name

    # ===================================================================================
    # Get system specs
    # ===================================================================================
    [String]$SystemInfo = System-And-DiskSpace $svr

    # ===================================================================================
    # SQL Errors
    # ===================================================================================
    [String]$Heading = "SQL Server Errors for: $svr"
    #Write-Host "1) $Heading" -BackgroundColor Blue -ForegroundColor White
	
	If ($ServerName.CheckSQLErrors -eq $true)
	{
		[string]$sql = @"
            DECLARE @Time_Start DATETIME;
            DECLARE @Time_End DATETIME;
            SET @Time_Start = GETDATE() - 2;
            SET @Time_End = GETDATE();

            DECLARE @ErrorLog TABLE
	            (logdate DATETIME
	            ,processinfo VARCHAR(255)
	            ,Message VARCHAR(500));

            INSERT @ErrorLog (logdate,processinfo,Message)
		            EXEC master.dbo.xp_readerrorlog 0, 1, NULL, NULL, @Time_Start, @Time_End, N'desc';

            SELECT logdate
	               ,Message
	               ,CASE WHEN Message LIKE '%without errors%' THEN 'normal'
			             ELSE 'alert'
		            END AS Alert
	            FROM
		            @ErrorLog
	            WHERE
		            (
		             (
		              Message LIKE '%error%' AND (Message NOT LIKE '%found 0 errors%' and Message NOT LIKE '%without errors%')
		             ) OR
		             Message LIKE '%failed%'
		            ) AND
		            processinfo NOT LIKE 'logon'
	            ORDER BY
		            logdate DESC;
"@;

		[string]$SQLErrorsResults = RunSQLCommand -SQLServer $svr -TSQL $sql -ReportHeading $Heading
    }
    # ===================================================================================
    # Failed Agent Job
    # ===================================================================================
    [String]$Heading = "Agent Job - Failed to complete for: $svr"
    #Write-Host "2) $Heading" -BackgroundColor Blue -ForegroundColor White
	If ($ServerName.CheckAgentJob -eq $true)
	{
		[string]$sql = @"
		    SELECT
			    JobStatus = 'FAILED'
		       ,JobName = CAST(sj.name AS VARCHAR(100))
		       ,StepID = CAST(sjs.step_id AS VARCHAR(5))
		       ,StepName = CAST(sjs.step_name AS VARCHAR(30))
		       ,StartDateTIME = CAST(REPLACE(CONVERT(VARCHAR, CONVERT(DATETIME, CONVERT(VARCHAR, sjh.run_date)), 102), '.', '-') + ' ' +
			    SUBSTRING(RIGHT('000000' + CONVERT(VARCHAR, sjh.run_time), 6), 1, 2) + ':' + SUBSTRING(RIGHT('000000' +
																										    CONVERT(VARCHAR, sjh.run_time),6), 3, 2) + ':' +
			    SUBSTRING(RIGHT('000000' + CONVERT(VARCHAR, sjh.run_time), 6), 5, 2) AS VARCHAR(30))
		       ,[Message] = sjh.message
			    ,Alert = 'alert'

		    FROM
			    msdb.dbo.sysjobs sj
			    JOIN msdb.dbo.sysjobsteps sjs
				    ON sj.job_id = sjs.job_id
			    JOIN msdb.dbo.sysjobhistory sjh
				    ON sj.job_id = sjh.job_id AND
				       sjs.step_id = sjh.step_id
		    WHERE
			    sjh.run_status IN (0, 3) AND
			    CAST(sjh.run_date AS FLOAT) * 1000000 + sjh.run_time > CAST(CONVERT(VARCHAR(8), GETDATE() - 1, 112) AS FLOAT) *
			    1000000 + 70000 --yesterday at 7am
	    UNION
	    SELECT
			    JobStatus = 'FAILED'
		       ,JobName = CAST(sj.name AS VARCHAR(100))
		       ,StepID = 'MAIN'  
		       ,StepName = 'MAIN' 
		       ,StartDateTIME = CAST(REPLACE(CONVERT(VARCHAR, CONVERT(DATETIME, CONVERT(VARCHAR, sjh.run_date)), 102), '.', '-') + ' ' +
			    SUBSTRING(RIGHT('000000' + CONVERT(VARCHAR, sjh.run_time), 6), 1, 2) + ':' + SUBSTRING(RIGHT('000000' +
																										    CONVERT(VARCHAR, sjh.run_time),6), 3, 2) + ':' +
			    SUBSTRING(RIGHT('000000' + CONVERT(VARCHAR, sjh.run_time), 6), 5, 2) AS VARCHAR(30))
		       ,[Message] = sjh.message
                ,Alert = 'alert'
		    FROM
			    msdb.dbo.sysjobs sj
			    JOIN msdb.dbo.sysjobhistory sjh
				    ON sj.job_id = sjh.job_id
		    WHERE
			    sjh.run_status IN (0, 3) AND
			    sjh.step_id = 0 AND
			    CAST(sjh.run_date AS FLOAT) * 1000000 + sjh.run_time > CAST(CONVERT(VARCHAR(8), GETDATE() - 1, 112) AS FLOAT) *
			    1000000 + 70000;
"@;
    
		[String]$FailedAgentJobResults = RunSQLCommand -SQLServer $svr -TSQL $sql -ReportHeading $Heading
    }
    # ===================================================================================
    # Check Missing Backups
    # ===================================================================================
    [String]$Heading = "Check Missing Backups: $svr"
    #Write-Host "3) $Heading" -BackgroundColor Blue -ForegroundColor White
	If ($ServerName.CheckMissingBackups -eq $true)
	{
		[string]$sql = @"
	    SELECT
		    DatabaseName = d.name
		    ,BackupType = 'Full'
		    ,LastFullBackup = ISNULL(CONVERT(VARCHAR, b.backupdate, 120), 'NEVER')
		    ,Alert = 'alert'
	    FROM
		    sys.databases d
		    LEFT JOIN (
					    SELECT
						    database_name
						    ,[Type]
						    ,MAX(backup_finish_date) backupdate
					    FROM
						    msdb.dbo.backupset
					    WHERE
						    [Type] LIKE 'D'
					    GROUP BY
						    database_name
						    ,[Type]
					    ) b
			    ON d.name = b.database_name
	    WHERE
		    (
				    backupdate IS NULL OR
				    backupdate < GETDATE() - 1
		    ) AND
		    d.name <> 'tempdb'
	    UNION
		    SELECT
			    DatabaseName = d.name
			    ,BackupType = 'Trn'
			    ,LastLogBackup = ISNULL(CONVERT(VARCHAR, b.backupdate, 120), 'NEVER')
			    ,Alert = 'alert'
		    FROM
			    sys.databases d
			    LEFT JOIN (
					        SELECT
							    database_name
						        ,[Type]
						        ,MAX(backup_finish_date) backupdate
						    FROM
							    msdb.dbo.backupset
						    WHERE
							    Type LIKE 'L'
						    GROUP BY
							    database_name
						        ,[Type]
					        ) b
				    ON d.name = b.database_name
		    WHERE
			    recovery_model = 1 AND
			    (
				        backupdate IS NULL OR
				        backupdate < GETDATE() - 1
			    ) AND
			    d.name <> 'tempdb';
"@;

		[string]$MissingBackups = RunSQLCommand -SQLServer $svr -TSQL $sql -ReportHeading $Heading
	}

    if ($SQLErrorsResults.Length -lt 4) { $SQLErrorsResults = ""} else {$SQLErrorsResults = $SQLErrorsResults.Substring(3)}

	if ($FailedAgentJobResults.Length -lt 4) { $FailedAgentJobResults = ""} else {$FailedAgentJobResults = $FailedAgentJobResults.Substring(3)}
	if ($MissingBackups.Length -lt 4) { $MissingBackups = ""} else {$MissingBackups = $MissingBackups.Substring(3)}
	
    [String]$MailBody = "$MailHeader $SystemInfo $SQLErrorsResults $FailedAgentJobResults $MissingBackups</body>"
    
    # ===================================================================================
    # Send mail
    # ===================================================================================
	Send-Mail -MailFromAddress "$svr@email.com" -MailAddress $emailTo -MailSubject "Server Check $svr" -MailBody $MailBody
    
}