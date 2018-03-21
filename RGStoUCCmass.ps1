# RGS to UCC mass export Tool 
# by Martin Wilens
# last modified 2018-03-20
# todo: errorhandling, audiofiles, holidays

<#
.SYNOPSIS 
convert RGS to UCC mass import files.

.DESCRIPTION 
Script to convert Microsoft Skype for Business Response Groups to Anywhere365 UCCs.
The files created can be used with the importtool mass.ps1
See https://www.workstreampeople.com/Anywhere365/golive/Default.htm#Platform_Elements/Management_Tool_Kit/Management_Tool_Kit_PowerMass_Config.html
tested with version: Mass v6.0.17391.4

.PARAMETER OutPath, logPath, domain, SQLserver


.NOTES
This Tools requires you to be CSAdministrator
the Anywhere365 mass.ps1 requires you to be SharePoint Admin
#>


param
(
  $OutPath = "\\ad.utwente.nl\ict\BWS\WilensMHG.pa\Anywhere365\Mass",
  $logPath = "\\ad.utwente.nl\ict\BWS\WilensMHG.pa\Anywhere365\Mass",
  $domain = "@utwente.nl",
  $SQLserver = "sqlprd.ad.utwente.nl;5001;Initial Catalog=AnyWhere365DB;Integrated Security=True"
)

# $ErrorActionPreference = "continue"

'# run this script after all UCCs are created to give then names and phonenumbers'| Set-Content  "$OutPath\settingsscript.ps1"
' Summary created by script for RGS to UCC mass export Tool'| Set-Content  "$OutPath\summary.txt"
$settTxt = (Get-Date).toString('yyyyMMdd-HHmm')
"  -- started at $settTxt --"| Add-Content "$OutPath\summary.txt"


# start with clean files
'UCC;Agent;Order;Formal' |                                                                                     Out-File  "$OutPath\01. Agents.csv"
'UCC;Title;ShowOnWallboard;ForwardToSip;ForwardWithDTMF;StartCountDownSeconds;EndCountDownSeconds;Availability;EscapeSkill;"Content Type"' |Out-File  "$OutPath\02. Skills.csv"
'UCC;Skill;Score;Agent' |                                                                                      Out-File  "$OutPath\03. SkillsPerAgent.csv"
'UCC;Day;Start;End'|                                                                                           Out-File  "$OutPath\04. BusinessHours.csv"
'UCC;Title;Start datetime;End datetime;IVRQuestion'|                                                           Out-File  "$OutPath\05. Holidays.csv"
'UCC;Title;Action;Parent;Question;AudioQuestion;Choice;"Choice Timeout";Answer;AudioAnswer;Skill;Name;Order;Queue;Workflow;"Content Type"' |Out-File  "$OutPath\06. IVRQuestions.csv"
'UCC;Key;Value'|                                                                                               Out-File  "$OutPath\07. Settings.csv"
'UCC;Title;Day;"Specific days";Start time;"Specific start time";List;Item;Column;Value;Active;"Run now once"'| Out-File  "$OutPath\08. TimerJobs.csv"
'UCC;Welcome;ValueStart;ValueEnd;Goodbye;WelcomeAudio;GoodbyeAudio;Modality;Order'|                            Out-File  "$OutPath\09. QualityMonitorConfig.csv"
'UCC;Title;Modality;Skill;ParentQuestion;Priority;EnableRouting;AlwaysOn;"Content Type"'|                     Out-File  "$OutPath\11. Endpoints.csv"

$audiofiles =""
$Agentscsv        = @()
$Skillscsv        = @()
$SkillPerAgentscsv= @()
$HolidaysCsv      = @()
$businessHoursCsv = @()
$IVRQuestionsCsv  = @()
$SettingsCsv      = @()
$HolidaysCsv      = @()
$Endpointscsv     = @()

Function AddAgent{
param(
        [string]$UCCname ,
        [string]$Agent ,
        [string]$Order ,
        [string]$Formal 
     )
  $item = New-Object PSObject
  $item | Add-Member -type NoteProperty -Name 'UCC'    -Value "$UCCname"
  $item | Add-Member -type NoteProperty -Name 'Agent'  -Value "$Agent"
  $item | Add-Member -type NoteProperty -Name 'Order'  -Value "$Order"
  $item | Add-Member -type NoteProperty -Name 'Formal' -Value "$Formal"
  
  return, $item
}

Function AddSkillperAgent{
param(
        [string]$UCC ,
        [string]$Skill ,
        [string]$Score ,
        [string]$Agent 
     )
  $item = New-Object PSObject
  $item | Add-Member -type NoteProperty -Name 'UCC'   -Value "$UCCname"
  $item | Add-Member -type NoteProperty -Name 'Skill' -Value "$Skill"
  $item | Add-Member -type NoteProperty -Name 'Score' -Value "$Score"
  $item | Add-Member -type NoteProperty -Name 'Agent' -Value "$Agent"
  
  return, $item
}

Function AddSkill{
param(
        [string]$UCCname ,
        [string]$SkillTitle ,
        [string]$ShowOnWallboard ,
        [string]$ForwardToSip ,
        [string]$ForwardWithDTMF ,
        [string]$StartCountDownSeconds ,
        [string]$EndCountDownSeconds ,
        [string]$Availability ,
        [string]$EscapeSkill ,
        [string]$ContentType 
     )
  $item = New-Object PSObject
  $item | Add-Member -type NoteProperty -Name 'UCC'                 -Value "$UCCname"
  $item | Add-Member -type NoteProperty -Name 'Title'               -Value "$SkillTitle"
  $item | Add-Member -type NoteProperty -Name 'ShowOnWallboard'     -Value "$ShowOnWallboard"
  $item | Add-Member -type NoteProperty -Name 'ForwardToSip'        -Value "$ForwardToSip"
  $item | Add-Member -type NoteProperty -Name 'ForwardWithDTMF'     -Value "$ForwardWithDTMF"
  $item | Add-Member -type NoteProperty -Name 'StartCountDownSeconds' -Value $StartCountDownSeconds
  $item | Add-Member -type NoteProperty -Name 'EndCountDownSeconds' -Value "$EndCountDownSeconds"
  $item | Add-Member -type NoteProperty -Name 'Availability'        -Value "$Availability"
  $item | Add-Member -type NoteProperty -Name 'EscapeSkill'         -Value "$EscapeSkill"
  $item | Add-Member -type NoteProperty -Name 'Content Type'        -Value "$ContentType"

  return, $item
}

Function AddIVR{
param(
        [string]$UCCname ,
        [string]$Title ,
        [string]$Action ,
        [string]$Parent ,
        [string]$Question ,
        [string]$AudioQuestion ,
        [string]$Choice ,
        [string]$CTimeout ,
        [string]$Answer ,
        [string]$AudioAnswer , 
        [string]$Skill ,
        [string]$Name ,
        [string]$Order ,
        [string]$Queue ,
        [string]$Workflow ,
        [string]$ContentType 
     )
  $item = New-Object PSObject
  $item | Add-Member -type NoteProperty -Name 'UCC'           -Value "$UCCname"
  $item | Add-Member -type NoteProperty -Name 'Title'         -Value "$Title"
  $item | Add-Member -type NoteProperty -Name 'Action'        -Value "$Action"
  $item | Add-Member -type NoteProperty -Name 'Parent'        -Value "$Parent"
  $item | Add-Member -type NoteProperty -Name 'Question'      -Value "$Question"
  $item | Add-Member -type NoteProperty -Name 'AudioQuestion' -Value "$AudioQuestion"
  $item | Add-Member -type NoteProperty -Name 'Choice'        -Value "$Choice"
  $item | Add-Member -type NoteProperty -Name '"Choice Timeout"' -Value "$CTimeout"
  $item | Add-Member -type NoteProperty -Name 'Answer'        -Value "$Answer"
  $item | Add-Member -type NoteProperty -Name 'AudioAnswer'   -Value "$AudioAnswer"
  $item | Add-Member -type NoteProperty -Name 'Skill'         -Value "$Skill"
  $item | Add-Member -type NoteProperty -Name 'Name'          -Value "$Name"
  $item | Add-Member -type NoteProperty -Name 'Order'         -Value "$Order"
  $item | Add-Member -type NoteProperty -Name 'Queue'         -Value "$Queue"
  $item | Add-Member -type NoteProperty -Name 'Workflow'      -Value "$Workflow"
  $item | Add-Member -type NoteProperty -Name '"Content Type"' -Value "$ContentType"

  return, $item
}

Function AddSettings{
param(
        [string]$UCCname ,
        [string]$key ,
        [string]$Value  
     )
  $item = New-Object PSObject
  $item | Add-Member -type NoteProperty -Name 'UCC'   -Value "$UCCname"
  $item | Add-Member -type NoteProperty -Name 'key'   -Value "$key"
  $item | Add-Member -type NoteProperty -Name 'Value' -Value "$Value"
  
  return, $item
}

Function AddHolidays{
param(
        [string]$UCCname ,
        [string]$Title ,
        [string]$StartDate ,
        [string]$EndDate ,
        [string]$IVRQuestion  
     )
  $item = New-Object PSObject
  $item | Add-Member -type NoteProperty -Name 'UCC'            -Value "$UCCname"
  $item | Add-Member -type NoteProperty -Name 'Title'          -Value "$Title"
  $item | Add-Member -type NoteProperty -Name 'Start datetime' -Value "$StartDate"
  $item | Add-Member -type NoteProperty -Name 'End datetime'   -Value "$EndDate"
  $item | Add-Member -type NoteProperty -Name 'IVRQuestion'    -Value "$IVRQuestion"
  
  return, $item
}

Function AddEndpoint{
param(
        [string]$UCCname ,
        [string]$Title ,
        [string]$Modality ,
        [string]$Skill ,
        [string]$ParentQuestion ,
        [string]$Priority ,
        [string]$EnableRouting ,
        [string]$AlwaysOn ,
        [string]$ContentType  
     )
  $item = New-Object PSObject
  $item | Add-Member -type NoteProperty -Name 'UCC'            -Value "$UCCname"
  $item | Add-Member -type NoteProperty -Name 'Title'          -Value "$Title"
  $item | Add-Member -type NoteProperty -Name 'Modality'       -Value "$Modality"
  $item | Add-Member -type NoteProperty -Name 'Skill'          -Value "$Skill"
  $item | Add-Member -type NoteProperty -Name 'ParentQuestion' -Value "$ParentQuestion"
  $item | Add-Member -type NoteProperty -Name 'Priority'       -Value "$Priority"
  $item | Add-Member -type NoteProperty -Name 'EnableRouting'  -Value "$EnableRouting"
  $item | Add-Member -type NoteProperty -Name 'AlwaysOn'       -Value "$AlwaysOn"
  $item | Add-Member -type NoteProperty -Name 'Content Type'   -Value "$ContentType"
  
  return, $item
}

# $RGSname = "Servicedesk IT"
# $RGSname = (Get-CsRgsWorkflow)[3]
foreach($RGSname in Get-CsRgsWorkflow) { 

    # 00. AudioFiles
    # audiofiles not yet implemented  $RGSname.CustomMusicOnHoldFile.OriginalFileName

    $UCCname    = ("ucc_"+(($RGSname).name)  –replace “ “,”_” ).ToLower()
<#    if  ([bool](($RGSname).DefaultAction.prompt.TextToSpeechPrompt)) { 
        $WelcomePrompt =($RGSname).DefaultAction.prompt.TextToSpeechPrompt
        } 
    elseif ( (($RGSname).DefaultAction.prompt.AudioFilePrompt) -ne ""){
        $PromptAudioFile=($RGSname).DefaultAction.prompt.AudioFilePrompt
        $WelcomePrompt ="warning Audiofile $PromptAudioFile not yet inplemented"
        } else {
        $WelcomePrompt =""
        }
#>
    # I asume  DefaultAction   is always a queue
    $QueueId    = ($RGSname).DefaultAction.QueueID
    $QueueIdSplit=($QueueId.ToString()).split("/")[1]
    $RGQueueId = (Get-CsRgsQueue|where {$_.Identity -like "*$QueueIdSplit"})
	$RGSgroupID = ($RGQueueId).AgentGroupIDList
	$agentgroup = ($RGSgroupID|Get-CsRgsAgentGroup)
    $RGSflow = ($RGSname.Identity |Get-CsRgsWorkflow)

    # 01. Agents.csv               UCC;Agent;Order;Formal
    If ($agentgroup.AgentsByUri.Count -eq 0){     write-host ($RGSname.Name)" has NO members"}
        Else {
        $Form = $AgentGroup.ParticipationPolicy 
        ( $agentgroup.AgentsByUri )|ForEach-Object{
            $Agnt=$_.LocalPath
            $Agentscsv += AddAgent -UCCname "$UCCname" -Agent "$Agnt" -Order "1" -Formal "$Form"
            }
        }


	$SkillTT = $RGQueueId.TimeoutThreshold 	#(CountdownAvailabilitySkill:EndCountDownSeconds)
    if ( $SkillTT-eq $null)  {
        $SkillTT       = ""
        $SkilStart     = ""
        $EscSkill      = ""
        $skill1a       = "AvailabilitySkill"
        $SkillTitle    = "AvSkill"
        $IVRTiTreshAct = ""
        $SkillTU       = ""
        } else {
	    $SkilStart     = "0"
        $EscSkill      = "EscSkill"
        $skill1a       = "CountdownAvailabilitySkill"
        $SkillTitle    = "CdAvSkill"
        $IVRTiTreshAct = $RGQueueId.TimeoutAction.Action 	#(FwQueue, Disconnect, Voicemail, Forward)
        $SkillTU       = $RGQueueId.TimeoutAction.Uri 	#(EscapeSkill:ForwardToSip)
        }
    #row 1: (Countdown)AvailabilitySkill
    $Skillscsv +=   AddSkill -UCCname "$UCCname" -SkillTitle "$SkillTitle" -ShowOnWallboard "TRUE" -ForwardToSip ""              -ForwardWithDTMF ""     -StartCountDownSeconds "$SkilStart" -EndCountDownSeconds "$SkillTT" -Availability "Available" -EscapeSkill "$EscSkill" -ContentType "$skill1a"
# ==========================================================================================
    # nonbusinesshours Skill and IVR
    $IVRNonBusPrompt      = $RGSflow.NonBusinessHoursAction.Prompt
    $IVRNonBusA           = $RGSflow.NonBusinessHoursAction.Action
    $SkillNonBusUri       = $RGSflow.NonBusinessHoursAction.Uri

    if ($IVRNonBusA     -eq "Disconnect") {
        $BusnClosedAction = "Disconnect"
        $NBHtitel         = ""
        }
    elseif (($IVRNonBusA -eq "TransferToPstn") -or ($IVRNonBusA -eq "TransferToUri")) {
        $BusnClosedAction = "Skill"
        $NBHtitel         = "NBHtoSIP_Skill"
        }
    elseif ($IVRNonBusA -eq "TransferToVoicemailUri") {
        $BusnClosedAction = "Skill"
        $SkillNonBusUri   = "$SkillNonBusUri;opaque=app:voicemail"
        $NBHtitel         = "NBHtoVoicemail_Skill"
        }
    else {
        $IVRNonBusA       = ""
        }
    if ([bool]$RGSflow.NonBusinessHoursAction.Prompt.AudioFilePrompt) {
        $IVRNonBusPAudio = $RGSflow.NonBusinessHoursAction.Prompt.AudioFilePrompt.OriginalFileName
        $IVRNonBusPrompt =$IVRNonBusPrompt.TextToSpeechPrompt + "(" + $IVRNonBusPAudio +")"
        $audiofiles += "### copy NonBusinessHours Audiofile """+"$IVRNonBusPAudio"" to ""00. AudioFiles\$UCCname\"""+"`r`n"
        New-Item -ItemType directory -Path "$OutPath\00. AudioFiles\$UCCname\" -Force
        "### check $UCCname $NBHtitel and IVR Message Closed"  | Add-Content "$OutPath\summary.txt"
        }
    if ($BusnClosedAction -eq "Skill") {
		# NBHtoSIP_Skill or NBHtoVoicemail_Skill
        $Skillscsv += AddSkill -UCCname "$UCCname" -SkillTitle "$NBHtitel" -ShowOnWallboard ""     -ForwardToSip "$SkillNonBusUri" -ForwardWithDTMF "TRUE" -StartCountDownSeconds ""           -EndCountDownSeconds ""         -Availability ""          -EscapeSkill ""          -ContentType "ForwardSkill"
	    }
    if ($BusnClosedAction -ne "") {
        $IVRQuestionsCsv += AddIVR -UCCname "$UCCname" -Title "Message Closed"  -Action "$BusnClosedAction" -Parent "" -Question "$IVRNonBusPrompt" -AudioQuestion "" -Choice "1" -CTimeout "" -Answer "" -AudioAnswer "" -Skill "$NBHtitel"   -name "" -Order "" -Queue "" -Workflow "" -ContentType ""
        }
    $IVRNonBusPrompt =$IVRNonBusPrompt -replace "`t|`n|`r",""


#   =========================================================================================
    # Queue Overflow Skill and IVR
    $IVROverflPrompt = "" # RGS overflow has no prompt like "We are sorry. There are too many people waiting at the moment"
	$IVROverfAct          = $RGQueueId.OverflowAction.Action #(FwQueue, Disconnect, Voicemail, Forward)
    $SkillOverUri         = $RGQueueId.OverflowAction.Uri 	        #(IVR Overflow:ForwardToSip)
    $SkillOTr             = $RGQueueId.OverflowThreshold 	#(Settings:OverflowThreshold)  
    if ( $SkillOTr-eq $null)  {
        $SkillOTr         = ""
        $IVROverfAct      = ""
        $IVROverfUri      = ""
        $OverfAction      = ""
        $OverfTitel       = ""
        }
    elseif ($IVROverfAct-eq "TransferToUri") {
        $OverfAction      = "Skill"
        $OverfTitel       = "OverflowToSIP_Skill"
        }
    elseif ($SkillOverAct -eq "TransferToVoicemailUri") {
        $OverfAction      = "Skill"
        $SkillOverfUri    = "$SkillOverUri;opaque=app:voicemail"
        $OverfTitel       = "OverflowToVoicemail_Skill"
        }
    elseif ($SkillOverAct -eq "TransferToQueue") {
          # for name of "queue to transfer to"
          $RGQueueId2     = $RGQueueId.OverflowAction.QueueID
          $QueueIdSplit2  =($RGQueueId2.ToString()).split("/")[1]
          $TransferToName = ( (Get-CsRgsQueue|where {$_.Identity -like "*$QueueIdSplit2"}) ).name
          "$UCCname : TransferToQueue ""$TransferToName"" not yet inplmented"   | Add-Content "$OutPath\summary.txt"
        $OverfAction      = "Skill"
        $SkillOverfUri    = "TransferToQueue"
        $OverfTitel       = "TransferToQueue_Skill"
        }
    elseif ($SkillOverAct -eq "Disconnect") {
        $OverfAction        = "Disconnect"
        $OverfTitel         = ""
        }
    else {
        #  ?
        }
    
    if ($OverfAction -eq "Skill") {
        $Skillscsv += AddSkill -UCCname "$UCCname" -SkillTitle "$OverfTitel" -ShowOnWallboard ""     -ForwardToSip "$SkillOverUri" -ForwardWithDTMF "TRUE" -StartCountDownSeconds ""           -EndCountDownSeconds ""         -Availability ""          -EscapeSkill ""          -ContentType "ForwardSkill"
	    }
    if ($OverfAction -ne "") {
        $IVRQuestionsCsv += AddIVR -UCCname "$UCCname" -Title "Message Overflow"  -Action "$OverfAction" -Parent "" -Question "$IVROverflPrompt" -AudioQuestion "" -Choice "1" -CTimeout "" -Answer "" -AudioAnswer "" -Skill "$OverfTitel"   -name "" -Order "" -Queue "" -Workflow "" -ContentType ""
        }

#   =========================================================================================
    # Holiday Skill and IVR
    $IVRHolyPrompt   = $RGSflow.HolidayAction.prompt
	$IVRHolyAct      = $RGSflow.HolidayAction.Action #(FwQueue, Disconnect, Voicemail, Forward)
    $SkillHolyUri    = $RGSflow.HolidayAction.Uri 	#(ForwardToSip)
    if ( $IVRHolyAct-eq $null)  {
        $IVRHolyAct     = ""
        $IVRHolyPrompt  = ""
        $IVRHolyrfAct   = ""
        $IVRHolyUri     = ""
        $HolyAction     = ""
        $HolyTitel      = ""
        }
    elseif ( (($IVRHolyAct -eq "TransferToPstn") -or ($IVRHolyAct -eq "TransferToUri")) -and [bool]$SkillHolyUri ) {
        $HolyAction     = "Skill"
        $HolyTitel      = "HolyToSIP_Skill"
        }
    elseif ($IVRHolyAct -eq "TransferToVoicemailUri" -and [bool]$SkillHolyUri ) {
        $HolyAction     = "Skill"
        $SkillHolyUri   = "$SkillHolyUri;opaque=app:voicemail"
        $HolyTitel      = "HolylToVoicemail_Skill"
        }
    elseif ($IVRHolyAct -eq "Disconnect"){
        $HolyAction     = "Disconnect"
        $SkillHolyUri   = ""
        $HolyTitel      = ""
        }
    else {
        $IVRHolyAct     = ""
        }
    if ([bool]$IVRHolyPrompt.AudioFilePrompt) {
        $IVRHolyPAudio = $RGSflow.OverflowAction.Prompt.AudioFilePrompt.OriginalFileName
        $IVRHolyPrompt =$IVRHolyPrompt.TextToSpeechPrompt + "(" + $IVRHolyPAudio +")"
        $audiofiles += "### copy Holiday AudioFile ""$IVRHolyPAudio"" to ""00. AudioFiles\$UCCname\"""+"`r`n"
        New-Item -ItemType directory -Path "$OutPath\00. AudioFiles\$UCCname\" -Force
        "### check $UCCname $Holytitel and IVR Message Closed or Holiday"   | Add-Content "$OutPath\summary.txt"
        } else {
        $IVRHolyPrompt =$IVRHolyPrompt.TextToSpeechPrompt
        }
    $IVRHolyPrompt =$IVRHolyPrompt -replace "`t|`n|`r",""
    if ($HolyAction -eq "Skill") {
        $Skillscsv += AddSkill -UCCname "$UCCname" -SkillTitle "$HolyTitel" -ShowOnWallboard ""     -ForwardToSip "$SkillHolyUri" -ForwardWithDTMF "TRUE" -StartCountDownSeconds ""           -EndCountDownSeconds ""         -Availability ""          -EscapeSkill ""          -ContentType "ForwardSkill"
	    }
    if ($HolyAction -ne "") {
        ## todo:  -HolidayQuestion "Holiday" not working?
        $IVRQuestionsCsv += AddIVR -UCCname "$UCCname" -HolidayQuestion "Holiday"  -Action "$HolyAction" -Parent "" -Question "$IVRHolyPrompt" -AudioQuestion "" -Choice "1" -CTimeout "" -Answer "" -AudioAnswer "" -Skill "$HolyTitel"   -name "" -Order "" -Queue "" -Workflow "" -ContentType ""
        }

# ===========================================================================================



    $SetHT   = 5 ## todo: not the same as?  ($AgentGroup).AgentAlertTime	# (Settings:HuntTimeout)
    $SetLang = $RGSflow.Language 	#(Settings:CultureInfo)
    $SetRM   = ($AgentGroup).RoutingMethod	#(Settings:Serial, longest idle, Parallel, Attendant = Settings:HuntingMethod)
    # Attendant = to offer a new call to all agents (parallel), regardless of their current presence. -> Anywhere365 LowestHuntPresence: Away
    $IvAudC  = $RGSflow.CustomMusicOnHoldFile
    If ($RGSname.Anonymous) {$SetOutbAr="$True"} else {$SetOutbAr="$False"}
    

      # 02. Skills.csv               UCC;Title;ShowOnWallboard;ForwardToSip;ForwardWithDTMF;StartCountDownSeconds;EndCountDownSeconds;Availability;EscapeSkill;Content Type
      
# ForwardToSip  sip:$SkillNonBusUri;opaque=app:voicemail 
        # NonBusinessHoursAction.Prompt
     #       }

   #row 2: EscapeSkill  
    If ( [bool]$SkillTU ) {
        $Skillscsv +=   AddSkill -UCCname "$UCCname" -SkillTitle "EscSkill" -ShowOnWallboard "" -ForwardToSip "$SkillTU" -ForwardWithDTMF "TRUE" -StartCountDownSeconds "" -EndCountDownSeconds "" -Availability "" -EscapeSkill "" -ContentType "ForwardSkill"
        }
    
    #03. SkillsPerAgent.csv       UCC;Skill;Score;Agent
    If ($agentgroup.AgentsByUri.Count -ne 0)
        {
        ( $agentgroup.AgentsByUri )|ForEach-Object{
            $Agnt=$_.LocalPath
            $SkillPerAgentscsv     += AddSkillperAgent -UCCname "$UCCname" -Skill "$SkillTitle" -Score "100" -Agent "$Agnt" 
     #       If ($EscSkill  -eq "EscSkill" -and [bool]$EscSkill ) {
     #           $SkillPerAgentscsv += AddSkillperAgent -UCCname "$UCCname" -Skill "$EscSkill"   -score "100" -Agent "$Agnt" 
     #           }
            }
        }




    # 04. BusinessHours.csv        UCC;Day;Start;End
    $busH = $RGSflow.BusinessHoursID | Get-CsRgsHoursOfBusiness
    foreach ($Hr in "Hours1","Hours2") {
        foreach ($day in "Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"){
        if (($bush.($day+$Hr)).OpenTime)    {
            $DayOpen = ($bush.($day+$Hr)).OpenTime.ToString()   
            $DayClose = ($bush.($day+$Hr)).CloseTime.ToString()
            $item = New-Object PSObject
            $item | Add-Member -type NoteProperty -Name 'UCC'   -Value "$UCCname"
            $item | Add-Member -type NoteProperty -Name 'Day'   -Value "$day"
            $item | Add-Member -type NoteProperty -Name 'Start' -Value "$DayOpen"
            $item | Add-Member -type NoteProperty -Name 'End'   -Value "$DayClose"
            $businessHoursCsv += $item
         } 
        }
     }


    # 05. Holidays.csv             UCC;Title;Start datetime;End datetime;IVRQuestion
    <# we don't have  holydays in our RGS, so this convertion is not yet implemented
    $RGSHolSetId = Get-CsRgsHolidaySet | Select-Object Identity -ExpandProperty HolidayList
    ForEach ($HolDay in $RGSHolSetId) {
            $HolDayN = $HolDay.Name.ToString()
            $HolDayS = $HolDay.StartDate.ToString()
            $HolDayE = $HolDay.EndDate.ToString()
            $item = New-Object PSObject
            $item | Add-Member -type NoteProperty -Name 'UCC'       -Value "$UCCname"
            $item | Add-Member -type NoteProperty -Name 'Name'      -Value "$HolDayN"
            $item | Add-Member -type NoteProperty -Name 'StartDate' -Value "$HolDayS"
            $item | Add-Member -type NoteProperty -Name 'EndDate'   -Value "$HolDayE"
            $HolidaysCsv += $item
            }
    #>
	if ($HolyAction -eq "Skill") { 
		$HolIVRQuest = "Holiday"
		} else {
		$HolIVRQuest = "Message Closed"
		}
    $HolidaysCsv += AddHolidays -UCCname "$UCCname" -Title "Kerst 2018"                  -StartDate "12/24/2018 12:00 AM" -EndDate "1/2/2019 12:00 AM" -IVRQuestion "$HolIVRQuest"
    $HolidaysCsv += AddHolidays -UCCname "$UCCname" -Title "Goede VRIJdag 30 maart 2018" -StartDate "3/30/2018 12:00 AM" -EndDate "3/31/2018 12:00 AM" -IVRQuestion "$HolIVRQuest"



    # 06. IVRQuestions.csv         UCC;Title;Action;Parent;Question;AudioQuestion;Choice;"Choice Timeout";Answer;AudioAnswer;Skill;Name;Order;Queue;Workflow;Content Type
    $IVRQuestionsCsv += AddIVR -UCCname "$UCCname" -Title "Welcome Message" -Action "Skill"        -Parent "" -Question "$WelcomePrompt" -AudioQuestion ""       -Choice "1" -CTimeout "" -Answer "" -AudioAnswer "" -Skill "$SkillTitle" -name "" -Order "" -Queue "" -Workflow "" -ContentType ""

 	

	
    # 07. Settings.csv             UCC;Key;Value

    $SettingsCsv +=   AddSettings -UCCname "$UCCname" -key "CDRConnectionString" -Value "Server=$SQLserver"
    $SettingsCsv +=   AddSettings -UCCname "$UCCname" -key "ApplicationId"       -Value "urn:application:$UCCname"
    $SettingsCsv +=   AddSettings -UCCname "$UCCname" -key "CultureInfo"         -Value "$SetLang"
    $SettingsCsv +=   AddSettings -UCCname "$UCCname" -key "HuntTimeout"         -Value "$SetHT"
    $SettingsCsv +=   AddSettings -UCCname "$UCCname" -key "OverflowThreshold"   -Value "$SkillOTr"
	$SettingsCsv +=   AddSettings -UCCname "$UCCname" -key "WriteSummaryToSharepoint"  -Value "yes"
    if ($SetOutbAr -eq "True") {
      $SettingsCsv += AddSettings -UCCname "$UCCname" -key "UseOutboundAudioRecording" -Value "True"
      }
    if ($SetRM -eq "Attendant") {
      $SettingsCsv += AddSettings -UCCname "$UCCname" -key "LowestHuntPresence" -Value "Away"
      $SetRM = "Parallel"
      }
    $SettingsCsv +=   AddSettings -UCCname "$UCCname" -key "HuntingMethod"       -Value "$SetRM"

    # 08. TimerJobs.csv  UCC;Title;Day;"Specific days";Start time;"Specific start time";List;Item;Column;Value;Active;"Run now once"
    ## for example: timerjob for holiday message or welcome text goodmorning, goodafternoon

	
    # 09. QualityMonitorConfig.csv  UCC;Welcome;ValueStart;ValueEnd;Goodbye;WelcomeAudio;GoodbyeAudio;Modality;Order
	
	
	# 10. Supervisorss.csv UCC;Supervisor
	
	
	# 11. Endpoints.csv
	# Modality, Skill1, ParentQuestion, Priority and AlwaysOn: Only used for ModalityEndpoint
	# EnableRouting: Only used for ModalityEndpoint and MainEndpoint
	# Content Type: SystemEndpoint, MainEndpoint, DefaultRoutingEndpoint or ModalityEndpoint

	$endpTitle  = ("sip:$UCCname"+"$domain").ToString()
	$endpTitle1 = "sip:$UCCname"+"001"+"$domain"
	$endpTitle2 = "sip:$UCCname"+"002"+"$domain"
	$endpTitle3 = "sip:$UCCname"+"003"+"$domain"
    $endpTitle3 = "sip:$UCCname"+"003"+"$domain"
    $Endpointscsv +=   AddEndpoint -UCCname "$UCCname" -Title "$endpTitle"  -Modality "" -Skill "" -ParentQuestion "" -Priority "" -EnableRouting "yes" -AlwaysOn ""  -ContentType "MainEndpoint"
    $Endpointscsv +=   AddEndpoint -UCCname "$UCCname" -Title "$endpTitle1" -Modality "" -Skill "" -ParentQuestion "" -Priority "" -EnableRouting ""    -AlwaysOn ""  -ContentType "SystemEndpoint"
    $Endpointscsv +=   AddEndpoint -UCCname "$UCCname" -Title "$endpTitle2" -Modality "" -Skill "" -ParentQuestion "" -Priority "" -EnableRouting ""    -AlwaysOn ""  -ContentType "SystemEndpoint"
    $Endpointscsv +=   AddEndpoint -UCCname "$UCCname" -Title "$endpTitle3" -Modality "" -Skill "" -ParentQuestion "" -Priority "" -EnableRouting ""    -AlwaysOn ""  -ContentType "SystemEndpoint"
    if ($SetOutbAr -eq "$True") { 
      $DRendpoint = "sip:$UCCname"+"_dr"+$domain
      $Endpointscsv += AddEndpoint -UCCname "$UCCname" -Title $DRendpoint   -Modality "" -Skill1 "" -ParentQuestion "" -Priority "" -EnableRouting ""    -AlwaysOn ""  -ContentType "DefaultRoutingEndpoint"
      "# run c:\script\A365\extra endpoint\InstallExtraEndpoint.ps1   for ""$UCCname"" to add """+$DRendpoint+""""| Add-Content "$OutPath\settingsscript.ps1"
      }
    if ( ($agentgroup.AgentsByUri.Count -gt 3) -and ($SetRM -eq "Parallel") ) {
      # add  extra SystemEndpoint
      $endpTitle4 = "sip:$UCCname"+"004"+"$domain"
      $Endpointscsv += AddEndpoint -UCCname "$UCCname" -Title "$endpTitle4" -Modality "" -Skill1 "" -ParentQuestion "" -Priority "" -EnableRouting ""    -AlwaysOn ""  -ContentType "SystemEndpoint"
      "# run c:\script\A365\extra endpoint\InstallExtraEndpoint.ps1   for ""$UCCname"" to add """+$DRendpoint+""""| Add-Content "$OutPath\settingsscript.ps1"
      }


    # Common things
   	$RGSname = $RGSflow.name
	$RGSdispnr = $RGSflow.DisplayNumber
	$RGSlUri = $RGSflow.LineUri
	"Set-CsTrustedApplicationEndpoint -Identity ""sip:$($UCCname)@$domain"" -DisplayName   ""$($RGSname)"""  | Add-Content "$OutPath\settingsscript.ps1"
    "Set-CsTrustedApplicationEndpoint -Identity ""sip:$($UCCname)@$domain"" -DisplayNumber ""$($RGSdispnr)"""| Add-Content "$OutPath\settingsscript.ps1"
    "Set-CsTrustedApplicationEndpoint -Identity ""sip:$($UCCname)@$domain"" -LineURI       ""$($RGSlUri)"""  | Add-Content "$OutPath\settingsscript.ps1"
    
    #for summary
    $newRow = New-Object -Type PSObject -Property @{
  'Name'                   = $RGSname
  'DisplayNumber'          = $RGSdispnr
  'LineUri'                = $RGSlUri
  'NonBusinessHoursAction' = $IVRNonBusA
  'NonBusinessHoursURL'    = $SkillNonBusUri
  'NonBusinessHoursPrompt' = $IVRNonBusPrompt
  'BusinessHoursID'        = $busH
  'HolidayAction'          = $IVRHolyAct
  'HolidayUri'             = $SkillHolyUri
  'HolidayIDList'          = $holL
  'HolidayPrompt'          = $IVRHolyPrompt
  'CustomMusicOnHoldFile'  = $IvAudC
  'Language'               = $SetLang
  'AgentAlertTime'         = $SetHT
  'RoutingMethod'          = $SetRM
  'Agents'                 = $Agentscsv
  'ParticipationPolicy'    = $Form     	# Formal / Informal
  'TimeoutThreshold'       = $SkillTT   	#(CountdownAvailabilitySkill:EndCountDownSeconds)
  'TimeoutActionAction'    = $IVRTiTreshAct	#(FwQueue, Disconnect, Voicemail, Forward)
  'TimeoutActionUri'       = $SkillTU    	#(EscapeSkill:ForwardToSip)
  'OverflowThreshold'      = $SkillOTr  	#(Settings:OverflowThreshold)
  'OverflowActionAction'   = $IVROverfAct   #(FwQueue, Disconnect, Voicemail, Forward)
  'OverflowAction.Uri'     = $SkillOverUri 	#(EscapeSkill:ForwardToSip)
   }
    $newRow | Add-Content "$OutPath\summary.txt" 
	if ( [bool]$IvAudC ) {
		$audiofiles += "### copy CustomMusicOnHoldFile ""$IvAudC"" to ""00. AudioFiles\$UCCname\""`r`n"
        New-Item -ItemType directory -Path "$OutPath\00. AudioFiles\$UCCname\" -Force
		}
  
}


$Agentscsv        | Export-Csv -NoTypeInformation -Delimiter ";" -Append  "$OutPath\01. Agents.csv"
$Skillscsv        | export-csv -NoTypeInformation -Delimiter ";" -Append  "$OutPath\02. Skills.csv"
$SkillPerAgentscsv |Export-Csv -NoTypeInformation -Delimiter ";" -Append  "$OutPath\03. SkillsPerAgent.csv"
$businessHoursCsv | export-csv -NoTypeInformation -Delimiter ";" -Append  "$OutPath\04. BusinessHours.csv"
$HolidaysCsv      | export-csv -NoTypeInformation -Delimiter ";" -Append  "$OutPath\05. Holidays.csv"
$IVRQuestionsCsv  | export-csv -NoTypeInformation -Delimiter ";" -Append  -Force "$OutPath\06. IVRQuestions.csv"
$SettingsCsv      | export-csv -NoTypeInformation -Delimiter ";" -Append  "$OutPath\07. Settings.csv"
$Endpointscsv     | export-csv -NoTypeInformation -Delimiter ";" -Append  -Force "$OutPath\11. Endpoints.csv"
$audiofiles       | Add-Content "$OutPath\summary.txt"
$settTxt = (Get-Date).toString('yyyyMMdd-HHmm')
"# -- generated at $settTxt --"| Add-Content "$OutPath\settingsscript.ps1"
"  -- ended at $settTxt --"    | Add-Content "$OutPath\summary.txt"
