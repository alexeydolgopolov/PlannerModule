param ( 
    [switch]$Verbose, 
    $sPlan,
    $dPlan,
    [switch]$CopyComments
)



function Get-PlanGroup {
    param ( $planId )
    $sGroupId = $null
    Get-UnifiedGroupsList | % { 
        $id = $_.id
        Get-PlannerPlansList -GroupID $_.id | % {        
            if ( $_.id -eq $planId ) { 
                $sGroupId = $id
            }
        }    
    }
    return $sGroupId 
}

function Copy-Comments { 
    param ( $taskName, $sTkaskId, $dTaskId, $AuthToken, $sPlanId, $dPlanId, $sGroupId, $dGroupId )

    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$sGroupId/threads" -Headers $AuthToken -Method GET -ContentType 'application/json; charset=utf-8' | select -ExpandProperty value | ? topic -match "Comments on task " | % {
        $thread = $_
        $posts = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$sGroupId/threads/$($_.id)/posts" -Headers $AuthToken -Method GET -ContentType 'application/json; charset=utf-8' | select -ExpandProperty value
        if ( $thread.topic -match $taskName) {
            $newThread = @{
                topic = ('Comments on task "{0}"' -f $taskName)
                posts = @()
            }
            $posts[0] | % {  $newThread.posts += @{ body = $_.body}}
            $newThreadJson = $newThread | ConvertTo-Json -Depth 100 
            $newThreadJson = $newThreadJson -replace $sPlanId, $dPlanId
            $newThreadJson = $newThreadJson -replace $sTkaskId, $dTaskId
            $newT = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$dGroupId/threads" -Headers $AuthToken -Method POST -Body $newThreadJson -ContentType 'application/json; charset=utf-8'
            $posts[1..$posts.count]| % {
                $newPost = @{
                    post = @{ 
                        body = $_.body
                    }
                }
                $newPostJson = $newPost | ConvertTo-Json  -Depth 100
                $newPostJson = $newPostJson -replace $sPlanId, $dPlanId
                $newPostJson = $newPostJson -replace $sTkaskId, $dTaskId
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$dGroupId/threads/$($newT.id)/reply" -Headers $AuthToken -Method POST -Body $newPostJson  -ContentType 'application/json; charset=utf-8'
            }
        }
        
    }
}

if ( $Verbose ) {
    $VerbosePreference = "continue"
}

# check module if not installed
$AadModule = Get-Module -Name "AzureAD" -ListAvailable | Sort-Object -Property Version -Descending | Select-Object -First 1
if ($AadModule -eq $null) {
    Install-Module -Name "AzureAD" -Scope CurrentUser -Force -AllowClobber
}

pushd $PSScriptRoot
Import-Module .\PlannerModule\PlannerModule.psm1 -Force
popd
Update-PlannerModuleEnvironment -ClientId "67f1395a-801f-4f6f-afeb-a5305bcb578a" -redirectUri "https://login.microsoftonline.com/common/oauth2/nativeclient" -silent
$token = Connect-Planner  -ReturnToken # ForceInteractive


if ( $sPlan -eq $null ) {
    $sGroupId = Get-UnifiedGroupsList | sort -Property displayName | select displayName,id | Out-GridView -OutputMode Single -Title "Select SOURCE Group" 
    $sGroupId = $sGroupId.id
    $sPlan = Get-PlannerPlansList -GroupID $sGroupId | sort -Property title | select title,id | Out-GridView -OutputMode Single -Title "Select SOURCE Plan" 
    $sPlan  = $sPlan.id
} else {
    $sGroupId = Get-PlanGroup -planId $sPlan
}

if ( $dPlan -eq $null ) {
    $dGroupId = Get-UnifiedGroupsList | sort -Property displayName | select displayName,id | Out-GridView -OutputMode Single -Title "Select Destination Group" 
    $dGroupId = $dGroupId.id
    $dPlan = Get-PlannerPlansList -GroupID $dGroupId | sort -Property title | select title,id | Out-GridView -OutputMode Single -Title "Select Destination Plan" 
    $dPlan  = $dPlan.id
} else {
    $dGroupId = Get-PlanGroup -planId $dPlan
}

if ( $sPlan -eq $null -or  $dPlan -eq $null ) {
    Write-Host "source and destination plan should be provided" -ForegroundColor Red
    exit 1
}

$sBuckets = Get-PlannerPlanBuckets -PlanID $sPlan
$dBuckets = Get-PlannerPlanBuckets -PlanID $dPlan
$sTasks = Get-PlannerPlanTasks -PlanID $sPlan | ? percentComplete -ne 100 
$dTasks = Get-PlannerPlanTasks -PlanID $dPlan

$ChangeList = @()

$sBuckets | % {
    $sBucketName = $_.name
    $dBucket = $dBuckets | ? name -eq $sBucketName
    if ( $dBucket -eq $null ) {
        Write-Verbose "$sBucketName - will be created"
        $ChangeList += [PSCustomObject]@{ 
            action = "create bucket"
            name = $sBucketName
            sId = $_.id
            sObj = $_
        } 
    }
}



$sTasks | % {
    $sTaskName = $_.title
    $dTask = $dTasks | ? title -eq $sTaskName
    if ( $dTask -eq $null ) {
        Write-Verbose "$sBucketName - will be created"
        $ChangeList += [PSCustomObject]@{ 
            action = "create task"
            name = $sTaskName
            sId = $_.id
            details = Get-PlannerTaskDetails -TaskID $_.id
            sObj = $_
        } 
    }
}

$bucketToCreate = $ChangeList | ? action -eq "create bucket"
$tasksToCreate = $ChangeList | ? action -eq "create task"
$dPlanName = Get-PlannerPlan -PlanID $dPlan | select -ExpandProperty title

if ( $bucketToCreate.count -gt 0 ) {
    Write-Host "Following buckets to be created in the plan $dPlanName ($dPlan)"
    $bucketToCreate | % {
        Write-Host "    $($_.name)"
    }
    $yesNo = Read-Host -Prompt "continue with the creation? [y/N]"
    if ( $yesNo -ne "y" ) { exit 0 }
    $bucketToCreate | % {
        New-PlannerBucket -PlanID $dPlan -BucketName $_.name | Out-Null
    }
}


if ( $tasksToCreate.count -gt 0 ) {
    Write-Host "Following tasks to be created in the plan $dPlanName ($dPlan)"
    $tasksToCreate | % {
        Write-Host "    $($_.name)"
    }
    $yesNo = Read-Host -Prompt "continue with the creation? [y/N]"
    if ( $yesNo -ne "y" ) { exit 0 }
    $dBuckets = Get-PlannerPlanBuckets -PlanID $dPlan
    $tasksToCreate | % {
        $sBucketID = $_.sObj.bucketID
        $sBucketIDname = $sBuckets | ? id -eq $sBucketID | select -ExpandProperty name
        $BucketID = $dBuckets | ? name -eq $sBucketIDname | select -ExpandProperty id
        $taskObject = @{
            PlanID = $dPlan
            TaskName = $_.name
            BucketID = $BucketID
        }
        $taskDescription = @()
        $users = @()
        if ( $_.sObj.startDateTime ) { $taskObject.startDateTime = $_.sObj.startDateTime }
        if ( $_.sObj.dueDateTime ) { $taskObject.dueDateTime = $_.sObj.dueDateTime }
        
        if ( $_.details.description ) { $taskDescription += $_.details.description }
        if ( $_.sObj.assignments.PSObject.Properties.name.count -gt 0 ) { 
            $_.sObj.assignments.PSObject.Properties.name | % {
                try { 
                   $assigneeDetails =  Get-AADUserDetails $_
                   $taskDescription += "Assigned to: $($assigneeDetails.userPrincipalName)"
                } catch {
                    $taskDescription += "Assigned to: $($_)"                    
                }
            } 
        }
        if ( $_.sObj.percentComplete ) { $taskObject.percentComplete = $_.sObj.percentComplete }
        
        Write-host "creating task $($taskObject.TaskName)" -ForegroundColor Yellow
        $newTask = New-PlannerTask @taskObject

        if ( $taskDescription.Count -gt 0 ) { 
            Write-host "adding description"
            Add-PlannerTaskDescription -TaskID $newTask.id  -Description ( $taskDescription -Join "`r`n" )
        }
 
        $taskChecklist = $_.details.checklist
        if ( $taskDescription ) {
            $taskChecklistArray = @()
            $taskChecklist.PSObject.Properties.name | % { $taskChecklistArray += $taskChecklist.$_ }
            Write-host "adding task checklist"
            $taskChecklistArray | sort -property orderHint -Descending | % {
                Add-PlannerTaskChecklist -TaskID $newTask.id -Title $_.title -IsChecked $_.isChecked
            }
        }
        if ( $CopyComments ) {
            Copy-Comments -taskName $taskObject.TaskName -sTkaskId $_.sObj.id -dTaskId $newTask.id -AuthToken $token -sPlanId $sPlan -dPlanId $dPlan -sGroupId $sGroupId -dGroupId $dGroupId
        }
    }
}
<#
Links 
    Planner API https://docs.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-1.0
    Ref module atricle https://www.scconfigmgr.com/2019/06/06/powershell-module-for-microsoft-planner/
    Some link for C# https://laurakokkarinen.com/how-to-sort-tasks-using-planner-order-hint-and-microsoft-graph/
#>