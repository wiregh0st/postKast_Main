
Write-Host "#####################################################################################
#                                                                                   #
#                Welcome to postKast developed by wiregh0st                         #
#                                                                                   #
#                                                                                   #        
#            You must change the variables at the top of PostKast.ps1               #
#            script if the file locations change at any time. If you do             #
#            not change the variables, the script will not execute properly!        #
#                                                                                   #
#####################################################################################"

Write-Host "`n"
Write-Host "Sleeping 120s before starting..."
Start-Sleep -s 10

$namespaceName = "root\cimv2\mdm\dmmap"
$className = "MDM_EnterpriseModernAppManagement_AppManagement01"
$wmiObj = Get-WmiObject -Namespace $namespaceName -Class $className
$result = $wmiObj.UpdateScanMethod()
write-host "here"