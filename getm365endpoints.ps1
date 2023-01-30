<#

source:       https://github.com/krugerd/getm365endpoints
version:      0.21

#####################################################################################

MIT License

Copyright (c) 2023 David Kruger

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#>

#####################################################################################

#set TenantName
$TenantName = "TenantName"

#set scriptversion
$scriptversion = "0.21"

#set working directory and put us in that directory for all files
$workdir = $env:temp + "\getm365endpoints"

if (!(Test-Path $workdir)) { 
  new-item -path $env:temp -name "getm365endpoints" -itemtype "directory" | out-null
}

set-location $workdir

#log
(get-date -f "yyyy-MM-dd HH:mm:ss") + " - script version $scriptversion - script start" | out-file -append $workdir\log.txt

#####################################################################################

#set required
$required= 'True','False'

#set IPv6
$IPv6 = "false" 

#set lastVersion
$lastVersion = "0000000000" 

#set categories
$categories = 'Optimize', 'Allow', 'Default'

#set serviceAreas
$serviceAreas = 'Exchange','Skype','SharePoint','Common'

#set serviceAreaDisplayNames
$serviceAreaDisplayNames = 'Exchange Online','Skype for Business Online and Microsoft Teams','SharePoint Online and OneDrive for Business','Microsoft 365 Common and Office Online'

#set endpoint types
$endpoints = 'ips','urls'

#set porttypes
$porttypes = 'udpPorts','tcpPorts'

#set bloop
$bloop = "#" * 75

#set msblurb
$msblurb = "
Note:
Optimize endpoints are required for connectivity to every Office 365 service and represent over 75% of Office 365 bandwidth, connections, and volume of data
These endpoints are the most sensitive to network performance, latency, and availability

Microsoft guidance on Optimization Methods:
OPTIMIZE:
Bypass network devices and services that perform traffic interception, SSL decryption, deep packet inspection, and content filtering
Bypass on-premises proxy devices and cloud-based proxy services
Facilitate direct connectivity for VPN users by implementing split tunneling

ALLOW:
Bypass endpoints on network devices and services that perform traffic interception, SSL decryption, deep packet inspection, and content filtering

DEFAULT:
No optimization required - treat like any other internet traffic"

#####################################################################################

Function Set-CIDR2SM {
$_ -replace "/8"," 255.0.0.0"`
   -replace "/9"," 255.128.0.0"`
   -replace "/10"," 255.192.0.0"`
   -replace "/11"," 255.224.0.0"`
   -replace "/12"," 255.240.0.0"`
   -replace "/13"," 255.248.0.0"`
   -replace "/14"," 255.252.0.0"`
   -replace "/15"," 255.254.0.0"`
   -replace "/16"," 255.255.0.0"`
   -replace "/17"," 255.255.128.0"`
   -replace "/18"," 255.255.192.0"`
   -replace "/19"," 255.255.224.0"`
   -replace "/20"," 255.255.240.0"`
   -replace "/21"," 255.255.248.0"`
   -replace "/22"," 255.255.252.0"`
   -replace "/23"," 255.255.254.0"`
   -replace "/24"," 255.255.255.0"`
   -replace "/25"," 255.255.255.128"`
   -replace "/26"," 255.255.255.192"`
   -replace "/27"," 255.255.255.224"`
   -replace "/28"," 255.255.255.240"`
   -replace "/29"," 255.255.255.248"`
   -replace "/30"," 255.255.255.252"`
   -replace "/31"," 255.255.255.254"`
   -replace "/32"," 255.255.255.255"`
}

Function Set-CIDR2WM {
$_ -replace "/8"," 0.255.255.255"`
   -replace "/9"," 0.127.255.255"`
   -replace "/10"," 0.63.255.255"`
   -replace "/11"," 0.31.255.255"`
   -replace "/12"," 0.15.255.255"`
   -replace "/13"," 0.7.255.255"`
   -replace "/14"," 0.3.255.255"`
   -replace "/15"," 0.1.255.255"`
   -replace "/16"," 0.0.255.255"`
   -replace "/17"," 0.0.127.255"`
   -replace "/18"," 0.0.63.255"`
   -replace "/19"," 0.0.31.255"`
   -replace "/20"," 0.0.15.255"`
   -replace "/21"," 0.0.7.255"`
   -replace "/22"," 0.0.3.255"`
   -replace "/23"," 0.0.1.255"`
   -replace "/24"," 0.0.0.255"`
   -replace "/25"," 0.0.0.127"`
   -replace "/26"," 0.0.0.63"`
   -replace "/27"," 0.0.0.31"`
   -replace "/28"," 0.0.0.15"`
   -replace "/29"," 0.0.0.7"`
   -replace "/30"," 0.0.0.3"`
   -replace "/31"," 0.0.0.1"`
   -replace "/32"," 0.0.0.0"`
}

#####################################################################################
# SECTION 1: Check for latest version of IPs and URLs and get if new
#####################################################################################

write-output "file:`t`tgetm365endpoints
ver:`t`t$scriptversion
working dir:`t$workdir"

#generate GUID
$clientRequestId = New-Guid

#GET LATEST VERSION

#check if version.txt exists, if not create one
if (Test-Path -pathtype leaf $workdir\version.txt) {
    $lastVersion = Get-Content $workdir\version.txt
}
else {
    write-output $lastVersion > $workdir\version.txt
}

#check latest version
$version = Invoke-RestMethod -Uri ("https://endpoints.office.com/version/Worldwide?clientRequestId=" + $clientRequestId)

#log
(get-date -f "yyyy-MM-dd HH:mm:ss") + " - local version: $lastversion, latest version: " + $version.latest | out-file -append $workdir\log.txt

#if its newer, do the things
if ($version.latest -gt $lastVersion) {

  Write-Output "`n  Checking for latest version...

  local version:`t$lastVersion
  latest version:`t$($version.latest)"

#  Click OK to get updates..."

 #make version folder
 if (!(Test-Path $workdir\$($version.latest))) { 
  new-item -path $env:temp -name "getm365endpoints\$($version.latest)" -itemtype "directory" | out-null
 } else {

  write-output "Warning: Folder $workdir\$($version.latest) already exists!  Exiting..."
  #log
  (get-date -f "yyyy-MM-dd HH:mm:ss") + " - folder $workdir\$($version.latest) already exists: exiting - script end" | out-file -append $workdir\log.txt
  exit
}

set-location $workdir\$($version.latest)


  #log
  (get-date -f "yyyy-MM-dd HH:mm:ss") + " - updates needed: clicked OK to get updates" | out-file -append $workdir\log.txt

  #GET CHANGES
  Invoke-RestMethod -Uri ("https://endpoints.office.com/changes/Worldwide/" + $lastversion + "?ClientRequestId=" + $clientRequestId + "&format=csv") | set-content changes-all.csv

  #GET ENDPOINTS (full csv)
  Invoke-RestMethod -Uri ("https://endpoints.office.com/endpoints/Worldwide?TenantName=" + $TenantName + "&clientRequestId=" + $clientRequestId + "&NoIPv6=" + $NoIPv6 + "&format=csv") | set-content mscsv-all.csv
  #id,serviceArea,serviceAreaDisplayName,urls,ips,tcpPorts,udpPorts,expressRoute,category,required,notes

  (Get-Content mscsv-all.csv) -replace 'TenantName', '<tenant>' | Set-Content mscsv-all.csv

  #create notes.csv
  import-csv mscsv-all.csv | select ID,ips,urls,tcpports,udpports,required,category,notes | ? notes -ne "" | export-csv Optional-notes.csv -notypeinformation

  #update version.txt
  $version.latest | Out-File $workdir\version.txt

} else { 
  Write-Output "`n  Checking for latest version...

  local version:`t$lastVersion
  latest version:`t$($version.latest)

  You already have the latest version!  Exiting..."
#  Click OK to Exit..." 

  #log
  (get-date -f "yyyy-MM-dd HH:mm:ss") + " - updates not needed: script end" | out-file -append $workdir\log.txt

  exit
}


#####################################################################################
# SECTION 2: Create endpoint-category-required files
#####################################################################################

$bob = import-csv mscsv-all.csv

foreach ($endpoint in $endpoints) {
  foreach ($category in $categories) {
    foreach ($require in $required) {

       if ($require -eq "True") { 
           $req = "Required" } else {
           $req = "Optional" 
         }

          $doug = $bob | ? $endpoint -ne "" | ? category -eq $category | ? required -eq $require

        if ($doug) {	
          $doug.$endpoint -split "," | sort | select-object -unique | add-content $endpoint-$category-$req.txt
        }
      }
    }
  }


#####################################################################################
# SECTION 3: ONEFILE BY ID (mscsv-all.txt)
#####################################################################################

$bob = import-csv mscsv-all.csv
$outfile = "mscsv-all.txt"

  #onefile - byID
  foreach($id in $bob.id){

    $doug = $bob | ? ID -eq $id

    "`n$bloop`nID: $id`nCategory: " + $doug.category +"`nServiceArea: " + $doug.serviceareadisplayname + "`nRequired: " + $doug.required | out-file $outfile -append
  
     if ($doug.notes) {
      "Notes: " + $doug.notes | add-content $outfile
    }

   if ($doug.expressRoute -eq "True") { 
      $ER = "Yes" } else {
      $ER = "No" 
    }
    "ExpressRoute: " + $ER | add-content $outfile
    
   "`r" | add-content $outfile

    if ($doug.tcpPorts) {
      "tcpPorts: " + $doug.tcpPorts | add-content $outfile
    }

    if ($doug.udpPorts) {
      "udpPorts: " + $doug.udpPorts | add-content $outfile
    }
  
    if ($doug.ips) {
      $doug.ips -split "," | sort | select-object -unique | add-content $outfile
    }
  
   if ($doug.urls) {
     $doug.urls -split "," | sort | select-object -unique | add-content $outfile
   }
}


#####################################################################################
# SECTION 4: LISTS IPs and URLs
#####################################################################################


foreach($endpoint in $endpoints) {
  
  #endpoints all - no tmp files
  $doug = $bob | ? $endpoint -ne ""
  $doug.$endpoint -split "," | sort | select-object -unique | set-content all-by$endpoint.txt

  if ($endpoint -eq "ips") {
    #replace CIDR with subnet mask
    Get-Content all-by$endpoint.txt | Foreach-Object { Set-CIDR2SM } | set-content all-by$endpoint-subnetmask.txt
    #replace CIDR with ACL wildcard mask
    Get-Content all-by$endpoint.txt | Foreach-Object { Set-CIDR2WM } | set-content all-by$endpoint-wildcardmask.txt
  }

<#
  #byID
  foreach($id in $bob.id){

    $doug = $bob | ? $endpoint -ne "" | ? ID -eq $id

    if ($doug) {
      "`n$bloop`nID: $id`nCategory: " + $doug.category +"`nServiceArea: " + $doug.serviceareadisplayname + "`nRequired: " + $doug.required + "`nNotes: " + $doug.notes + "`ntcpPorts: " + $doug.tcpPorts + " udpPorts: " + $doug.udpPorts | out-file all-byID.txt -append
      $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byID.txt
    }
  }
#>

  ###byrequired
  foreach($require in $required) {

    $doug = $bob | ? $endpoint -ne "" | ? required -eq $require

    if ($require -eq "True") { 
      $req = "Required" } else {
      $req = "Optional" 
    }
    
    "`n$bloop`n$req" | out-file all-byrequired.txt -append
    $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byrequired.txt
 
    #byrequired,byport
    foreach ($type in $porttypes) {
      $bonkports = $bob.$type -ne "" | select-object -unique | sort

        foreach($port in $bonkports) {
    
          $doug = $bob | ? $endpoint -ne "" | ? required -eq $require | ? $type -eq $port

          if ($doug) {
            "`n$bloop`n$req`n${type}: $port" | out-file all-byrequired-byport.txt -append
            $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byrequired-byport.txt
          }	

	  #byrequired,byport,bycat
          foreach ($category in $categories) {
    
            $doug = $bob | ? $endpoint -ne "" | ? required -eq $require | ? $type -eq $port | ? category -eq $category

            if ($doug) {
              "`n$bloop`n$req ($category)`n${type}: $port" | out-file all-byrequired-byport-bycategory.txt -append
              $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byrequired-byport-bycategory.txt

       #for fw spreadsheet tab2
	"$category`n${type}: $port" | out-file tmp0-$endpoint-$req-byport-bycategory.txt -append
       	$doug.$endpoint -split "," | sort | select-object -unique | add-content tmp0-$endpoint-$req-byport-bycategory.txt

 	      
            }	
          }
        } 	
      } 
    }


  ###byservicearea
  foreach($servicearea in $serviceareadisplaynames){

    $doug = $bob | ? $endpoint -ne "" | ? serviceareadisplayname -eq $servicearea
    "`n$bloop`nService Area: $servicearea" | out-file all-byservicearea.txt -append
    $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byservicearea.txt

    #byservicearea,bycat
    foreach($category in $categories) {

      $doug = $bob | ? $endpoint -ne "" | ? serviceareadisplayname -eq $servicearea | ? category -eq $category

      if ($doug) {
        "`n$bloop`nService Area: $servicearea`nCategory: $category" | out-file all-byservicearea-bycategory.txt -append
        $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byservicearea-bycategory.txt
      }

      #byservicearea,bycat,byrequired
      foreach($require in $required) {

        $doug = $bob | ? $endpoint -ne "" | ? serviceareadisplayname -eq $servicearea | ? category -eq $category | ? required -eq $require

        if ($require -eq "True") {
          $req = "Required" } else {
          $req = "Optional"
        }
     
        if ($doug) {
          "`n$bloop`nService Area: $servicearea`nCategory: $category`nRequired: $req" | out-file all-byservicearea-bycategory-byrequired.txt -append
          $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byservicearea-bycategory-byrequired.txt
        }
       
        #byservicearea,bycat,byrequired,byport
        foreach ($type in $porttypes) {
          $bonkports = $bob.$type -ne "" | select-object -unique | sort

          foreach($port in $bonkports) {
    
            $doug = $bob | ? $endpoint -ne "" | ? serviceareadisplayname -eq $servicearea | ? category -eq $category | ? required -eq $require | ? $type -eq $port

            if ($doug) {
              "`n$bloop`nService Area: $servicearea`nCategory: $category`nRequired: $req`n${type}: $port" | out-file all-byservicearea-bycategory-byrequired-byport.txt -append
              $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byservicearea-bycategory-byrequired-byport.txt
            }	
            	
          }
        }
      }
    }
  }

  ###byport
  foreach ($type in $porttypes) {
    $bonkports = $bob.$type -ne "" | select-object -unique | sort

    foreach($port in $bonkports) {
    
      $doug = $bob | ? $endpoint -ne "" | ? $type -eq $port

        if ($doug) {
          "`n$bloop`n${type}: $port" | out-file all-byport.txt -append
          $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byport.txt
       }

<#
      #by port,byreq
       foreach($require in $required) {

         $doug = $bob | ? $endpoint -ne "" | ? $type -eq $port | ? required -eq $require
           
         if ($require -eq "True") {
             $req = "Required" } else {
             $req = "Optional"
           }
      
        if ($doug) {
           "`n$bloop`n${type}: $port`n($req)" | out-file all-byport-byrequired.txt -append
           $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byport-byrequired.txt
         }
           #new here by port,byreq,bycat
	}
#>


       #by port,bycat
       foreach($category in $categories) {

         $doug = $bob | ? $endpoint -ne "" | ? $type -eq $port | ? category -eq $category
   
         if ($doug) {
           "`n$bloop`n${type}: $port`nCategory: $category" | out-file all-byport-bycategory.txt -append
           $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byport-bycategory.txt
         }

         #by port,bycat,byreq
         foreach($require in $required) {

           $doug = $bob | ? $endpoint -ne "" | ? $type -eq $port | ? category -eq $category | ? required -eq $require
   
           if ($require -eq "True") {
             $req = "Required" } else {
             $req = "Optional"
           }
         
           if ($doug) {
             "`n$bloop`n${type}: $port`nCategory: $category ($req)" | out-file all-byport-bycategory-byrequired.txt -append
             $doug.$endpoint -split "," | sort | select-object -unique | add-content all-byport-bycategory-byrequired.txt
           }
         }
       }
     }
   }

  ###by category
  foreach($category in $categories) {
  
    $doug = $bob | ? $endpoint -ne "" | ? category -eq $category
    
    "`n$bloop`nCategory: $category ($endpoint)" | out-file all-bycategory.txt -append
    $doug.$endpoint -split "," | sort | select-object -unique | add-content all-bycategory.txt

    #by category,byrequired
    foreach($require in $required) {

    $doug = $bob | ? $endpoint -ne "" | ? category -eq $category | ? required -eq $require

      if ($require -eq "True") {
        $req = "Required" } else {
        $req = "Optional"
      }

      if ($doug) {
        "`n$bloop`nCategory: $category ($req)" | out-file all-bycategory-byrequired.txt -append
        $doug.$endpoint -split "," | sort | select-object -unique | add-content all-bycategory-byrequired.txt
        $doug.$endpoint -split "," | sort | select-object -unique | add-content tmp1-$category-$req.txt

      }  
    }

    #bycategory,byport
    foreach ($type in $porttypes) {
      $bonkports = $bob.$type -ne "" | select-object -unique | sort
      
      foreach($port in $bonkports) {

      $doug = $bob | ? $endpoint -ne "" | ? category -eq $category | ? $type -eq $port
   
      if ($doug) {
        "`n$bloop`nCategory: $category (${type}: $port)" | out-file all-bycategory-byport.txt -append
        $doug.$endpoint -split "," | sort | select-object -unique | add-content all-bycategory-byport.txt
      }
    }
  } 

    #by category,byservicearea
    foreach($servicearea in $serviceareadisplaynames) {

      $doug = $bob | ? $endpoint -ne "" | ? category -eq $category | ? serviceareadisplayname -EQ $servicearea
   
        if ($doug) {
          "`n$bloop`nCategory: $category`nService Area: $servicearea" | out-file all-bycategory-byservicearea.txt -append
         $doug.$endpoint -split "," | sort | select-object -unique | add-content all-bycategory-byservicearea.txt
       }

    #by category,byservicearea,byport       
    foreach ($type in $porttypes) {
      $bonkports = $bob.$type -ne "" | select-object -unique | sort
      
      foreach($port in $bonkports) {

        $doug = $bob | ? $endpoint -ne "" | ? category -eq $category | ? serviceareadisplayname -EQ $servicearea | ? $type -eq $port

         if ($doug) {
           "`n$bloop`nCategory: $category`nService Area: $servicearea`n${type}: $port" | out-file all-bycategory-byservicearea-byport.txt -append
           $doug.$endpoint -split "," | sort | select-object -unique | add-content all-bycategory-byservicearea-byport.txt
	 }
      }  
    } 

       #by category,byservicearea,byrequired
       foreach($require in $required) {
         
       if ($require -eq "True") { 
           $req = "Required" } else {
           $req = "Optional" 
         }

       $doug = $bob | ? $endpoint -ne "" | ? category -eq $category | ? serviceareadisplayname -EQ $servicearea | ? required -eq $require
   
        if ($doug) {
          "`n$bloop`nCategory: $category`nService Area: $servicearea ($req)" | out-file all-bycategory-byservicearea-byrequired.txt -append
         $doug.$endpoint -split "," | sort | select-object -unique | add-content all-bycategory-byservicearea-byrequired.txt
       }

        #bycategory,byservicearea,byrequired,byport
        foreach ($type in $porttypes) {
          $bonkports = $bob.$type -ne "" | select-object -unique | sort

          foreach($port in $bonkports) {

            $doug = $bob | ? $endpoint -ne "" | ? category -eq $category | ? serviceareadisplayname -EQ $servicearea | ? required -EQ $require | ? $type -eq $port

            if ($doug) {
               "`n$bloop`nCategory: $category`nService Area: $servicearea ($req)`nPorts: $type $port" | out-file all-bycategory-byservicearea-byrequired-byport.txt -append
               $doug.$endpoint -split "," | sort | select-object -unique | add-content all-bycategory-byservicearea-byrequired-byport.txt
	        }

          }  #end port in bonkports
        }  #end type in porttypes (bycat,byser,byreq,byport)
      }  #end bycat,byservicearea,byreq
    }  #end bycat,byservicearea
  }  #category in categories
}  #endpoint in endpoints


#####################################################################################
# SECTION 5:URLs (SORT, UNIQUE, NO NESTED URLs)
#####################################################################################

$regex = '([A-Z0-9-]+\.[A-Z0-9-]+${1})'  #(AZ09- some sites have -)+ at end of string exactly once

$bang = get-childitem -name all-byurls.txt 

foreach ($b in $bang) {
  select-string $b -Pattern $regex | % { $_.Matches.Value } > tmp0lasttwo.txt 
  get-content tmp0lasttwo.txt | select -unique | sort | foreach {"*." + $_} | set-content all-byurls-notnested.txt
}


#####################################################################################
# SECTION 6:DIFF
#####################################################################################

compare-object (get-content tmp1-Optimize-Required.txt) (get-content tmp1-Allow-Required.txt) -includeequal > diff-OptimizeRequired-AllowRequired.txt
compare-object (get-content tmp1-Allow-Required.txt) (get-content tmp1-Allow-Optional.txt) -includeequal > diff-AllowRequired-AllowOptional.txt
compare-object (get-content tmp1-Default-Required.txt) (get-content tmp1-Default-Optional.txt) -includeequal > diff-DefaultRequired-DefaultOptional.txt

compare-object (get-content tmp1-Optimize-Required.txt) (get-content tmp1-Allow-Required.txt) -includeequal | 
    ForEach-Object {
        if ($_.SideIndicator -eq '=>') {
          write-output $_.InputObject >> tmp1-AllowRequired-NotInOptimizeRequired.txt
        }
       }

compare-object (get-content tmp1-Allow-Required.txt) (get-content tmp1-Allow-Optional.txt) -includeequal | 
    ForEach-Object {
        if ($_.SideIndicator -eq '=>') {
          write-output $_.InputObject >> tmp1-AllowOptional-NotInAllowRequired.txt
        }
       }

compare-object (get-content tmp1-Default-Required.txt) (get-content tmp1-AllowOptional-NotInAllowRequired.txt) -includeequal | 
    ForEach-Object {
        if ($_.SideIndicator -eq '=>') {
          write-output $_.InputObject >> tmp1-AllowOptional-NotInAllowRequiredorDefaultRequired.txt
        }
       }

compare-object (get-content tmp1-Optimize-Required.txt) (get-content tmp1-AllowOptional-NotInAllowRequiredorDefaultRequired.txt) -includeequal | 
    ForEach-Object {
        if ($_.SideIndicator -eq '=>') {
          write-output $_.InputObject >> tmp1-AllowOptional-NotInAllowRequiredorDefaultRequiredorOptimizeRequired.txt
        }
       }

compare-object (get-content tmp1-Default-Required.txt) (get-content tmp1-Default-Optional.txt) -includeequal | 
    ForEach-Object {
        if ($_.SideIndicator -eq '=>') {
          write-output $_.InputObject >> tmp1-DefaultOptional-NotInDefaultRequired.txt
        }
       }

#####################################################################################
# SECTION 7:ONEFILEs
#####################################################################################

###onefile

$now = get-date -f "yyyyMMdd"
$outfile = "m365-ALL-$now.txt"

"m365 List of IPs and URLs

Created on:`t$now
Based on:`tMicrosoft IP and URL Web Service (version $($version.latest))" | out-file $outfile

"`n" + $bloop + $msblurb | add-content $outfile

write-output "`n$bloop`nOptimize (Required)" | add-content $outfile
get-content tmp1-optimize-required.txt | add-content $outfile

write-output "`n$bloop`nAllow (Required - not in Optimize Required already)" | add-content $outfile
get-content tmp1-AllowRequired-NotInOptimizeRequired.txt | add-content $outfile

write-output "`n$bloop`nDefault (Required)" | add-content $outfile
get-content tmp1-Default-Required.txt | add-content $outfile

write-output "`n$bloop`nAllow (Optional - not in Allow/Default/Optimize Required already)`n**FOR YOUR REVIEW - refer to OPTIONAL-NOTES.CSV**" | add-content $outfile
get-content tmp1-AllowOptional-NotInAllowRequiredorDefaultRequiredorOptimizeRequired.txt | add-content $outfile

write-output "`n$bloop`nDefault (Optional - not in Default Required already)`n**FOR YOUR REVIEW - refer to OPTIONAL-NOTES.CSV**" | add-content $outfile
get-content tmp1-DefaultOptional-NotIndefaultRequired.txt | add-content $outfile


###onefile for FW
$outfile = "m365-FW-$now.txt"

"List of IPs and URLs for REQUIRED and OPTIONAL firewall flows (tcp and udp ports)

Created on:`t$now
Based on:`tMicrosoft IP and URL Web Service (version $($version.latest))" | out-file $outfile

"`n" + $bloop + $msblurb | add-content $outfile

get-content all-byrequired-byport-bycategory.txt | add-content $outfile

#####################################################################################
# SECTION 8: PAC file
#####################################################################################

$outfile = "m365-PAC-$now.txt"
$pactxt = "urls-optimize-required.txt","urls-allow-required.txt","urls-default-required.txt"

"m365 sample PAC file

Created on:`t$now
Based on:`tMicrosoft IP and URL Web Service (version $($version.latest))" | out-file $outfile

"`n" + $bloop + $msblurb + "`n`n" + $bloop | add-content $outfile

foreach ($file in $pactxt) {
  write-output "`n${file}:`n" >> $outfile
    foreach ($line in get-content $file) {
      "`t|| shExpMatch(host, `"" + $line + "`")" | add-content $outfile
    }
}

#####################################################################################
# SECTION 9: create excel spreadsheet
#####################################################################################

write-output "`nCreating excel file...`n"

$excel = New-Object -ComObject excel.application
$excel.visible = $True
$workbook = $excel.Workbooks.Add()
#$workbook.Worksheets.Item(3).Delete()

#select and name
$ws = $workbook.Worksheets.Item(1)
$ws.Name = 'all'

Function Set-Categories {
 $ws.Cells.Item(1,2)='OPTIMIZE'
 $ws.Cells.Item(1,2).Font.ColorIndex=2
 $ws.Cells.Item(1,2).Interior.ColorIndex=1
 $ws.Cells.Item(1,2).Font.Size = 14
 $ws.Cells.Item(1,2).Font.Bold=$True
 #$ws.Cells.Item(1,2).Font.Name = "Segoe UI"

 $ws.Cells.Item(1,3)='ALLOW'
 $ws.Cells.Item(1,3).Font.ColorIndex=2
 $ws.Cells.Item(1,3).Interior.ColorIndex=1
 $ws.Cells.Item(1,3).Font.Size = 14
 $ws.Cells.Item(1,3).Font.Bold=$True
 
 $ws.Cells.Item(1,4)='DEFAULT'
 $ws.Cells.Item(1,4).Font.ColorIndex=2
 $ws.Cells.Item(1,4).Interior.ColorIndex=1
 $ws.Cells.Item(1,4).Font.Size = 14
 $ws.Cells.Item(1,4).Font.Bold=$True
}

Set-Categories

###REQUIRED
$row =1
$col=2

$records = get-content tmp1-optimize-required.txt

foreach($record in $records) 
{ 
    $row++
    $ws.cells.item($row,$col)=$record
} 

$lrow = $row

$records = get-content tmp1-AllowRequired-NotInOptimizeRequired.txt

$row =1
$col=3

foreach($record in $records) 
{ 
    $row++
    $ws.cells.item($row,$col)=$record
} 


if ($row -gt $lrow) { $lrow = $row }  

$records = get-content tmp1-Default-Required.txt

$row =1
$col=4

foreach($record in $records) 
{ 
    $row++
    $ws.cells.item($row,$col)=$record
} 

if ($row -gt $lrow) { $lrow = $row } 


###OPTIONAL
#no optimize/optional
    $ws.cells.item($lrow+1,2)="n/a"

$records = get-content tmp1-AllowOptional-NotInAllowRequiredorDefaultRequiredorOptimizeRequired.txt

$row =$lrow
$col=3

foreach($record in $records) 
{ 
    $row++
    $ws.cells.item($row,$col)=$record
} 

$records = get-content tmp1-DefaultOptional-NotIndefaultRequired.txt

$row =$lrow
$col=4

foreach($record in $records) 
{ 
    $row++
    $ws.cells.item($row,$col)=$record
} 

$olrow = $row  #optional endpoints last row

$ws.Cells.Item(2,1)='REQUIRED ENDPOINTS'
$ws.Cells(2, 1).Orientation = 90
#$ws.Cells(2, 1).VerticalAlignment = "xlTop"   #fix this
$ws.Cells(2, 1).Interior.ColorIndex = 4 #green
$ws.Cells.Item(2,1).Font.Size = 14
$ws.Cells.Item(2,1).Font.Bold=$True
$ws.Cells.Item(2,1).Font.Name = "Segoe UI"

$ws.Cells.Item($lrow+1,1)='OPTIONAL ENDPOINTS' 
$ws.Cells($lrow+1, 1).Orientation = 90
#$ws.Cells($lrow+1, 1).VerticalAlignment = "xlTop"   #fix this
$ws.Cells($lrow+1, 1).Interior.ColorIndex = 15 #grey
$ws.Cells.Item($lrow+1,1).Font.Size = 14
$ws.Cells.Item($lrow+1,1).Font.Bold=$True
$ws.Cells.Item($lrow+1,1).Font.Name = "Segoe UI"

#merging
$excel.DisplayAlerts = $false
$MergeCells = $ws.Range("A2:A"+$lrow)  
$MergeCells.Select() | out-null
$MergeCells.MergeCells = $true

$MergeCells = $ws.Range("A" + ($lrow+1) + ":A" + $olrow)
$MergeCells.Select() | out-null
$MergeCells.MergeCells = $true

#freeze top row
$excel.Rows.Item("2:2").Select() | out-null
$excel.ActiveWindow.FreezePanes = $true

#autofit
$ws.UsedRange.EntireColumn.AutoFit() | Out-Null


#### FW TAB

#select and name
$ws = $workbook.worksheets.add()
$ws.Activate()
$ws.Name = 'fw-ports'

Set-Categories

###REQUIRED
$row =2
$col=1

$records = get-content tmp0-ips-Required-byport-bycategory.txt

  foreach($record in $records) { 
    if ($record -eq "Optimize") { $col =2 }
    elseif ($record -eq "Allow") { $col =3 }
    elseif ($record -eq "Default") { $col =4 }
    elseif ($record -match 'pPorts:') { 
      $ws.cells.item($row,2)=$record
      $MergeCells = $ws.Range("B" + $row + ":D" + $row) #fixed
      $MergeCells.Select() | out-null
      $MergeCells.MergeCells = $true
      $ws.Cells.Item($row,2).Interior.ColorIndex=4
      $row++
    } 
    else { 
      $ws.cells.item($row,$col)=$record
      $row++ 
    } 
  } 


$records = get-content tmp0-urls-Required-byport-bycategory.txt

  foreach($record in $records) { 
    if ($record -eq "Optimize") { $col =2 }
    elseif ($record -eq "Allow") { $col =3 }
    elseif ($record -eq "Default") { $col =4 }
    elseif ($record -match 'pPorts:') { 
      $ws.cells.item($row,2)=$record
      $MergeCells = $ws.Range("B" + $row + ":D" + $row) #fixed
      $MergeCells.Select() | out-null 
      $MergeCells.MergeCells = $true
      $ws.Cells.Item($row,2).Interior.ColorIndex=4
      $row++
    } 
    else { 
      $ws.cells.item($row,$col)=$record
      $row++ 
    } 
  } 

$ws.Cells.Item(2,1)='REQUIRED ENDPOINTS'
$ws.Cells(2, 1).Orientation = 90
#$ws.Cells(2, 1).VerticalAlignment = "xlTop"   #fix this
$ws.Cells(2, 1).Interior.ColorIndex = 4 #green
$ws.Cells.Item(2,1).Font.Size = 14
$ws.Cells.Item(2,1).Font.Bold=$True
#$ws.Cells.Item(2,1).Font.Name = "Segoe UI"

$ws.Cells.Item($row,1)='OPTIONAL ENDPOINTS' 
$ws.Cells($row, 1).Orientation = 90
#$ws.Cells($row, 1).VerticalAlignment = "xlTop"   #fix this
$ws.Cells($row, 1).Interior.ColorIndex = 15 #grey
$ws.Cells.Item($row,1).Font.Size = 14
$ws.Cells.Item($row,1).Font.Bold=$True
#$ws.Cells.Item($row,1).Font.Name = "Segoe UI"

$MergeCells = $ws.Range("A2:A"+($row-1)) 
$MergeCells.Select() | out-null
$MergeCells.MergeCells = $true
$lreqrow = $row

###OPTIONAL

$records = get-content tmp0-ips-optional-byport-bycategory.txt

  foreach($record in $records) { 
    if ($record -eq "Optimize") { $col =2 }
    elseif ($record -eq "Allow") { $col =3 }
    elseif ($record -eq "Default") { $col =4 }
    elseif ($record -match 'pPorts:') { 
      $ws.cells.item($row,2)=$record
      $MergeCells = $ws.Range("B" + $row + ":D" + $row) #fixed
      $MergeCells.Select() | out-null
      $MergeCells.MergeCells = $true
      $ws.Cells.Item($row,2).Interior.ColorIndex=15
      $row++
    } 
    else { 
      $ws.cells.item($row,$col)=$record
      $row++ 
    } 
  } 

$records = get-content tmp0-urls-optional-byport-bycategory.txt

  foreach($record in $records) { 
    if ($record -eq "Optimize") { $col =2 }
    elseif ($record -eq "Allow") { $col =3 }
    elseif ($record -eq "Default") { $col =4 }
    elseif ($record -match 'pPorts:') { 
      $ws.cells.item($row,2)=$record
      $MergeCells = $ws.Range("B" + $row + ":D" + $row) #fixed
      $MergeCells.Select() | out-null
      $MergeCells.MergeCells = $true
      $ws.Cells.Item($row,2).Interior.ColorIndex=15
      $row++
    } 
    else { 
      $ws.cells.item($row,$col)=$record
      $row++ 
    } 
  } 

$MergeCells = $ws.Range("A" + $lreqrow + ":A" + ($row-1)) 
$MergeCells.Select() | out-null
$MergeCells.MergeCells = $true

#freeze top row
$excel.Rows.Item("2:2").Select() | out-null
$excel.activewindow.FreezePanes = $true  #fixthis activewindow

#autofit
$ws.UsedRange.EntireColumn.AutoFit() | Out-Null

#gridlines
#$Excel.ActiveWindow.DisplayGridlines = $True

#select worksheet 1
$ws = $workbook.Worksheets.Item(1)
$ws.Activate()
$ws.Cells.Item("1,1").Select() | out-null

#saving & closing the file
$workbook.SaveAs("$workdir\$($version.latest)\m365 Endpoints v$($version.latest).xlsx")
$excel.Quit()


#####################################################################################
# SECTION 10: cleanup
#####################################################################################

write-output "Done!  Check for files in $workdir"

#log
(get-date -f "yyyy-MM-dd HH:mm:ss") + " - files saved to $workdir - script end" | out-file -append $workdir\log.txt

#cleanup
remove-item tmp0*
remove-item tmp1*
remove-item diff*

set-location $workdir

copy *.* $workdir\$($version.latest)

invoke-item $workdir

#end
