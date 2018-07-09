<#
    .SYNOPSIS
    Get-ExchangePQT.ps1
    Exchange-Processor Query Tool

    Michel de Rooij
    michel@eightwone.com
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.2, July 9th, 2018
    
    .DESCRIPTION
    The script is an alternative to the Processor Query Tool and automates the steps 
    to calculate the SPECint 2006 Rate Value (www.spec.org) of your planned
    processor model when Exchange Server 2010/2013 configurations. For virtualized
    environments, you can determine the SPECint for a specific virtual processor 
    ratio in combination with allocated number of vCPUs.
    
    Using PowerShell, you can also perform other tasks, such as:
    - Calculate the average SPECint2006 Rate Value for a certain CPU/cores configuration.
    - Query which systems meet a certain total megacycles requirement.
    - Query for systems meeting a megacycle requirement, taking a certain overhead% into account.


    The fields returned are (not all are visible by default):
    Vendor, System, CPU, Cores, Chips, CoresPerChip, Speed, Result, ResultPerCore,
    Baseline, MCyclesPerCore, MCyclesTotal, OS and Published.

    .LINK
    http://eightwone.com
    
    .NOTES
    The tool requires internet access to retrieve information from spec.org.
    Progress meter of Invoke-WebRequest disabled to speed up download process.
    Note that SPECint may also return systems which can not run Windows, e.g. SGI.

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial community release
    1.1     Added vCPU:pCPU ratio support, and fixed # of Cores / Chips
    1.2     Renamed to Get-ExchangePQT
            Changed Published to [datetime] for sorting
            Code cleanup
    
    .PARAMETER CPU
    Filter on processor name (partial matching).

    .PARAMETER Vendor
    Filter on vendor name (partial matching).

    .PARAMETER System
    Filter on system name (partial matching).

    .PARAMETER Type
    Type of calculation to perform, Possible values are 2010 for 
    Exchange Server 2010 and 2013 for Exchange Server 2013/2016 (Default).

    .PARAMETER MinCores
    Only return specs of systems with more than this number of cores.
    Can not be used together with MaxCores or Cores.

    .PARAMETER MaxCores
    Only return specs of systems with less than this number of cores.
    Can not be used together with MinCores or Cores.

    .PARAMETER Cores
    Only return specs of systems with this number of cores.
    Can not be used together with MinCores or MaxCores.

    .PARAMETER minChips
    Only return specs of systems more than this number of processors.
    Can not be used together with maxChips or Chips.

    .PARAMETER maxChips
    Only return specs of systems with less than this number of processors.
    Can not be used together with minChips or Chips.

    .PARAMETER Chips
    Only return specs of systems with this number of processors.
    Can not be used together with minChips or maxChips.

    .PARAMETER MinMegaCycles
    Specify the minimum number of total megagacycles returned items should meet.

    .PARAMETER Overhead
    Specify the percentage of overhead to take into account when specifying 
    MinMegaCycles. Default is 0 (0%).

    .PARAMETER Ratio
    Specify the vCPU:pCPU ratio. For example, specify 2 to use a 2:1 vCPU to 
    pCPU ratio. Default is 1 (1:1).

    .PARAMETER vCPU
    Specify the number of vCPU allocated. Default is the number of cores of the system.

    .EXAMPLE
    Calculate the average SpecInt rate for 20 core systems containing an E5-2670 CPU
    .\Get-ExchangePQT.ps1 -CPU 'e5-2670' -Cores 20 | Measure-Object -Average -Property Result

    .EXAMPLE
    Calculate the average SpecInt rate for 20 core systems with an E5-2670 CPU, using a 2:1 
    vCPU:pCPU ratio, allocating 12 vCPU cores
    .\Get-ExchangePQT.ps1 -CPU 'e5-2670' -Cores 20 -Ratio 2-vCPU 12 | Measure-Object -Average -Property Result

    .EXAMPLE
    Calculate average SPECint 2006 rate value a hex-core x5450 systems
    .\Get-ExchangePQT.ps1 -CPU x5470 | Where { $_.Cores -eq 8 } | Measure -Average Result

    .EXAMPLE
    Return the 10 last published entries, displaying system name, number of cores and rate.
    .\Get-ExchangePQT.ps1 | Sort Published -Desc | Select System, Cores, Result, Published -First 10

    .EXAMPLE
    Search all specs for systems using x5470 CPUs, with a minimum of 15,000 megacycles 
    and 20% overhead remaining (net total megacycles = 15,000 + 20% = 18,000)
    .\Get-ExchangePQT.ps1 -CPU x5470  -MinMegaCycles 15000 -Overhead 20 
    
    Search all specs for systems containing x3430 CPUs and export it to a CSV file
    .\Get-ExchangePQT.ps1 -CPU x3430 | Export-CSV -NoTypeInformation .\specint_x5470.csv

#>

[cmdletbinding( DefaultParameterSetName='Default')]
param(
	[parameter( Mandatory=$false, ParameterSetName='Default')]
	[string]$CPU=$null,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
        [string]$Vendor=$null,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
        [string]$System=$null,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
        [ValidateRange(0,100)]
	[int]$Overhead=0,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
	[ValidateRange(1,2)]
	[float]$Ratio,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
	[ValidateRange(1,100)]
	[int]$vCPU,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
	[ValidateRange(0,999999999)]
	[int]$MinMegaCycles=0,
	[parameter( Mandatory=$false, ParameterSetName="Default")]
	[ValidateSet(2010, 2013)] 
	[string]$Type=2013,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
	[parameter( Mandatory=$true, ParameterSetName="CoresMin")]
        [ValidateRange(0,999)]
        [int]$MinCores=0,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
	[parameter( Mandatory=$true, ParameterSetName="CoresMax")]
        [ValidateRange(0,999)]
        [int]$MaxCores=999,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
	[parameter( Mandatory=$true, ParameterSetName="CoresRange")]
        [ValidateRange(0,999)]
        [int]$Cores,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
	[parameter( Mandatory=$true, ParameterSetName="ChipsRange")]
        [ValidateRange(0,999)]
        [int]$Chips,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
	[parameter( Mandatory=$true, ParameterSetName="ChipsMin")]
        [ValidateRange(0,999)]
        [int]$MinChips=0,
	[parameter( Mandatory=$false, ParameterSetName='Default')]
	[parameter( Mandatory=$true, ParameterSetName="ChipsMax")]
        [ValidateRange(0,999)]
        [int]$MaxChips=999
    )

process {

    If( $MinCores -and $MaxCores) {
        If( $MaxCores -lt $MinCores) {
            Write-Error 'Max Cores is not equal or larger than Min Cores'
            Exit -1
        }
    }

    # Construct URL
    $URL= "http://www.spec.org/cgi-bin/osgresults?conf=rint2006&op=dump;format=csvdump&proj-BASE=0&proj-COPIES=0&proj-CPU_MHZ=256"

    If( $Vendor) {
        $URL+= "&proj-COMPANY=256&critop-COMPANY=0&crit-COMPANY=$Vendor"
    }
    If( $System) {
        $URL+= "&proj-SYSTEM=256&critop-SYSTEM=0&crit-SYSTEM=$System"
    }
    If( $CPU) {
        $URL+= "&proj-CPU=0&critop-CPU=0&crit-CPU=$CPU"
    }
    If( $Cores -gt 0) {
        $MinCores= $Cores
        $MaxCores= $Cores
    }
    If( $Chips -gt 0) {
        $MinChips= $Chips
        $MaxChips= $Chips
    }

    Write-Verbose ('Querying SpecInt2006 information (Chips: {0}-{1} Cores: {2}-{3})' -f $MinChips, $MaxChips, $MinCores, $MaxCores)

    # Disable progress bar to speed up downloading ..
    $OldProgress= $ProgressPreference
    $ProgressPreference= "SilentlyContinue"
    $Page= Invoke-WebRequest -Uri $URL -UseBasicParsing -Method Get
    $ProgressPreference= $OldProgress
    $Data= ConvertFrom-CSV $Page 
    If( $Overhead -gt 0) {
        $MegacyclesTreshold= [int]((100+$Overhead)/100*$MinMegaCycles);
    }
    Else {
        $MegacyclesTreshold= $MinMegaCycles;
    }

    Write-Verbose "Total megacycles targeted $MinMegaCycles, with $Overhead% overhead treshold is $MegacyclesTreshold"

    # Baseline score 
    If( $Type -eq 2010) {
        # HP DL380 G5 x5470 3.33GHz, 8 cores (3,333 MHz), score = 150 or 18.75/core
        $MCyclesBaseline= 3333
        $MCyclesCoreScore= 18.75
    }
    Else {
        # HP DL380p G8 Intel Xeon E5-2650 2 GHz, score = 540 or 33.75/core
        $MCyclesBaseline= 2000
        $MCyclesCoreScore= 33.75
    }
    Write-Verbose "Calculating for Exchange Server $Type (Reference megacycles baseline $MCyclesBaseline, core $MCyclesCoreScore)"

    # Specify default display properties
    $defaultProps= @('Vendor','System','Cores','Chips','CoresPerChip','Speed','Result','Published')
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultProps)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

    Write-Verbose "$($Data.Count) number of potential entries found, narrowing results .."

    $Data | ForEach-Object {

        $Baseline= [float]$_.'Baseline'
        $Result= ([float]$_.'Result')
        If( $Result -gt 0 -and $Baseline -gt 0) {
            $ThisCores= [int]($_.'# Cores')
            $ThisChips= [int]($_.'# Chips')
            if( $vCPU -gt 0) {
                $Result= [math]::round(($Result / ( $Cores * $Ratio)) * $vCPU, 2);            
            }
            $ResultPerCore= [math]::round( $Result / $ThisCores, 2);
            $MCyclesPerCore= [math]::round(( $Result / $ThisCores * $MCyclesBaseline) / $MCyclesCoreScore, 2);
            $MCyclesTotal= $MCyclesPerCore * $ThisCores;
            If( $MCyclesTotal -ge $MegacyclesTreshold -and $ThisCores -ge $MinCores -and $ThisCores -le $MaxCores -and $ThisChips -ge $MinChips -and $ThisChips -le $MaxChips) {
                $Props = @{
                    'Vendor'= $_.'Hardware Vendor	';
                    'System'=$_.'System';
                    'CPU'=$_.'Processor';
                    'Speed'=$_.'Processor MHz';
                    'Cores'= $ThisCores;
                    'Chips'=$ThisChips;
                    'CoresPerChip'=$_.'# Cores Per Chip';
                    'OS'=$_.'Operating System';
                    'Result'= $Result;
		    'ResultPerCore'= $ResultPerCore;
                    'Baseline'= $Baseline;
                    'MCyclesPerCore'= $MCyclesPerCore;
                    'MCyclesTotal'= $MCyclesTotal;
                    'Published'= [datetime](([datetime]::Parse( $_.'Published')).ToShortDateString());
                }
                $Object= New-Object -Typename PSObject -Prop $Props
                $Object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                Write-Output $Object
            }
            Else {
                # System does not meet megacycles treshold
            }
        }
    }
}