#requires -Version 3
param(
    [Parameter(Mandatory=$false, HelpMessage='Weekly or Daily')]
    [ValidateSet( 'Weekly','Daily' )] 
    [string]$ChartTime = 'Daily', 

    [Parameter(Mandatory=$false, HelpMessage='Temp, Wind or Wave')]
    [ValidateSet( 'Temp', 'Wind', 'Wave' )] 
    [string]$ChartType = 'Wind'
)

function Make-TempLineChart() {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage='Chart Data')]
        [Object[]]$data,
        [Parameter(Mandatory=$True)]
        [string]$ChartTitle = '<<TITLE>>',
        [Parameter(Mandatory=$True)]
        [string]$outFile = "$scriptpath\SplineArea.png"

    )

    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

    # chart object
    $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart1.Width = 1920
    $chart1.Height = 1080
    $chart1.BackColor = [System.Drawing.Color]::White

    # title 
    [void]$chart1.Titles.Add($ChartTitle)
    $chart1.Titles[0].Font = "Arial,13pt"
    $chart1.Titles[0].Alignment = "topLeft"

    # chart area 
    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartarea.Name = "ChartArea1"
    $chartarea.AxisY.Title = "Temp"
    $chartarea.AxisY.Interval = 5

    #Min is Min of data
    $minmax = $data | Measure-Object -Property atmp1,wtemp1 -Maximum -Minimum
    $min = if ($minmax[0].Minimum -lt $minmax[1].Minimum) {$minmax[0].Minimum} else {$minmax[1].Minimum}
    $max = if ($minmax[0].Maximum -gt $minmax[1].Maximum) {$minmax[0].Maximum} else {$minmax[1].Maximum}
    #nearest 5
    $chartarea.AxisY.Minimum = [math]::Truncate($min / 5) * 5
    $chartarea.AxisY.Maximum = [math]::ceiling($max / 5) * 5


    $chartarea.AxisX.Title = "Date"
    if ($data.Count -gt 200) {
        $chartarea.AxisX.Interval = 7
    } else {
        $chartarea.AxisX.Interval = 3
    }
    $chartarea.AxisX.MajorGrid.Enabled = $False
    $chart1.ChartAreas.Add($chartarea)

    # legend 
    $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
    $legend.name = "Legend1"
    $chart1.Legends.Add($legend)

    # data series
    [void]$chart1.Series.Add("AirTemp")
    $chart1.Series["AirTemp"].ChartType = "Line"
    $chart1.Series["AirTemp"].BorderWidth = 3
    $chart1.Series["AirTemp"].IsVisibleInLegend = $true
    $chart1.Series["AirTemp"].chartarea = "ChartArea1"
    $chart1.Series["AirTemp"].Legend = "Legend1"
    $chart1.Series["AirTemp"].color = "#62B5CC"
    $data | ForEach-Object {$chart1.Series["AirTemp"].Points.addxy( $_.date , $_.atmp1) } | Out-Null

    # data series
    [void]$chart1.Series.Add("WaterTemp")
    $chart1.Series["WaterTemp"].ChartType = "Line"
    $chart1.Series["WaterTemp"].IsVisibleInLegend = $true
    $chart1.Series["WaterTemp"].BorderWidth = 3
    $chart1.Series["WaterTemp"].chartarea = "ChartArea1"
    $chart1.Series["WaterTemp"].Legend = "Legend1"
    $chart1.Series["WaterTemp"].color = "#E3B64C"
    $data | ForEach-Object {$chart1.Series["WaterTemp"].Points.addxy( $_.date , $_.wtemp1) } | Out-Null

    # save chart
    $chart1.SaveImage($outFile,"png")
}

function Make-WindLineChart() {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage='Chart Data')]
        [Object[]]$data,
        [Parameter(Mandatory=$True)]
        [string]$ChartTitle = '<<TITLE>>',
        [Parameter(Mandatory=$True)]
        [string]$outFile = "$scriptpath\SplineArea.png"

    )

    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

    # chart object
    $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart1.Width = 1920
    $chart1.Height = 1080
    $chart1.BackColor = [System.Drawing.Color]::White

    # title 
    [void]$chart1.Titles.Add($ChartTitle)
    $chart1.Titles[0].Font = "Arial,13pt"
    $chart1.Titles[0].Alignment = "topLeft"

    # chart area 
    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartarea.Name = "ChartArea1"
    $chartarea.AxisY.Title = "Wind"
    $chartarea.AxisY.Interval = 5

    #Min is Min of data wspd1,gust1
    $minmax = $data | Measure-Object -Property wspd1,gust1 -Maximum -Minimum
    $min = if ($minmax[0].Minimum -lt $minmax[1].Minimum) {$minmax[0].Minimum} else {$minmax[1].Minimum}
    $max = if ($minmax[0].Maximum -gt $minmax[1].Maximum) {$minmax[0].Maximum} else {$minmax[1].Maximum}
    #nearest 5
    $chartarea.AxisY.Minimum = [math]::Truncate($min / 5) * 5
    $chartarea.AxisY.Maximum = [math]::ceiling($max / 5) * 5


    $chartarea.AxisX.Title = "Date"
    if ($data.Count -gt 200) {
        $chartarea.AxisX.Interval = 7
    } else {
        $chartarea.AxisX.Interval = 3
    }
    $chartarea.AxisX.MajorGrid.Enabled = $False
    $chart1.ChartAreas.Add($chartarea)

    # legend 
    $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
    $legend.name = "Legend1"
    $chart1.Legends.Add($legend)

    # data series
    [void]$chart1.Series.Add("WindSpeed")
    $chart1.Series["WindSpeed"].ChartType = "Line"
    $chart1.Series["WindSpeed"].BorderWidth = 3
    $chart1.Series["WindSpeed"].IsVisibleInLegend = $true
    $chart1.Series["WindSpeed"].chartarea = "ChartArea1"
    $chart1.Series["WindSpeed"].Legend = "Legend1"
    $chart1.Series["WindSpeed"].color = "#62B5CC"
    $data | ForEach-Object {$chart1.Series["WindSpeed"].Points.addxy( $_.date , $_.wspd1) } | Out-Null

    # data series
    [void]$chart1.Series.Add("Gust")
    $chart1.Series["Gust"].ChartType = "Line"
    $chart1.Series["Gust"].IsVisibleInLegend = $true
    $chart1.Series["Gust"].BorderWidth = 3
    $chart1.Series["Gust"].chartarea = "ChartArea1"
    $chart1.Series["Gust"].Legend = "Legend1"
    $chart1.Series["Gust"].color = "#E3B64C"
    $data | ForEach-Object {$chart1.Series["Gust"].Points.addxy( $_.date , $_.gust1) } | Out-Null

    # save chart
    $chart1.SaveImage($outFile,"png")
}

function Make-WaveLineChart() {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage='Chart Data')]
        [Object[]]$data,
        [Parameter(Mandatory=$True)]
        [string]$ChartTitle = '<<TITLE>>',
        [Parameter(Mandatory=$True)]
        [string]$outFile = "$scriptpath\SplineArea.png"

    )

    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

    # chart object
    $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart1.Width = 1920
    $chart1.Height = 1080
    $chart1.BackColor = [System.Drawing.Color]::White

    # title 
    [void]$chart1.Titles.Add($ChartTitle)
    $chart1.Titles[0].Font = "Arial,13pt"
    $chart1.Titles[0].Alignment = "topLeft"

    # chart area 
    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartarea.Name = "ChartArea1"
    $chartarea.AxisY.Title = "Wind"
    $chartarea.AxisY.Interval = 5

    #Min is Min of data - wvhgt,dompd
    $minmax = $data | Measure-Object -Property wvhgt,dompd -Maximum -Minimum
    $min = if ($minmax[0].Minimum -lt $minmax[1].Minimum) {$minmax[0].Minimum} else {$minmax[1].Minimum}
    $max = if ($minmax[0].Maximum -gt $minmax[1].Maximum) {$minmax[0].Maximum} else {$minmax[1].Maximum}
    #nearest 5
    $chartarea.AxisY.Minimum = [math]::Truncate($min / 5) * 5
    $chartarea.AxisY.Maximum = [math]::ceiling($max / 5) * 5


    $chartarea.AxisX.Title = "Date"
    if ($data.Count -gt 200) {
        $chartarea.AxisX.Interval = 7
    } else {
        $chartarea.AxisX.Interval = 3
    }
    $chartarea.AxisX.MajorGrid.Enabled = $False
    $chart1.ChartAreas.Add($chartarea)

    # legend 
    $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
    $legend.name = "Legend1"
    $chart1.Legends.Add($legend)

    # data series
    [void]$chart1.Series.Add("WaveHeight")
    $chart1.Series["WaveHeight"].ChartType = "Line"
    $chart1.Series["WaveHeight"].BorderWidth = 3
    $chart1.Series["WaveHeight"].IsVisibleInLegend = $true
    $chart1.Series["WaveHeight"].chartarea = "ChartArea1"
    $chart1.Series["WaveHeight"].Legend = "Legend1"
    $chart1.Series["WaveHeight"].color = "#62B5CC"
    $data | ForEach-Object {$chart1.Series["WaveHeight"].Points.addxy( $_.date , $_.wvhgt) } | Out-Null

    # data series
    [void]$chart1.Series.Add("WavePeriod")
    $chart1.Series["WavePeriod"].ChartType = "Line"
    $chart1.Series["WavePeriod"].IsVisibleInLegend = $true
    $chart1.Series["WavePeriod"].BorderWidth = 3
    $chart1.Series["WavePeriod"].chartarea = "ChartArea1"
    $chart1.Series["WavePeriod"].Legend = "Legend1"
    $chart1.Series["WavePeriod"].color = "#E3B64C"
    $data | ForEach-Object {$chart1.Series["WavePeriod"].Points.addxy( $_.date , $_.dompd) } | Out-Null

    # save chart
    $chart1.SaveImage($outFile,"png")
}


$URL = "http://www.iiseagrant.org/wilmettebuoy/doGraphs.php"
$scriptpath = Split-Path -parent $MyInvocation.MyCommand.Definition

switch ("$ChartType")
{
  'temp' { 
        $body = "type=watemp&range="
        $Title = ( "Wilmette Bouy Temperatures retrieved on " + (Get-Date) )
        $File = "$scriptpath\Tempurature.png"
        break
    }
  'wave' { 
        $body = "type=whap&range=" 
        $Title = ( "Wilmette Bouy Wave Height retrieved on " + (Get-Date) )
        $File = "$scriptpath\Wave.png"
        break
    }
  'wind' { 
        $body = "type=wind&range=" 
        $Title = ( "Wilmette Bouy Wind and Wind Gusts retrieved on " + (Get-Date) )
        $File = "$scriptpath\Wind.png"
        break
    }
  default { 
        $body = "type=wind&range=" 
        $Title = ( "Wilmette Bouy Wind and Wind Gusts retrieved on " + (Get-Date) )
        $File = "$scriptpath\Wind.png"
        break
    }     
}
$PostResult = Invoke-WebRequest -Uri $URL -Body ( $body + $ChartTime).ToLower() -Method Post

if ($PostResult.StatusCode -eq 200) {
    $Data = $PostResult.Content | ConvertFrom-Json

    switch ("$ChartType")
    {
        'temp' {
            #Have to convert from C to F
            foreach ($d in $Data) {
                $d.atmp1 = ([double]$d.atmp1 * 1.8 + 32)
                $d.wtemp1 = ([double]$d.wtemp1 * 1.8 + 32)
            }
            Make-TempLineChart -data $Data -ChartTitle $Title -outfile $File
            break
        }
        'wave' { 
            #data is wspd1,gust1
            Make-WaveLineChart -data $Data -ChartTitle $Title -outfile $File
            break
        }
        'wind' { 
            #data is wvhgt,dompd
            Make-WindLineChart -data $Data -ChartTitle $Title -outfile $File
            break
        }
        default { 
            #data is wspd1,gust1
            Make-WindLineChart -data $Data -ChartTitle $Title -outfile $File
            break
        }     
    }
    
    start $file
}