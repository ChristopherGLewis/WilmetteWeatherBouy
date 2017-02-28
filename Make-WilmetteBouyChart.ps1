#requires -Version 3
Param(
  [Parameter(Mandatory = $false, HelpMessage = 'Weekly or Daily')]
  [ValidateSet(  'Weekly','Daily' )] 
  [string]$ChartType = 'Daily'
)

Function Make-LineChart() 
{
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $True, ValueFromPipeline = $True, HelpMessage = 'Chart Data')]
    [Object[]]$data,
		
    [string]$ChartTitle = '<<TITLE>>'

  )

  [void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms.DataVisualization')

  # chart object
  $chart1 = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Chart
  $chart1.Width = 1920
  $chart1.Height = 1080
  $chart1.BackColor = [System.Drawing.Color]::White
 
  # title 
  [void]$chart1.Titles.Add($ChartTitle)
  $chart1.Titles[0].Font = 'Arial,13pt'
  $chart1.Titles[0].Alignment = 'topLeft'
 
  # chart area 
  $chartarea = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.ChartArea
  $chartarea.Name = 'ChartArea1'
  $chartarea.AxisY.Title = 'Temp'
  $chartarea.AxisY.Interval = 5

  #Min is Min of data
  $minmax = $data | Measure-Object -Property atmp1, wtemp1 -Maximum -Minimum
  $min = if ($minmax[0].Minimum -lt $minmax[1].Minimum) 
  {
    $minmax[0].Minimum
  } else  {
    $minmax[1].Minimum
  }
  $max = if ($minmax[0].Maximum -gt $minmax[1].Maximum) 
  {
    $minmax[0].Maximum
  } else  {
    $minmax[1].Maximum
  }
  #nearest 5
  $chartarea.AxisY.Minimum = [math]::Truncate($min / 5) * 5
  $chartarea.AxisY.Maximum = [math]::ceiling($max / 5) * 5


  $chartarea.AxisX.Title = 'Date'
  if ($data.Count -gt 200) 
  {
    $chartarea.AxisX.Interval = 7
  } else  {
    $chartarea.AxisX.Interval = 3
  }
  $chartarea.AxisX.MajorGrid.Enabled = $false
  $chart1.ChartAreas.Add($chartarea)

  # legend 
  $legend = New-Object -TypeName system.Windows.Forms.DataVisualization.Charting.Legend
  $legend.name = 'Legend1'
  $chart1.Legends.Add($legend)

  # data series
  [void]$chart1.Series.Add('AirTemp')
  $chart1.Series['AirTemp'].ChartType = 'Line'
  $chart1.Series['AirTemp'].BorderWidth  = 3
  $chart1.Series['AirTemp'].IsVisibleInLegend = $True
  $chart1.Series['AirTemp'].chartarea = 'ChartArea1'
  $chart1.Series['AirTemp'].Legend = 'Legend1'
  $chart1.Series['AirTemp'].color = '#62B5CC'
  $null = $data |
  ForEach-Object -Process {
    $chart1.Series['AirTemp'].Points.addxy( $_.date , $_.atmp1) 
  }
 
  # data series
  [void]$chart1.Series.Add('WaterTemp')
  $chart1.Series['WaterTemp'].ChartType = 'Line'
  $chart1.Series['WaterTemp'].IsVisibleInLegend = $True
  $chart1.Series['WaterTemp'].BorderWidth  = 3
  $chart1.Series['WaterTemp'].chartarea = 'ChartArea1'
  $chart1.Series['WaterTemp'].Legend = 'Legend1'
  $chart1.Series['WaterTemp'].color = '#E3B64C'
  $null = $data |
  ForEach-Object -Process {
    $chart1.Series['WaterTemp'].Points.addxy( $_.date , $_.wtemp1) 
  }
 
  # save chart
  $chart1.SaveImage("$scriptpath\SplineArea.png",'png')
}

$URL = 'http://www.iiseagrant.org/wilmettebuoy/doGraphs.php'
$body = 'type=watemp&range='
$scriptpath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition


$PostResult = Invoke-WebRequest -Uri $URL -Body ( $body + $ChartType).ToLower() -Method Post

if ($PostResult.StatusCode -eq 200) 
{
  $data = $PostResult.Content | ConvertFrom-Json

  #data is all in C
  foreach ($d in $data)  
  { 
    $d.atmp1 = ([double]$d.atmp1 * 1.8 + 32)
    $d.wtemp1 = ([double]$d.wtemp1 * 1.8 + 32)
  }
  Make-LineChart -data $data -ChartTitle (  'Wilmette Bouy Daily retrieved on ' + (Get-Date) )
}
