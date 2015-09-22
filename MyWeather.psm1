#requires -version 3.0


<#
 Jeffery Hicks
 http://jdhitsolutions.com/blog
 follow on Twitter: http://twitter.com/JeffHicks

 
 "Those who forget to script are doomed to repeat their work."

 Learn more about PowerShell:
 http://jdhitsolutions.com/blog/essential-powershell-resources/
 
  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************
#>


Function Get-Woeid {


Param(
[parameter(Position=0,
Mandatory,
ValueFromPipeline,
HelpMessage = "Enter a place name or postal (zip) code.")]
[string[]]$Search,
[switch]$AsXML
)

Begin {
    Write-Verbose "Starting $($myinvocation.mycommand)"
}

Process {
    foreach ($item in $search) {
        Write-Verbose "Querying for $search"
        $uri = "https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20geo.places%20where%20text%3D'$ITEM'%20limit%201"
        
        Write-Verbose $uri
        [xml]$xml = Invoke-RestMethod -Uri $uri

        if ($AsXML) {
            Write-Verbose "Writing XML document"
            $xml
            Write-Verbose "Ending function"
            #bail out since we're done
            Return
        }

        if ($xml.query.results.place.woeid) {
            Write-Verbose "Parsing XML into an ordered hashtable"
            $hash = [ordered]@{
                WOEID   = $xml.query.results.place.woeid
                Locale  = $xml.query.results.place.locality1.'#text'
                Region  = $xml.query.results.place.admin1.'#text'
                Postal  = $xml.query.results.place.postal.'#text'
                Country = $xml.query.results.place.country.code            
            }
            Write-Verbose "Writing custom object"
            New-Object -TypeName PSObject -Property $hash

        } #if $xml
        else {
            Write-Warning "Failed to find anything for $item"
        }
    } #foreach
} #process

End {
    Write-Verbose "Ending $($myinvocation.mycommand)"
 }
} #end function

Function Get-Weather {


[cmdletbinding(DefaultParameterSetName="detail")]

Param(
[Parameter(Position=0,
Mandatory,
HelpMessage="Enter a WOEID value",
ValueFromPipeline,
ValueFromPipelineByPropertyName
)]
[ValidateNotNullorEmpty()]
[Alias("id")]
[string[]]$woeid,
    
[Parameter(Position=1,ParameterSetName="detail")]
[ValidateSet("Basic","Extended","All")]
[ValidateNotNullorEmpty()]
[String]$Detail = "Basic",

[ValidateSet("f","c")]
[String]$Unit = "f",
    
[Parameter(ParameterSetName="online")]
[Switch]$Online,

[Parameter(ParameterSetName="xml")]
[Switch]$AsXML
    
)

Begin {    
	Write-Verbose "Starting $($myinvocation.mycommand)"
    Write-Verbose "PSBoundparameters"
    Write-Verbose ($PSBoundParameters | out-String).Trim()
	
	#define property sets
	$BasicProp = @("Date","Location","Temperature","Condition","ForecastCondition",
	"ForecastLow","ForecastHigh")
	$ExtendedProp = $BasicProp
	$ExtendedProp += @("WindChill","WindSpeed","Humidity","Barometer",
	"Visibility","Tomorrow")
	$AllProp=$ExtendedProp
	$AllProp += @("Sunrise","Sunset","City","Region","Latitude",
	"Longitude","ID","URL")
	
	#unit must be lower case
	$Unit = $Unit.ToLower()
	Write-Verbose "Using unit $($unit.toLower())"
    
    #define base url string
    [string]$uribase = "http://weather.yahooapis.com/forecastrss"
    Write-Verbose "Base = $uribase"
 } #begin
 
Process {
    Write-Verbose "Processing"
    
	foreach ($id in $woeid) {
		Write-Verbose "Getting weather info for woeid: $id"
	    
	    #define a uri for the given WOEID
        [string]$uri = "{0}?w={1}&u={2}" -f $uribase,$id,$unit
	     if ($online) {
	    	Write-Verbose "Opening $uri in web browser"
	    	Start-Process -FilePath $uri
            #bail out since there's nothing else to do.
            Return
	    }
        Write-Verbose "Downloading $uri"
	    [xml]$xml= Invoke-WebRequest -uri $uri

        if ($AsXML) {
            Write-Verbose "Writing XML document"
            $xml
            Write-Verbose "Ending function"
            #bail out since we're done
            Return
        }

	    if ($xml.rss.channel.item.Title -eq "City not found") {
	        Write-Warning "Could not find a location for $id"}
	    else {
	    	Write-Verbose "Processing xml"
	    	
	        #initialize a new hash table
	        $properties=@{}
	    	<#
            get the yweather nodes	
        	yweather information comes from a different namespace so we'll
            use Select-XML to extract the data. Parsing out all data regardless
            of requested detail since it doesn't take much effort. Later,
            only the requested detail will be written to the pipeline.
            #>

        	#define the namespace hash
        	$namespace = @{yweather=$xml.rss.yweather}
        	$units = (Select-Xml -xml $xml -XPath "//yweather:units" -Namespace $namespace ).node
	    	
	        $properties.Add("Condition",$xml.rss.channel.item.condition.text)
	        $properties.Add("Temperature","$($xml.rss.channel.item.condition.temp) $($units.temperature)")
            #convert Date to a [datetime] object
            $dt = $xml.rss.channel.item.condition.date.Substring(0,$xml.rss.channel.item.condition.Date.LastIndexOf("m")+1) -as [datetime]
	        $properties.Add("Date",$dt)
	    
	        #get forecast
            $properties.add("ForecastDate",$xml.rss.channel.item.forecast[0].date)
	        $properties.add("ForecastCondition",$xml.rss.channel.item.forecast[0].text )
	        $properties.Add("ForecastLow","$($xml.rss.channel.item.forecast[0].low) $($units.temperature)" )
	        $properties.Add("ForecastHigh","$($xml.rss.channel.item.forecast[0].high) $($units.temperature)" )
	        
	        #build tomorrow's foreacst
	        $t = $xml.rss.channel.item.forecast[1]
	        $tomorrow = "{0} {1} {2} Low {3}{4}: High: {5}{6}" -f $t.day,$t.Date,$t.Text,$t.low, $($units.temperature),$t.high, $($units.temperature)
	        $properties.add("Tomorrow",$tomorrow)
	        
	        #get optional information
          	$properties.Add("Latitude",$xml.rss.channel.item.lat)
        	$properties.Add("Longitude",$xml.rss.channel.item.long)
        	$city = $xml.rss.channel.location.city
        	$properties.Add("City",$city)
        	$region = $xml.rss.channel.location.region
            $country = $xml.rss.channel.location.country
        	
        	if (-not ($region)) {
               #if no region found then use country
               Write-Verbose "No region found. Using Country"
                $region=$country
            }
            $properties.Add("Region",$region)
            $location = "{0}, {1}" -f $city,$region
	        $properties.Add("Location",$location)	    	 
	        $properties.Add("ID",$id)
        	
            #get additional yweather information        	
        	$wind = (Select-Xml -xml $xml -XPath "//yweather:wind" -Namespace $namespace ).node
        	$astronomy = (Select-Xml -xml $xml -XPath "//yweather:astronomy" -Namespace $namespace ).node
        	$atmosphere = (Select-Xml -xml $xml -XPath "//yweather:atmosphere" -Namespace $namespace ).node
        	
        	$properties.Add("WindChill","$($wind.chill) $($units.temperature)")
        	$properties.Add("WindSpeed","$($wind.speed) $($units.speed)")
        	$properties.Add("Humidity","$($atmosphere.humidity)%")
        	$properties.Add("Visibility","$($atmosphere.visibility) $($units.distance)")
        	
            #decode rising
        	switch ($atmosphere.rising) {
        		0 {$state="steady"}
        		1 {$state="rising"}
        		2 {$state="falling"}
        	}
         	$properties.Add("Barometer","$($atmosphere.pressure) $($units.pressure) and $state")
        	$properties.Add("Sunrise",$astronomy.sunrise)
        	$properties.Add("Sunset",$astronomy.sunset)
        	$properties.Add("url",$uri)
           
		   #create new object to hold values
		   $obj = New-Object -TypeName PSObject -Property $properties
		   
		   #write object and properties. Default is Basic
		   Switch ($detail) {
           "All"  {
		   	     Write-Verbose "Using All properties"
		   	     $obj | Select-Object -Property $AllProp
		    } #all
		   "Extended"  {
		   	    Write-Verbose "Using Extended properties"
		   	    $obj | Select-Object -Property $ExtendedProp
		   } #extended
		   Default {
		   	     Write-Verbose "Using Basic properties"
		   	     $obj | Select-Object -Property $BasicProp
		     } #default
	       } #Switch
	   } #processing XML
   } #foreach $id

 } #process
 
 End {
 	Write-Verbose "Ending $($myinvocation.mycommand)"
 }  #end 
    
} #end function

###############################################################################
#define some aliases

Set-Alias -Name gw -Value get-weather
Set-Alias -Name gwid -Value get-woeid

Export-ModuleMember -alias * -Function *

