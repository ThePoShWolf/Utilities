<#
.SYNOPSIS
Converts bytes to a more readable KB, MB, GB, or TB
.DESCRIPTION
Format-Bytes takes a bytes sized input and formats it to be KB, MB, GB, or TB depending on its size.
.PARAMETER $Num
The number to convert to KB, MB, GB, or TB.
.EXAMPLE
Format-Bytes 123456789

Output: 117.74 MB
.EXAMPLE
$x = (123456,456789456,789456123456)
$x | Format-Bytes

Output: 
120.56 KB
435.63 MB
735.24 GB
.INPUTS
A number of any type.
.OUTPUTS
A number converted to a string with it's type identifier.
.NOTES
Author: Town of An
.LINK
KB
MB
GB
TB
#>
Function Format-Bytes
{
	Param
	(
		[Parameter(ValueFromPipeline=$true,Mandatory=$true)]
		[float[]]$number
	)
	Begin
	{
		
	}
	Process
	{
		ForEach($num in $number)
		{
			If($num -lt 1KB)
			{
				Return "$num B"
			}
			If($num -lt 1MB)
			{
				$num = $num/1KB
				$num = "{0:N2}" -f $num
				Return "$num KB"
			}
			If($num -lt 1GB)
			{
				$num = $num/1MB
				$num = "{0:N2}" -f $num
				Return "$num MB"
			}
			If($num -lt 1TB)
			{
				$num = $num/1GB
				$num = "{0:N2}" -f $num
				Return "$num GB"
			}
			Else
			{
				$num = $num/1TB
				$num = "{0:N2}" -f $num
				Return "$num TB"
			}
		}
	}
	End
	{
		
	}
}