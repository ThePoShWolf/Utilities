Function Format-Bytes {
    Param
    (
        [Parameter(
            ValueFromPipeline = $true
        )]
        [ValidateNotNullOrEmpty()]
        [float]$number
    )
    Begin{
        $sizes = 'KB','MB','GB','TB','PB'
    }
    Process {
        # New for loop
        for($x = 0;$x -lt $sizes.count; $x++){
            if ($number -lt [int64]"1$($sizes[$x])"){
                if ($x -eq 0){
                    return "$number B"
                } else {
                    $num = $number / [int64]"1$($sizes[$x-1])"
                    $num = "{0:N2}" -f $num
                    return "$num $($sizes[$x-1])"
                }
            }
        }
        <# Original way
        if ($number -lt 1KB) {
            "$number B"
        } elseif ($number -lt 1MB) {
            $number = $number / 1KB
            $number = "{0:N2}" -f $number
            "$number KB"
        } elseif ($number -lt 1GB) {
            $number = $number / 1MB
            $number = "{0:N2}" -f $number
            "$number MB"
        } elseif ($number -lt 1TB) {
            $number = $number / 1GB
            $number = "{0:N2}" -f $number
            "$number GB"
        } elseif ($number -lt 1PB) {
            $number = $number / 1TB
            $number = "{0:N2}" -f $number
            "$number TB"
        } else {
            $number = $number / 1PB
            $number = "{0:N2}" -f $number
            "$number PB"
        }
        #>
    }
    End{}
}