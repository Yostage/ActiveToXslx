function ConvertFrom-XLSx {
  param ([parameter(             Mandatory=$true,
                         ValueFromPipeline=$true, 
           ValueFromPipelineByPropertyName=$true)]
         [string]$path , 
         [switch]$PassThru
        )

  begin { $objExcel = New-Object -ComObject Excel.Application }
Process { if ((test-path $path) -and ( $path -match ".xl\w*$")) {
                    $path = (resolve-path -Path $path).path 
                $savePath = $path -replace ".xl\w*$",".csv"
              $objworkbook=$objExcel.Workbooks.Open( $path)
              $objworkbook.SaveAs($savePath,6) # 6 is the code for .CSV 
              $objworkbook.Close($false) 
              if ($PassThru) {Import-Csv -Path $savePath } 
          }
          else {Write-Host "$path : not found"} 
        } 
   end  { $objExcel.Quit() }
}


function repro()
{ 
  ConvertFrom-XLSx $file
  # TODO: remove first line of the CSV
  $records = import-csv $file
  # strip out all the wrong records, which have no name?
  $filtered = $records | ? { !(bogus-record $_)}
}

function bogus-record($rec)
{
  # note $records[3] is bogus
  return ($rec.('Participant: Name') -eq "" -or
          $rec.('Participant: Name') -eq "Participant: Name")
}

# rewrite the properties:
function write-record($rec)
{
  write-output "Name: $($rec.('Primary P/G: Name'))"
  write-output "Address: $($rec.('Primary P/G: Address'))"
  write-output "Phone: $($rec.('Primary P/G: Cell phone number'))"
  write-output "Email Address: $($rec.('Primary P/G: Email address'))"
  write-output "Child's name/DOB: $($rec.('Participant: Name')) $($rec.('Participant: Date of birth'))"
  write-output "Class: $($rec.('Session name'))"
  write-output ""
  # todo: class, date of registration   
}

