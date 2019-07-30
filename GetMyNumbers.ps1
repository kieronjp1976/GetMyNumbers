##### Kieron Palmer July 2019#############################
##### Update the five variables below with your own file paths and initials.
##################################################################
##################################################################
$file="C:\temp\GetMyNumbers\terg.csv"
$Initial1="KJP"
$Initial2="KP"  # If you only have one intitial enter the same one twice
$path="C:\temp\GetMyNumbers\output"
$pdf="C:\temp\GetMyNumbers\pdf.pdf"
###############################################################################




$data = import-csv $file | where-object {$_.processor_initials -Match $initial1 -or $_.ringer_initials -match $initial1 -or $_.processor_initials -Match $initial2 -or $_.ringer_initials -match $initial2}


class Bird # Create a class to hold the data defined below
{
[string]$SpeciesName
[int]$New
[int]$Pulli
[int]$Subsequent
}



function New-SpreadSheet # This can move to a module in future. 
{
$script:excel = New-Object -ComObject excel.application
$excel.visible = $false
$script:workbook = $excel.Workbooks.Add()
$script:Worksheet= $workbook.Worksheets.Item(1)
########################################  Add the Column names:  ROW, COLUMN
$Worksheet.Name = "Bird Numbers"
#$Worksheet.Cells.Item(1,1) = 'Bird Numbers'
$Worksheet.Cells.Item(1,1) = 'Species'
$Worksheet.Cells.Item(1,2) = 'New birds'
$Worksheet.Cells.Item(1,3) = 'Pulli'
$Worksheet.Cells.Item(1,4) = 'Retrap'
$Worksheet.Cells.Item(1,5) = 'Total'
}
#############################################################################
function Save-Spreadsheet
{
$excel.DisplayAlerts = $False # suppress overwrite alert


$workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $pdf)
$workbook.SaveAs($path) 
$workbook.Close
$excel.Quit()
}

#####################################################################################################

# This loops through the imported records, takes out individual species names and puts them in the hash table $specieslist
$specieslist=@{} 
foreach ($record in $data) # Loop through the records#
    {
    if ($specieslist.keys -notcontains $record.species_name) # If hash table doesnt contain a species.....
        {
        $specieslist.add($record.species_name, (new-object bird))  # Add the species to the hash table as a key and an a object based on our class to the value of that key
        $specieslist.($record.species_name).speciesname = $record.species_name
        $specieslist.($record.species_name).new = 0
        $specieslist.($record.species_name).pulli = 0
        $specieslist.($record.species_name).subsequent = 0
        }   
    }
#############################################################################################################



foreach ($encounter in $data) # loops through tallying up the new, retraps and pulli and adds them to the HT
{
 if ($encounter.record_type -eq "N")
            {
               $specieslist.($encounter.species_name).new ++
            }
elseif ($encounter.record_type -eq "S")
            {
               $specieslist.($encounter.species_name).subsequent ++
            }
                   
 if ($encounter.age -match "1" )
            {
               $specieslist.($encounter.species_name).pulli ++
            }   
}

New-Spreadsheet # creates a spreadhseet using the function defined earlier

$i=2 # Row increment
foreach ($species in $specieslist.keys) #iterate the hash table and add the values into excel
{
$Worksheet.Cells.Item($i,1) = $specieslist.$species.speciesname
$Worksheet.Cells.Item($i,2) = $specieslist.$species.new
$Worksheet.Cells.Item($i,3) = $specieslist.$species.pulli
$Worksheet.Cells.Item($i,4) = $specieslist.$species.subsequent
$a = "=SUM(B" + $i + ":D" + $i + ")"
$Worksheet.Cells.Item($i,5) = $a

$i++
}

$Worksheet.UsedRange.columns.AutoFit() # Autofit the columns
$Range = $Worksheet.usedrange
$sort_col=$Worksheet.Range("B1")
#$Range.Sort($sort_col,2,$empty_Var,$empty_Var,$empty_Var,$empty_Var,$empty_Var,1)
$Range.Sort($sort_col,2)
$header=$worksheet.Range("A1:E1")
$header.Interior.ColorIndex = 16


Save-Spreadsheet
