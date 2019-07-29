$file="C:\temp\GetMyNumbers\terg.csv"
$reportpath="C:\temp\GetMyNumbers\report.csv"
$initial1="KJP"
$initial2="KP"
$data = import-csv $file | where {$_.processor_initials -Match $initial1 -or $_.ringer_initials -match $initial1 -or $_.processor_initials -Match $initial2 -or $_.ringer_initials -match $initial2}



class Bird
{
[string]$SpeciesName
[int]$New
[int]$Pulli
[int]$Subsequent
}

#####################################################################################################

# This loops through the records and takes out individual species names and put them in the HT specieslist
$specieslist=@{} 
foreach ($record in $data) # Loop through the records#
    {
    if ($specieslist.keys -notcontains $record.species_name) # If hash table doesnt contain a species.....
        {
       $specieslist.add($record.species_name, (new-object bird))  # Add the species to the hash table as a key 
        $specieslist.($record.species_name).speciesname = $record.species_name
        $specieslist.($record.species_name).new = 0
        $specieslist.($record.species_name).pulli = 0
        $specieslist.($record.species_name).subsequent = 0
        }   
    }
#############################################################################################################

#$b=@($specieslist).Keys  # i had to copy the hashtable to a new variable so that we arent modifyint he array that we ar elooppig through


foreach ($encounter in $data)
{
 if ($encounter.record_type -eq "N")
            {
               $specieslist.($encounter.species_name).new ++
            }
elseif ($encounter.record_type -eq "S")
            {
               $specieslist.($encounter.species_name).subsequent ++
            }
                   
 if ($encounter.age -match "1" )#-or $encounter.age -match "1J" )
            {
               $specieslist.($encounter.species_name).pulli ++
            }   
}
$specieslist.Values

$specieslist.Values|export-csv -Path $reportpath 
