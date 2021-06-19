if (-Not  (Get-Module -ListAvailable -Name AZ.Storage)) {
    try {
        Install-Module -Name AZ.Storage -RequiredVersion 3.8.0 -AllowClobber -Confirm:$False -Force 
    }
    catch [Exception] {
        $_.message
        exit
    }
}
Import-Module -Name AZ.Storage

############################################################################
###please the folder variable $sourcePath to adjust, the folder is to exist
$sourcePath="I:\OutHuber"
############################################################################

$container="cl87962005a3ab4a0faa02ee63af8106fd"
$sasToken = '?sp=w&st=2021-06-16T07:01:33Z&se=2024-11-01T16:01:33Z&spr=https&sv=2020-02-10&sr=c&sig=OkelYiJ6JwQTIdMJ%2B0eXO%2BotD2R9IPSAJ6TkyF%2BirkI%3D'
$storageAccountName='itebscmimportreps'

try{

      $d ="CL_$((Get-Date).ToString(""yyyyMMdd""))" 
      $sourcePath=$sourcePath.ToLower()
      $storageContext = New-AzStorageContext  -StorageAccountName $storageAccountName -SasToken $sasToken 
      $files= Get-ChildItem $sourcePath -Recurse -File
      foreach($file in $files)
      {
        $localfile =$file.FullName
        $path=$file.DirectoryName.ToLower().Replace($sourcePath,"")
        $fileName=$file.Name.ToLower()
        $remoteFile ="$($d)$($path)\$($fileName)"
        Set-AzStorageBlobContent  -Container  $container -File $localfile -Blob $remoteFile -Context $storageContext -Force
      }
      Write-Host "Done." -ForegroundColor Green
}catch{
    Write-Error  $_.Exception.Messag
}
