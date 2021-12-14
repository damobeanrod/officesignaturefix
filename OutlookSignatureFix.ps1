echo "         _              _                _                _                 _          _                      _       ";
echo "        _\ \           /\ \             /\ \             /\_\              /\ \       /\_\                   /\ \     ";
echo "       /\__ \         /  \ \           /  \ \           / / /  _          /  \ \     / / /         _         \_\ \    ";
echo "      / /_ \_\       / /\ \ \         / /\ \ \         / / /  /\_\       / /\ \ \    \ \ \__      /\_\       /\__ \   ";
echo "     / / /\/_/      / / /\ \ \       / / /\ \ \       / / /__/ / /      / / /\ \ \    \ \___\    / / /      / /_ \ \  ";
echo "    / / /          / / /  \ \_\     / / /  \ \_\     / /\_____/ /      / / /  \ \_\    \__  /   / / /      / / /\ \ \ ";
echo "   / / /          / / /   / / /    / / /   / / /    / /\_______/      / / /   / / /    / / /   / / /      / / /  \/_/ ";
echo "  / / / ____     / / /   / / /    / / /   / / /    / / /\ \ \        / / /   / / /    / / /   / / /      / / /        ";
echo " / /_/_/ ___/\  / / /___/ / /    / / /___/ / /    / / /  \ \ \      / / /___/ / /    / / /___/ / /      / / /         ";
echo "/_______/\__\/ / / /____\/ /    / / /____\/ /    / / /    \ \ \    / / /____\/ /    / / /____\/ /      /_/ /          ";
echo "\_______\/     \/_________/     \/_________/     \/_/      \_\_\   \/_________/     \/_________/       \_\/           ";
echo "        /\ \            /\ \             /\ \                                                                         ";
echo "       /  \ \          /  \ \           /  \ \                                                                        ";
echo "      / /\ \ \        / /\ \ \         / /\ \ \                                                                       ";
echo "     / / /\ \_\      / / /\ \ \       / / /\ \_\                                                                      ";
echo "    / /_/_ \/_/     / / /  \ \_\     / / /_/ / /                                                                      ";
echo "   / /____/\       / / /   / / /    / / /__\/ /                                                                       ";
echo "  / /\____\/      / / /   / / /    / / /_____/                                                                        ";
echo " / / /           / / /___/ / /    / / /\ \ \                                                                          ";
echo "/ / /           / / /____\/ /    / / /  \ \ \                                                                         ";
echo "\/_/     _      \/_________/     \/_/    \_\/               _              _                _                _        ";
echo "        /\ \       /\_\                   /\ \             _\ \           /\ \             /\ \             /\_\      ";
echo "       /  \ \     / / /         _         \_\ \           /\__ \         /  \ \           /  \ \           / / /  _   ";
echo "      / /\ \ \    \ \ \__      /\_\       /\__ \         / /_ \_\       / /\ \ \         / /\ \ \         / / /  /\_\ ";
echo "     / / /\ \ \    \ \___\    / / /      / /_ \ \       / / /\/_/      / / /\ \ \       / / /\ \ \       / / /__/ / / ";
echo "    / / /  \ \_\    \__  /   / / /      / / /\ \ \     / / /          / / /  \ \_\     / / /  \ \_\     / /\_____/ /  ";
echo "   / / /   / / /    / / /   / / /      / / /  \/_/    / / /          / / /   / / /    / / /   / / /    / /\_______/   ";
echo "  / / /   / / /    / / /   / / /      / / /          / / / ____     / / /   / / /    / / /   / / /    / / /\ \ \      ";
echo " / / /___/ / /    / / /___/ / /      / / /          / /_/_/ ___/\  / / /___/ / /    / / /___/ / /    / / /  \ \ \     ";
echo "/ / /____\/ /    / / /____\/ /      /_/ /          /_______/\__\/ / / /____\/ /    / / /____\/ /    / / /    \ \ \    ";
echo "\/_________/     \/_________/       \_\/           \_______\/     \/_________/     \/_________/     \/_/      \_\_\   ";
echo "                                                                                                                      ";




#Declare Locations where to clear the key value from.
$RegLocation = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\'
$RegLocation2 = 'HKCR:\CLSID\'
$KeyValue = '*0006F03A-0000-0000-C000-000000000046*'
$RegBackupLocation = 'C:\Regback'
New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR


#This will backup the keys to the regbackup location above
#Remove the colon from the variables so they can work with reg export
mkdir $RegBackupLocation
$rlbak = $RegLocation -Replace ':',''
reg export $rlbak $RegBackupLocation\RegLocation1.reg /y
$rlbak2 = $RegLocation2 -Replace ':',''
reg export $rlbak2 $RegBackupLocation\RegLocation2.reg /y


#The same routine twice against each location
Get-ChildItem $RegLocation -Rec -EA SilentlyContinue | ForEach-Object {
   $CurrentKey = (Get-ItemProperty -Path $_.PsPath)
   If ($CurrentKey -like "$KeyValue"){
     $CurrentKey|Remove-Item -Force -Verbose -Recurse -EA SilentlyContinue
   }
}


Get-ChildItem $RegLocation2 -Rec -EA SilentlyContinue | ForEach-Object {
   $CurrentKey = (Get-ItemProperty -Path $_.PsPath)
   If ($CurrentKey -like "$KeyValue"){
     $CurrentKey|Remove-Item -Force -Verbose -Recurse -EA SilentlyContinue
   }
}
